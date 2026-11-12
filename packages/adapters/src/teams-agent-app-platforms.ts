import type { PrismaClient } from "@rakazo/db";
import { runAsSystem } from "@rakazo/db";
import { getLogger } from "@rakazo/logging";
import type { MessagingPlatform } from "./chat-sdk-surface.js";
import type { TeamsAgentBot } from "./teams-messaging.js";
import { markChatRemoved, recordChatInstall } from "./teams-agent-apps/provision.js";
import { prismaAgentTeamsAppStore } from "./teams-agent-apps/store.js";
import { teamsAgentProvider, teamsPlatform } from "./teams-messaging.js";

/** One TEAMS_AGENT_APPS_ENABLED check, shared with the provisioner. */
export { teamsAgentAppsEnabled } from "./teams-agent-apps/env.js";

/**
 * Agent Teams apps (TEAMS_AGENT_APPS_ENABLED): every agent whose own Teams app the provisioner
 * created at runtime gets its own mounted Teams platform, loaded from `agent_teams_apps` with the
 * app password read from a secret store. Platforms are mounted and unmounted without a restart.
 */

/** Statuses whose app is in the Teams catalog and can receive activities. */
export const MOUNTABLE_AGENT_TEAMS_APP_STATUSES = ["published", "visible"] as const;

/** The parts of an `agent_teams_apps` row the router reads. */
export interface AgentTeamsAppRow {
  id: string;
  botId: string;
  status: string;
  entraAppId: string | null;
  secretRef: string | null;
}

/** Reads one app password by its secret name (Key Vault in production). */
export type AgentTeamsSecretReader = (secretRef: string) => Promise<string | undefined>;

/** The install bookkeeping the provisioner owns (`recordChatInstall`), injected. */
export interface AgentTeamsChatInstalls {
  recordChatInstall(agentAppId: string, chatId: string, installId: string | null): Promise<void>;
  markChatRemoved(agentAppId: string, chatId: string): Promise<void>;
}

/** What the router needs of the messaging surface (ChatSdkMessagingSurface). */
export interface TeamsAgentPlatformHost {
  mountedPlatform(provider: string): MessagingPlatform | undefined;
  mountPlatform(platform: MessagingPlatform): Promise<void>;
  unmountPlatform(provider: string): Promise<boolean>;
  platforms(): Array<{ provider: string }>;
}

export interface TeamsAgentAppRouterOptions {
  host: TeamsAgentPlatformHost;
  /** The one tenant every bot answers (TEAMS_APP_TENANT_ID). */
  tenantId: string;
  /** Env-configured agent bots (TEAMS_AGENT_BOTS); restored when a DB row goes away. */
  envBots: TeamsAgentBot[];
  loadApps: () => Promise<AgentTeamsAppRow[]>;
  readSecret: AgentTeamsSecretReader;
  installs: AgentTeamsChatInstalls;
  /** Refresh interval; 0 disables the timer. Default 30 seconds. */
  intervalMs?: number;
  /** Builds a platform; tests replace the real Teams adapter. */
  createPlatform?: (
    bot: TeamsAgentBot,
    tenantId: string,
    provider: string,
  ) => MessagingPlatform | undefined;
}

interface ManagedApp {
  rowId: string;
  appId: string;
  appPassword: string;
}

const DEFAULT_REFRESH_MS = 30_000;

function defaultCreatePlatform(bot: TeamsAgentBot, tenantId: string, provider: string) {
  return teamsPlatform(
    { teamsAppId: bot.appId, teamsAppPassword: bot.appPassword, teamsAppTenantId: tenantId },
    { provider, agentId: bot.agentId },
  );
}

export class TeamsAgentAppRouter {
  /** Platforms this router mounted, by provider. */
  private readonly managed = new Map<string, ManagedApp>();
  /** `<rowId>|<chatId>` already recorded as installed in this process. */
  private readonly seenChats = new Set<string>();
  private timer: ReturnType<typeof setInterval> | undefined;
  private running: Promise<void> = Promise.resolve();

  constructor(private readonly options: TeamsAgentAppRouterOptions) {}

  /** Providers this router has mounted, with the app id each one serves. */
  mounted(): Map<string, string> {
    return new Map([...this.managed].map(([provider, app]) => [provider, app.appId]));
  }

  start(): void {
    const interval = this.options.intervalMs ?? DEFAULT_REFRESH_MS;
    void this.refresh();
    if (interval > 0 && !this.timer) {
      this.timer = setInterval(() => void this.refresh(), interval);
      this.timer.unref?.();
    }
  }

  stop(): void {
    if (this.timer) clearInterval(this.timer);
    this.timer = undefined;
  }

  /** Reconciles mounted platforms with the table. Calls run one at a time. */
  refresh(): Promise<void> {
    const next = this.running.then(() =>
      this.reconcile().catch((error: unknown) => {
        getLogger().warn("agent Teams app refresh failed", {
          reason: error instanceof Error ? error.message : String(error),
        });
      }),
    );
    this.running = next;
    return next;
  }

  private async reconcile(): Promise<void> {
    const { host } = this.options;
    const mountable = new Set<string>(MOUNTABLE_AGENT_TEAMS_APP_STATUSES);
    const rows = (await this.options.loadApps()).filter(
      (row) => mountable.has(row.status) && row.entraAppId && row.secretRef,
    );
    const desired = new Map<string, ManagedApp & { botId: string }>();
    for (const row of rows) {
      const secretRef = row.secretRef;
      const entraAppId = row.entraAppId;
      if (!secretRef || !entraAppId) continue;
      let appPassword: string | undefined;
      try {
        appPassword = (await this.options.readSecret(secretRef))?.trim();
      } catch {
        appPassword = undefined;
      }
      if (!appPassword) {
        getLogger().warn("agent Teams app secret unavailable; not mounted", { botId: row.botId });
        continue;
      }
      desired.set(teamsAgentProvider(row.botId), {
        rowId: row.id,
        botId: row.botId,
        appId: entraAppId.toLowerCase(),
        appPassword,
      });
    }

    // Rows gone (removed, failed, deleted): unmount, then give back an env bot for that agent.
    for (const provider of [...this.managed.keys()]) {
      if (desired.has(provider)) continue;
      this.managed.delete(provider);
      await host.unmountPlatform(provider);
      const envBot = this.options.envBots.find(
        (bot) => teamsAgentProvider(bot.agentId) === provider,
      );
      if (envBot && !this.appIdTaken(envBot.appId, provider)) {
        const platform = this.build(envBot, provider);
        if (platform) await host.mountPlatform(platform);
      }
    }

    for (const [provider, app] of desired) {
      const current = this.managed.get(provider);
      if (current && current.appId === app.appId && current.appPassword === app.appPassword) {
        continue;
      }
      // Never two platforms on one app id: both would answer its activities.
      if (this.appIdTaken(app.appId, provider)) {
        getLogger().warn("agent Teams app reuses a mounted bot's app id; not mounted", {
          botId: app.botId,
        });
        continue;
      }
      const platform = this.build(
        { agentId: app.botId, appId: app.appId, appPassword: app.appPassword },
        provider,
        app.rowId,
      );
      if (!platform) continue;
      // Replaces an env-configured bot for the same agent: the DB row wins.
      await host.mountPlatform(platform);
      this.managed.set(provider, {
        rowId: app.rowId,
        appId: app.appId,
        appPassword: app.appPassword,
      });
    }
  }

  /** True when a platform other than `provider` already serves the app id. */
  private appIdTaken(appId: string, provider: string): boolean {
    for (const { provider: other } of this.options.host.platforms()) {
      if (other === provider) continue;
      if (this.options.host.mountedPlatform(other)?.routeKey === appId) return true;
    }
    return false;
  }

  private build(bot: TeamsAgentBot, provider: string, rowId?: string) {
    const create = this.options.createPlatform ?? defaultCreatePlatform;
    const platform = create(bot, this.options.tenantId, provider);
    if (!platform || !rowId) return platform;
    return { ...platform, verifiedWebhook: (payload: unknown) => this.observe(rowId, payload) };
  }

  /**
   * Install tracking from a verified activity to an agent app: the first activity in a chat, or
   * an installationUpdate add, records the install; an installationUpdate remove ends it.
   */
  private async observe(rowId: string, payload: unknown): Promise<void> {
    const activity = asRecord(payload);
    const chatId = teamsChatIdOfActivity(activity);
    if (!activity || !chatId) return;
    const key = `${rowId}|${chatId}`;
    const installId = typeof activity.id === "string" && activity.id ? activity.id : null;
    if (activity.type === "installationUpdate") {
      const action = typeof activity.action === "string" ? activity.action : "";
      if (action.startsWith("remove")) {
        this.seenChats.delete(key);
        await this.options.installs.markChatRemoved(rowId, chatId);
        return;
      }
      if (!action.startsWith("add")) return;
      await this.options.installs.recordChatInstall(rowId, chatId, installId);
      this.seenChats.add(key);
      return;
    }
    if (this.seenChats.has(key)) return;
    await this.options.installs.recordChatInstall(rowId, chatId, installId);
    this.seenChats.add(key);
  }
}

/** The Teams chat (or channel) id an activity belongs to, without a channel post's message id. */
export function teamsChatIdOfActivity(
  activity: Record<string, unknown> | undefined,
): string | undefined {
  const id = asRecord(activity?.conversation)?.id;
  if (typeof id !== "string" || !id) return undefined;
  return id.replace(/;messageid=\d+$/, "");
}

let active: TeamsAgentAppRouter | undefined;

/** Registers the process's router so `refreshAgentTeamsPlatforms` reaches it. */
export function setActiveTeamsAgentAppRouter(router: TeamsAgentAppRouter | undefined): void {
  active?.stop();
  active = router;
}

/**
 * Re-reads `agent_teams_apps` and mounts or unmounts agent Teams platforms now. The provisioner
 * calls this after an app is published or removed. A no-op while the flag is off.
 */
export async function refreshAgentTeamsPlatforms(): Promise<void> {
  await active?.refresh();
}

/** Loads mountable rows from `agent_teams_apps`. */
export function prismaAgentTeamsApps(prisma: PrismaClient): () => Promise<AgentTeamsAppRow[]> {
  return () =>
    runAsSystem("mount agent Teams apps", () =>
      prisma.agentTeamsApp.findMany({
        where: { status: { in: [...MOUNTABLE_AGENT_TEAMS_APP_STATUSES] } },
        select: { id: true, botId: true, status: true, entraAppId: true, secretRef: true },
      }),
    );
}

/**
 * Install bookkeeping backed by the provisioner's `recordChatInstall` / `markChatRemoved`: the
 * chat row is upserted and the app's first install moves "published" to "visible".
 */
export function prismaAgentTeamsChatInstalls(prisma: PrismaClient): AgentTeamsChatInstalls {
  const store = prismaAgentTeamsAppStore(prisma);
  return {
    async recordChatInstall(agentAppId, chatId, installId) {
      await recordChatInstall(agentAppId, chatId, installId, { store });
    },
    async markChatRemoved(agentAppId, chatId) {
      await markChatRemoved(agentAppId, chatId, { store });
    },
  };
}

function asRecord(value: unknown): Record<string, unknown> | undefined {
  return typeof value === "object" && value !== null
    ? (value as Record<string, unknown>)
    : undefined;
}
