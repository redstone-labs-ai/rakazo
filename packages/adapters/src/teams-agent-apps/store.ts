import type { PrismaClient } from "@rakazo/db";
import { runAsSystem } from "@rakazo/db";

export const AGENT_TEAMS_APP_STATUSES = [
  "creating",
  "registered",
  "published",
  "visible",
  "removing",
  "removed",
  "failed",
] as const;
export type AgentTeamsAppStatus = (typeof AGENT_TEAMS_APP_STATUSES)[number];

/** One `agent_teams_apps` row. */
export interface AgentTeamsAppRecord {
  id: string;
  botId: string;
  status: string;
  entraAppId: string | null;
  entraObjectId: string | null;
  secretRef: string | null;
  azureBotName: string | null;
  catalogAppId: string | null;
  manifestVersion: string | null;
  publishedAt: Date | null;
  firstInstalledAt: Date | null;
  lastError: string | null;
  /** Operation lease: while in the future, a provision or removal holds the row. */
  lockedUntil: Date | null;
  createdAt: Date;
  updatedAt: Date;
}

/** One `agent_teams_chats` row. */
export interface AgentTeamsChatRecord {
  id: string;
  agentAppId: string;
  chatId: string;
  installId: string | null;
  installedAt: Date;
  removedAt: Date | null;
  /** The shared Bot's handoff was posted in this chat; a retried removal skips it. */
  handoffPostedAt: Date | null;
}

export type AgentTeamsAppPatch = Partial<
  Omit<AgentTeamsAppRecord, "id" | "botId" | "lockedUntil" | "createdAt" | "updatedAt">
>;

/**
 * Persistence for provisioning. There is deliberately no delete: a removed agent keeps its row
 * with status `removed`, and nothing (bot or group chat removal) cascades into it.
 */
export interface AgentTeamsAppStore {
  findByBotId(botId: string): Promise<AgentTeamsAppRecord | null>;
  findById(id: string): Promise<AgentTeamsAppRecord | null>;
  /** Creates the row in `creating`, or returns the existing one for this bot. */
  ensure(botId: string): Promise<AgentTeamsAppRecord>;
  update(id: string, patch: AgentTeamsAppPatch): Promise<AgentTeamsAppRecord>;
  /**
   * Compare-and-set: applies the patch only while the row's status is one of `statuses`.
   * Null, with nothing written, when the status had moved on (e.g. removal started).
   */
  updateIfStatus(
    id: string,
    statuses: readonly string[],
    patch: AgentTeamsAppPatch,
  ): Promise<AgentTeamsAppRecord | null>;
  /**
   * Takes the row's operation lease until `until` when it is free (unset, or expired at `now`).
   * False when another provision or removal holds it.
   */
  acquireLock(id: string, until: Date, now: Date): Promise<boolean>;
  /** Releases the lease taken with `acquireLock`. */
  releaseLock(id: string): Promise<void>;
  /** Chats the agent app is still installed in. */
  activeChats(agentAppId: string): Promise<AgentTeamsChatRecord[]>;
  chatCount(agentAppId: string): Promise<number>;
  /** Inserts the chat, or reactivates it with the new install id. */
  upsertChat(input: {
    agentAppId: string;
    chatId: string;
    installId: string | null;
    installedAt: Date;
  }): Promise<AgentTeamsChatRecord>;
  markChatRemoved(id: string, removedAt: Date): Promise<void>;
  markHandoffPosted(id: string, postedAt: Date): Promise<void>;
}

/** The Prisma store. The tables hold no tenant data, so every call runs as system. */
export function prismaAgentTeamsAppStore(prisma: PrismaClient): AgentTeamsAppStore {
  const system = <T>(reason: string, fn: () => Promise<T>) => runAsSystem(reason, fn);
  return {
    findByBotId: (botId) =>
      system("read an agent Teams app", () =>
        prisma.agentTeamsApp.findUnique({ where: { botId } }),
      ),
    findById: (id) =>
      system("read an agent Teams app", () => prisma.agentTeamsApp.findUnique({ where: { id } })),
    ensure: (botId) =>
      system("create an agent Teams app", () =>
        prisma.agentTeamsApp.upsert({
          where: { botId },
          create: { botId, status: "creating" },
          update: {},
        }),
      ),
    update: (id, patch) =>
      system("update an agent Teams app", () =>
        prisma.agentTeamsApp.update({ where: { id }, data: patch }),
      ),
    updateIfStatus: (id, statuses, patch) =>
      system("update an agent Teams app if its status allows", async () => {
        const { count } = await prisma.agentTeamsApp.updateMany({
          where: { id, status: { in: [...statuses] } },
          data: patch,
        });
        return count ? prisma.agentTeamsApp.findUnique({ where: { id } }) : null;
      }),
    acquireLock: (id, until, now) =>
      system("lock an agent Teams app", async () => {
        const { count } = await prisma.agentTeamsApp.updateMany({
          where: { id, OR: [{ lockedUntil: null }, { lockedUntil: { lt: now } }] },
          data: { lockedUntil: until },
        });
        return count === 1;
      }),
    releaseLock: async (id) => {
      await system("unlock an agent Teams app", () =>
        prisma.agentTeamsApp.updateMany({ where: { id }, data: { lockedUntil: null } }),
      );
    },
    activeChats: (agentAppId) =>
      system("list an agent Teams app's chats", () =>
        prisma.agentTeamsChat.findMany({
          where: { agentAppId, removedAt: null },
          orderBy: { installedAt: "asc" },
        }),
      ),
    chatCount: (agentAppId) =>
      system("count an agent Teams app's chats", () =>
        prisma.agentTeamsChat.count({ where: { agentAppId, removedAt: null } }),
      ),
    upsertChat: ({ agentAppId, chatId, installId, installedAt }) =>
      system("record an agent Teams app chat install", () =>
        prisma.agentTeamsChat.upsert({
          where: { agentAppId_chatId: { agentAppId, chatId } },
          create: { agentAppId, chatId, installId, installedAt },
          // A new install is a new membership: its handoff has not been posted yet.
          update: { installId, installedAt, removedAt: null, handoffPostedAt: null },
        }),
      ),
    markChatRemoved: async (id, removedAt) => {
      await system("mark an agent Teams app chat removed", () =>
        prisma.agentTeamsChat.update({ where: { id }, data: { removedAt } }),
      );
    },
    markHandoffPosted: async (id, handoffPostedAt) => {
      await system("record an agent Teams app handoff", () =>
        prisma.agentTeamsChat.update({ where: { id }, data: { handoffPostedAt } }),
      );
    },
  };
}

/** An in-memory store for tests and local runs. */
export function memoryAgentTeamsAppStore(now: () => Date = () => new Date()): AgentTeamsAppStore & {
  apps: Map<string, AgentTeamsAppRecord>;
  chats: Map<string, AgentTeamsChatRecord>;
} {
  const apps = new Map<string, AgentTeamsAppRecord>();
  const chats = new Map<string, AgentTeamsChatRecord>();
  let sequence = 0;
  const nextId = (prefix: string) => `${prefix}_${++sequence}`;
  const byBot = (botId: string) => [...apps.values()].find((app) => app.botId === botId) ?? null;
  const patchApp = (app: AgentTeamsAppRecord, patch: AgentTeamsAppPatch) => {
    const updated = { ...app, ...patch, updatedAt: now() };
    apps.set(app.id, updated);
    return { ...updated };
  };
  return {
    apps,
    chats,
    findByBotId: async (botId) => {
      const app = byBot(botId);
      return app ? { ...app } : null;
    },
    findById: async (id) => {
      const app = apps.get(id);
      return app ? { ...app } : null;
    },
    ensure: async (botId) => {
      const existing = byBot(botId);
      if (existing) return { ...existing };
      const at = now();
      const app: AgentTeamsAppRecord = {
        id: nextId("app"),
        botId,
        status: "creating",
        entraAppId: null,
        entraObjectId: null,
        secretRef: null,
        azureBotName: null,
        catalogAppId: null,
        manifestVersion: null,
        publishedAt: null,
        firstInstalledAt: null,
        lastError: null,
        lockedUntil: null,
        createdAt: at,
        updatedAt: at,
      };
      apps.set(app.id, app);
      return { ...app };
    },
    update: async (id, patch) => {
      const app = apps.get(id);
      if (!app) throw new Error(`No agent Teams app ${id}`);
      return patchApp(app, patch);
    },
    updateIfStatus: async (id, statuses, patch) => {
      const app = apps.get(id);
      if (!app || !statuses.includes(app.status)) return null;
      return patchApp(app, patch);
    },
    acquireLock: async (id, until, at) => {
      const app = apps.get(id);
      if (!app || (app.lockedUntil && app.lockedUntil >= at)) return false;
      apps.set(id, { ...app, lockedUntil: until });
      return true;
    },
    releaseLock: async (id) => {
      const app = apps.get(id);
      if (app) apps.set(id, { ...app, lockedUntil: null });
    },
    activeChats: async (agentAppId) =>
      [...chats.values()]
        .filter((chat) => chat.agentAppId === agentAppId && !chat.removedAt)
        .map((chat) => ({ ...chat })),
    chatCount: async (agentAppId) =>
      [...chats.values()].filter((chat) => chat.agentAppId === agentAppId && !chat.removedAt)
        .length,
    upsertChat: async ({ agentAppId, chatId, installId, installedAt }) => {
      const existing = [...chats.values()].find(
        (chat) => chat.agentAppId === agentAppId && chat.chatId === chatId,
      );
      const chat: AgentTeamsChatRecord = existing
        ? { ...existing, installId, installedAt, removedAt: null, handoffPostedAt: null }
        : {
            id: nextId("chat"),
            agentAppId,
            chatId,
            installId,
            installedAt,
            removedAt: null,
            handoffPostedAt: null,
          };
      chats.set(chat.id, chat);
      return { ...chat };
    },
    markChatRemoved: async (id, removedAt) => {
      const chat = chats.get(id);
      if (chat) chats.set(id, { ...chat, removedAt });
    },
    markHandoffPosted: async (id, handoffPostedAt) => {
      const chat = chats.get(id);
      if (chat) chats.set(id, { ...chat, handoffPostedAt });
    },
  };
}
