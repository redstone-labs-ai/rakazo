import type {
  AgentTeamsAppStatus,
  AgentTeamsAppStore,
  AgentTeamsVisibilityStep,
} from "@rakazo/adapters/teams-agent-apps";
import { AGENT_TEAMS_APP_STATUSES } from "@rakazo/adapters/teams-agent-apps";
import type { Actor } from "@rakazo/contracts";
import type { PrismaClient } from "@rakazo/db";
import { runAsTenant } from "@rakazo/db";
import type { Hono } from "hono";

export type { AgentTeamsAppStatus };

/** UI step, from the provisioner's `visibilityStatus`. */
export type AgentTeamsAppStep = AgentTeamsVisibilityStep;

/** One `agent_teams_apps` row as the routes read it, with every chat it was installed in. */
export interface AgentTeamsAppRecord {
  status: AgentTeamsAppStatus;
  publishedAt: Date | null;
  firstInstalledAt: Date | null;
  lastError: string | null;
  chats: { chatId: string; installedAt: Date; removedAt: Date | null }[];
}

export interface AgentTeamsAppVisibility {
  step: AgentTeamsAppStep;
  label: string;
  estimated: boolean;
}

export type AgentTeamsAppJob = "provision" | "retry" | "deprovision";

export interface TeamsAgentAppsDeps {
  /** TEAMS_AGENT_APPS_ENABLED; off hides the feature and refuses every change. */
  enabled: boolean;
  sessionActor: (request: Request) => Promise<Actor | null>;
  /** Whether the actor owns the bot in their current space. */
  ownsBot: (actor: Actor, botId: string) => Promise<boolean>;
  /** Reads `agent_teams_apps` with its chats; null when the bot has none. */
  findApp: (botId: string) => Promise<AgentTeamsAppRecord | null>;
  /** `visibilityStatus(app, now, chatCount)` from the provisioner. */
  visibilityStatus: (
    app: AgentTeamsAppRecord,
    now: Date,
    chatCount: number,
  ) => AgentTeamsAppVisibility;
  /**
   * Starts `provisionAgentTeamsApp` / `deprovisionAgentTeamsApp` in the background (job queue
   * in production). Resolves once the work is queued, not when it finishes.
   */
  start: (job: AgentTeamsAppJob, botId: string) => Promise<void>;
  now?: () => Date;
}

export type TeamsAppResponse = {
  enabled: boolean;
  app: null | {
    status: AgentTeamsAppStatus;
    step: AgentTeamsAppStep;
    label: string;
    estimated: boolean;
    chatCount: number;
    chats: { chatId: string; installedAt: string }[];
    /** The redacted provider error, for a collapsed "Details" disclosure only. */
    lastError: string | null;
    /** What to tell the user about `lastError`; null when there is none. */
    userMessage: string | null;
    userMessageKind: TeamsAppErrorKind | null;
    publishedAt: string | null;
    firstInstalledAt: string | null;
  };
};

const PATH = "/api/bots/:botId/teams-app";

/** Why setup failed, in terms a bot owner can act on. The web card translates by kind. */
export type TeamsAppErrorKind = "replicating" | "catalog_auth" | "forbidden" | "unknown";

const USER_MESSAGES: Record<TeamsAppErrorKind, string> = {
  replicating:
    "Microsoft is still setting up this agent's app. This usually takes under a minute — Retry.",
  catalog_auth: "Couldn't publish to Teams. A Teams admin needs to sign in again.",
  forbidden: "Bot doesn't have permission to do this step. Ask an admin.",
  unknown: "Something went wrong setting up Teams. Retry, or ask an admin.",
};

/** Classifies a stored (already redacted) Graph/Azure error. Never echoes it. */
export function teamsAppErrorKind(lastError: string): TeamsAppErrorKind {
  if (
    /Request_ResourceNotFound|does not exist or one of its queried reference-property/i.test(
      lastError,
    )
  ) {
    return "replicating";
  }
  if (
    /invalid_grant|interaction_required|consent_required|AADSTS\d+|refresh token|catalog token/i.test(
      lastError,
    ) ||
    /Teams catalog app failed \((401|403)\b/i.test(lastError)
  ) {
    return "catalog_auth";
  }
  if (
    /\(403\b|Authorization_RequestDenied|AuthorizationFailed|Forbidden|app role assignment refused/i.test(
      lastError,
    )
  ) {
    return "forbidden";
  }
  return "unknown";
}

/** A short plain-English message for a stored error, or null when there is none. */
export function teamsAppUserMessage(
  lastError: string | null,
): { kind: TeamsAppErrorKind; message: string } | null {
  if (!lastError) return null;
  const kind = teamsAppErrorKind(lastError);
  return { kind, message: USER_MESSAGES[kind] };
}

export function prismaOwnsBot(prisma: PrismaClient) {
  return async (actor: Actor, botId: string) =>
    Boolean(
      await runAsTenant(actor, () =>
        prisma.bot.findFirst({
          where: { id: botId, spaceId: actor.spaceId, userId: actor.userId },
          select: { id: true },
        }),
      ),
    );
}

function isAgentTeamsAppStatus(value: string): value is AgentTeamsAppStatus {
  return (AGENT_TEAMS_APP_STATUSES as readonly string[]).includes(value);
}

/**
 * `findApp` over the provisioner's store. The routes list only active chats, so removed ones
 * are not loaded. An unknown status reads as "failed".
 */
export function storeFindApp(store: AgentTeamsAppStore): TeamsAgentAppsDeps["findApp"] {
  return async (botId) => {
    const row = await store.findByBotId(botId);
    if (!row) return null;
    const chats = await store.activeChats(row.id);
    return {
      status: isAgentTeamsAppStatus(row.status) ? row.status : "failed",
      publishedAt: row.publishedAt,
      firstInstalledAt: row.firstInstalledAt,
      lastError: row.lastError,
      chats: chats.map(({ chatId, installedAt, removedAt }) => ({
        chatId,
        installedAt,
        removedAt,
      })),
    };
  };
}

export function teamsAppView(
  app: AgentTeamsAppRecord | null,
  visibilityStatus: TeamsAgentAppsDeps["visibilityStatus"],
  now: Date,
): TeamsAppResponse["app"] {
  if (!app || app.status === "removed") return null;
  const chats = app.chats
    .filter((chat) => chat.removedAt === null)
    .map((chat) => ({ chatId: chat.chatId, installedAt: chat.installedAt.toISOString() }));
  const { step, label, estimated } = visibilityStatus(app, now, chats.length);
  const user = teamsAppUserMessage(app.lastError);
  return {
    status: app.status,
    step,
    label,
    estimated,
    chatCount: chats.length,
    chats,
    lastError: app.lastError,
    userMessage: user?.message ?? null,
    userMessageKind: user?.kind ?? null,
    publishedAt: app.publishedAt?.toISOString() ?? null,
    firstInstalledAt: app.firstInstalledAt?.toISOString() ?? null,
  };
}

/** `GET|POST|DELETE /api/bots/:botId/teams-app` and `POST …/retry`, for the bot's owner only. */
export function mountTeamsAgentAppRoutes(app: Hono, deps: TeamsAgentAppsDeps): void {
  const headers = { "cache-control": "no-store" };
  const now = deps.now ?? (() => new Date());

  // Unauthenticated → 401; another user's or unknown bot → 404 so ids are not probeable.
  async function owner(request: Request, botId: string) {
    const actor = await deps.sessionActor(request);
    if (!actor) return 401 as const;
    if (!(await deps.ownsBot(actor, botId))) return 404 as const;
    return actor;
  }

  app.get(PATH, async (c) => {
    const actor = await owner(c.req.raw, c.req.param("botId"));
    if (actor === 401) return c.json({ error: "Unauthorized" }, 401, headers);
    if (actor === 404) return c.json({ error: "Not found" }, 404, headers);
    if (!deps.enabled) return c.json({ enabled: false, app: null } satisfies TeamsAppResponse);
    const found = await runAsTenant(actor, () => deps.findApp(c.req.param("botId")));
    return c.json(
      {
        enabled: true,
        app: teamsAppView(found, deps.visibilityStatus, now()),
      } satisfies TeamsAppResponse,
      200,
      headers,
    );
  });

  async function mutate(
    request: Request,
    botId: string,
    job: AgentTeamsAppJob,
    allowed: (app: AgentTeamsAppRecord | null) => boolean,
  ) {
    const actor = await owner(request, botId);
    if (actor === 401) return { status: 401 as const, body: { error: "Unauthorized" } };
    if (actor === 404 || !deps.enabled)
      return { status: 404 as const, body: { error: "Not found" } };
    const found = await runAsTenant(actor, () => deps.findApp(botId));
    if (!allowed(found)) return { status: 409 as const, body: { error: "Conflict" } };
    await deps.start(job, botId);
    return { status: 202 as const, body: { accepted: true } };
  }

  const absent = (app: AgentTeamsAppRecord | null) => !app || app.status === "removed";

  app.post(PATH, async (c) => {
    const result = await mutate(c.req.raw, c.req.param("botId"), "provision", absent);
    return c.json(result.body, result.status, headers);
  });
  app.post(`${PATH}/retry`, async (c) => {
    const result = await mutate(
      c.req.raw,
      c.req.param("botId"),
      "retry",
      (found) => found?.status === "failed",
    );
    return c.json(result.body, result.status, headers);
  });
  app.delete(PATH, async (c) => {
    const result = await mutate(
      c.req.raw,
      c.req.param("botId"),
      "deprovision",
      (found) => !absent(found) && found?.status !== "removing",
    );
    return c.json(result.body, result.status, headers);
  });
}
