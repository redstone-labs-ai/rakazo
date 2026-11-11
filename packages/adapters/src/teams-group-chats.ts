import { getLogger } from "@rakazo/logging";
import {
  SHARED_TEAMS_BOT_APP_ID,
  teamsAgentBotsFromEnv,
  teamsAgentProvider,
} from "./teams-messaging.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const GRAPH_SCOPE = "https://graph.microsoft.com/.default";
const BOT_CONNECTOR_SCOPE = "https://api.botframework.com/.default";
const DEFAULT_SERVICE_URL = "https://smba.trafficmanager.net/teams/";
const DEFAULT_RSC_PERMISSIONS = ["ChatMessage.Read.Chat"];
/** Renew a token this long before it expires. */
const TOKEN_EARLY_RENEW_MS = 60_000;
/** Teams accepts up to 250 members in a group chat; the bot adds itself. */
export const TEAMS_GROUP_CHAT_MAX_MEMBERS = 249;
/** The most other agent bots one create call adds. */
export const TEAMS_GROUP_CHAT_MAX_BOTS = 10;

type FetchLike = (input: string, init?: RequestInit) => Promise<Response>;

/** One bot that creates group chats: its app credentials and its Teams app catalog id. */
export interface TeamsGroupChatBot {
  appId: string;
  appPassword: string;
  /** The Teams app id in the org app catalog, which `POST /chats/{id}/installedApps` binds. */
  catalogAppId?: string;
}

export interface TeamsGroupChatsConfig {
  tenantId: string;
  /** The shared bot. Agents without their own bot create chats with it. */
  shared?: TeamsGroupChatBot;
  /** Agent id to that agent's own bot. */
  agents: ReadonlyMap<string, TeamsGroupChatBot>;
  /** Send consentedPermissionSet when the app is installed (TEAMS_RSC_CONSENT_ENABLED). */
  rscConsentEnabled: boolean;
  /** The resource-specific permissions to consent to. They must match the app manifest. */
  rscPermissions: readonly string[];
  /** Bot Connector service URL for the tenant (TEAMS_SERVICE_URL). */
  serviceUrl: string;
  fetch?: FetchLike;
  now?: () => number;
}

export interface CreateTeamsGroupChatInput {
  agentId: string;
  /** Entra object ids or user principal names. */
  members: readonly string[];
  topic?: string;
  firstMessage: string;
  /**
   * Other agents whose own bots join the chat. Teams adds a bot to a chat only by installing its
   * app, so each of these bots installs its own app with its own credentials.
   */
  bots?: readonly string[];
}

/** The result for one other agent bot that was asked to join the chat. */
export interface TeamsGroupChatBotJoin {
  agentId: string;
  installed: boolean;
  error?: string;
}

export interface CreatedTeamsGroupChat {
  chatId: string;
  /** The Rakazo thread id of the chat, in the Teams adapter format. */
  threadId: string;
  /** The Rakazo provider whose bot owns the chat: `teams` or `teams-<agent id>`. */
  provider: string;
  appId: string;
  installed: boolean;
  messageId?: string;
  /** One entry for each agent in the input `bots`, in order. */
  bots: TeamsGroupChatBotJoin[];
}

/** A Graph or Bot Connector failure. `code` is the service error code, when it sent one. */
export class TeamsGroupChatError extends Error {
  constructor(
    message: string,
    readonly step: "token" | "create" | "install" | "message",
    readonly status: number,
    readonly code?: string,
    /** The chat, when Teams created it before the failure. */
    readonly chat?: { chatId: string; threadId: string; provider: string },
  ) {
    super(message);
    this.name = "TeamsGroupChatError";
  }
}

interface CachedToken {
  token: string;
  expiresAt: number;
}

/**
 * Creates Teams group chats as an agent's own bot. Each (app id, scope) pair has its own token
 * entry, so agents never read or replace each other's tokens. A request in flight is shared only
 * by calls for the same app and scope.
 */
export class TeamsGroupChats {
  private readonly tokens = new Map<string, Promise<CachedToken>>();
  private readonly fetch: FetchLike;
  private readonly now: () => number;

  constructor(private readonly config: TeamsGroupChatsConfig) {
    this.fetch = config.fetch ?? ((input, init) => fetch(input, init));
    this.now = config.now ?? Date.now;
  }

  /** The bot an agent creates chats with: its own, else the shared bot. */
  botFor(agentId: string): { bot: TeamsGroupChatBot; provider: string } | undefined {
    const own = this.config.agents.get(agentId);
    if (own) return { bot: own, provider: teamsAgentProvider(agentId) };
    return this.config.shared ? { bot: this.config.shared, provider: "teams" } : undefined;
  }

  async create(input: CreateTeamsGroupChatInput): Promise<CreatedTeamsGroupChat> {
    const selected = this.botFor(input.agentId);
    if (!selected) {
      throw new TeamsGroupChatError("No Teams bot is set up for this agent.", "token", 0);
    }
    const { bot, provider } = selected;
    const graphToken = await this.token(bot, GRAPH_SCOPE);
    const chatId = await this.createChat(graphToken, input);
    const threadId = teamsGroupThreadId(chatId, this.config.serviceUrl);
    let installed = false;
    let messageId: string | undefined;
    let joins: TeamsGroupChatBotJoin[] = [];
    try {
      if (bot.catalogAppId) {
        await this.installApp(graphToken, chatId, bot.catalogAppId);
        installed = true;
      }
      joins = await this.addBots(chatId, input.agentId, bot.appId, input.bots ?? []);
      const connectorToken = await this.token(bot, BOT_CONNECTOR_SCOPE);
      messageId = await this.postMessage(connectorToken, chatId, input.firstMessage);
    } catch (error) {
      // The chat exists now. Give its id to the caller, so it can register the chat.
      if (!(error instanceof TeamsGroupChatError)) throw error;
      throw new TeamsGroupChatError(error.message, error.step, error.status, error.code, {
        chatId,
        threadId,
        provider,
      });
    }
    return {
      chatId,
      threadId,
      provider,
      appId: bot.appId,
      installed,
      ...(messageId ? { messageId } : {}),
      bots: joins,
    };
  }

  /**
   * Adds other agents' bots. Each bot installs its own app, so no bot needs a permission to
   * install another app. A failure for one bot does not stop the others or the chat.
   */
  private async addBots(
    chatId: string,
    creatorAgentId: string,
    creatorAppId: string,
    agentIds: readonly string[],
  ): Promise<TeamsGroupChatBotJoin[]> {
    const joins: TeamsGroupChatBotJoin[] = [];
    const seenApps = new Set([creatorAppId]);
    for (const agentId of [...new Set(agentIds)]) {
      const own = this.config.agents.get(agentId);
      if (agentId === creatorAgentId || (own && seenApps.has(own.appId))) {
        joins.push({ agentId, installed: false, error: "This bot is already in the chat." });
        continue;
      }
      if (!own?.catalogAppId) {
        joins.push({
          agentId,
          installed: false,
          error: "This agent has no own Teams bot with a catalog app id.",
        });
        continue;
      }
      seenApps.add(own.appId);
      try {
        const token = await this.token(own, GRAPH_SCOPE);
        await this.installApp(token, chatId, own.catalogAppId);
        joins.push({ agentId, installed: true });
      } catch (error) {
        if (!(error instanceof TeamsGroupChatError)) throw error;
        joins.push({ agentId, installed: false, error: error.message });
      }
    }
    return joins;
  }

  private token(bot: TeamsGroupChatBot, scope: string): Promise<string> {
    const key = `${bot.appId}|${scope}`;
    const cached = this.tokens.get(key);
    const pending = cached
      ? cached.then((entry) =>
          entry.expiresAt - TOKEN_EARLY_RENEW_MS > this.now() ? entry : this.renew(key, bot, scope),
        )
      : this.renew(key, bot, scope);
    return pending.then((entry) => entry.token);
  }

  private renew(key: string, bot: TeamsGroupChatBot, scope: string): Promise<CachedToken> {
    const request = this.requestToken(bot, scope);
    this.tokens.set(key, request);
    // A failed request must not stay in the cache.
    request.catch(() => {
      if (this.tokens.get(key) === request) this.tokens.delete(key);
    });
    return request;
  }

  private async requestToken(bot: TeamsGroupChatBot, scope: string): Promise<CachedToken> {
    const response = await this.fetch(
      `https://login.microsoftonline.com/${encodeURIComponent(this.config.tenantId)}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "content-type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          grant_type: "client_credentials",
          client_id: bot.appId,
          client_secret: bot.appPassword,
          scope,
        }).toString(),
      },
    );
    const body = await readJson(response);
    const token = typeof body?.access_token === "string" ? body.access_token : "";
    if (!response.ok || !token) {
      const code = typeof body?.error === "string" ? body.error : undefined;
      throw new TeamsGroupChatError(
        `The bot could not get a token${code ? ` (${code})` : ""}.`,
        "token",
        response.status,
        code,
      );
    }
    const expiresIn = typeof body?.expires_in === "number" ? body.expires_in : 3600;
    return { token, expiresAt: this.now() + expiresIn * 1000 };
  }

  private async createChat(token: string, input: CreateTeamsGroupChatInput): Promise<string> {
    const body = {
      chatType: "group",
      ...(input.topic ? { topic: input.topic } : {}),
      members: input.members.map((member) => ({
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        roles: ["owner"],
        "user@odata.bind": `${GRAPH_BASE}/users('${encodeURIComponent(member)}')`,
      })),
    };
    const result = await this.call("create", `${GRAPH_BASE}/chats`, token, body);
    const id = typeof result?.id === "string" ? result.id : "";
    if (!id) throw new TeamsGroupChatError("Graph returned no chat id.", "create", 201);
    return id;
  }

  private async installApp(token: string, chatId: string, catalogAppId: string): Promise<void> {
    const body = {
      "teamsApp@odata.bind": `${GRAPH_BASE}/appCatalogs/teamsApps/${encodeURIComponent(catalogAppId)}`,
      ...(this.config.rscConsentEnabled
        ? {
            consentedPermissionSet: {
              resourceSpecificPermissions: this.config.rscPermissions.map((permissionValue) => ({
                permissionValue,
                permissionType: "Application",
              })),
            },
          }
        : {}),
    };
    await this.call(
      "install",
      `${GRAPH_BASE}/chats/${encodeURIComponent(chatId)}/installedApps`,
      token,
      body,
    );
  }

  private async postMessage(token: string, chatId: string, text: string) {
    const base = this.config.serviceUrl.endsWith("/")
      ? this.config.serviceUrl
      : `${this.config.serviceUrl}/`;
    const result = await this.call(
      "message",
      `${base}v3/conversations/${encodeURIComponent(chatId)}/activities`,
      token,
      { type: "message", text },
    );
    return typeof result?.id === "string" ? result.id : undefined;
  }

  /** One request. A failure is thrown at once: no step is retried. */
  private async call(
    step: "create" | "install" | "message",
    url: string,
    token: string,
    body: unknown,
  ): Promise<Record<string, unknown> | undefined> {
    const response = await this.fetch(url, {
      method: "POST",
      headers: { authorization: `Bearer ${token}`, "content-type": "application/json" },
      body: JSON.stringify(body),
    });
    const json = await readJson(response);
    if (response.ok) return json;
    const error = asRecord(json?.error);
    const code = typeof error?.code === "string" ? error.code : undefined;
    const detail = typeof error?.message === "string" ? error.message : response.statusText;
    throw new TeamsGroupChatError(
      `Teams ${STEP_LABELS[step]} failed (${response.status}${code ? ` ${code}` : ""}): ${detail}`,
      step,
      response.status,
      code,
    );
  }
}

const STEP_LABELS = {
  create: "chat create",
  install: "app install",
  message: "first message",
} as const;

/** A group chat's Rakazo thread id, as the Teams adapter encodes it. */
export function teamsGroupThreadId(chatId: string, serviceUrl: string): string {
  const encode = (value: string) => Buffer.from(value).toString("base64url");
  return `teams:${encode(chatId)}:${encode(serviceUrl)}`;
}

function asRecord(value: unknown): Record<string, unknown> | undefined {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined;
}

async function readJson(response: Response): Promise<Record<string, unknown> | undefined> {
  const text = await response.text().catch(() => "");
  if (!text) return undefined;
  try {
    return asRecord(JSON.parse(text));
  } catch {
    return undefined;
  }
}

/**
 * Group chat creation from env. Off unless RAKAZO_TEAMS_GROUP_CHATS is "true" and the tenant is
 * set. TEAMS_AGENT_BOTS entries may name their Teams app with `catalogAppId`; the shared bot uses
 * TEAMS_APP_CATALOG_ID.
 */
export function teamsGroupChatsFromEnv(
  env: Record<string, string | undefined>,
  options: { fetch?: FetchLike } = {},
): TeamsGroupChats | undefined {
  if (env.RAKAZO_TEAMS_GROUP_CHATS !== "true") return undefined;
  const tenantId = env.TEAMS_APP_TENANT_ID?.trim();
  if (!tenantId) {
    getLogger().warn("RAKAZO_TEAMS_GROUP_CHATS is on but TEAMS_APP_TENANT_ID is empty");
    return undefined;
  }
  const catalogIds = agentCatalogIds(env.TEAMS_AGENT_BOTS);
  const agents = new Map<string, TeamsGroupChatBot>();
  for (const bot of teamsAgentBotsFromEnv(env)) {
    const catalogAppId = catalogIds.get(bot.agentId);
    agents.set(bot.agentId, {
      appId: bot.appId,
      appPassword: bot.appPassword,
      ...(catalogAppId ? { catalogAppId } : {}),
    });
  }
  const sharedPassword = env.TEAMS_APP_PASSWORD?.trim();
  const sharedCatalog = env.TEAMS_APP_CATALOG_ID?.trim();
  const shared: TeamsGroupChatBot | undefined = sharedPassword
    ? {
        appId: env.TEAMS_APP_ID?.trim().toLowerCase() || SHARED_TEAMS_BOT_APP_ID,
        appPassword: sharedPassword,
        ...(sharedCatalog ? { catalogAppId: sharedCatalog } : {}),
      }
    : undefined;
  const permissions = (env.TEAMS_RSC_PERMISSIONS ?? "")
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);
  return new TeamsGroupChats({
    tenantId,
    ...(shared ? { shared } : {}),
    agents,
    rscConsentEnabled: env.TEAMS_RSC_CONSENT_ENABLED === "true",
    rscPermissions: permissions.length ? permissions : DEFAULT_RSC_PERMISSIONS,
    serviceUrl: env.TEAMS_SERVICE_URL?.trim() || DEFAULT_SERVICE_URL,
    ...(options.fetch ? { fetch: options.fetch } : {}),
  });
}

/** `catalogAppId` per agent from TEAMS_AGENT_BOTS. Bad JSON gives none. */
function agentCatalogIds(raw: string | undefined): Map<string, string> {
  const ids = new Map<string, string>();
  if (!raw?.trim()) return ids;
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return ids;
  }
  for (const [agentId, value] of Object.entries(asRecord(parsed) ?? {})) {
    const id = asRecord(value)?.catalogAppId;
    if (typeof id === "string" && id.trim()) ids.set(agentId, id.trim());
  }
  return ids;
}
