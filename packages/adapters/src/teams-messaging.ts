import type { TeamsAdapter, TeamsThreadId } from "@chat-adapter/teams";
import { getLogger } from "@rakazo/logging";
import type {
  Adapter,
  AdapterPostableMessage,
  ChatInstance,
  EmojiValue,
  FetchOptions,
  FormattedContent,
  Logger,
  WebhookOptions,
} from "chat";
import { stringifyMarkdown } from "chat";
import type { MessagingPlatform } from "./chat-sdk-surface.js";
import { boundInboundMarkdown } from "./chat-sdk-surface.js";

export interface TeamsCredentials {
  appId: string;
  appPassword: string;
  /** The one Microsoft Entra tenant the bot answers; every other tenant is refused. */
  tenantId: string;
}

/**
 * Microsoft Teams, mounted when the app id, password and tenant are all set.
 * Groups and channels use personal group linking, like iMessage and Telegram.
 * Teams names no roster and only delivers group messages that @mention the bot
 * unless the app manifest grants the ChatMessage.Read.Chat and
 * ChannelMessage.Read.Group resource-specific permissions.
 */
export function teamsPlatform(credentials: {
  teamsAppId?: string | undefined;
  teamsAppPassword?: string | undefined;
  teamsAppTenantId?: string | undefined;
}): MessagingPlatform | undefined {
  const { teamsAppId, teamsAppPassword, teamsAppTenantId } = credentials;
  if (!teamsAppId || !teamsAppPassword || !teamsAppTenantId) return undefined;
  return {
    provider: "teams",
    capabilities: { direct: true, groups: true, receipt: "typing", files: true },
    adapter: new LazyTeamsAdapter({
      appId: teamsAppId,
      appPassword: teamsAppPassword,
      tenantId: teamsAppTenantId,
    }),
    channelName: teamsChannelName,
    normalizeThreadId: teamsChannelThreadId,
    isPermanentSendError: isTeamsUnreachable,
    directoryUser: teamsDirectoryUser,
  };
}

/**
 * The sender's Microsoft Entra account. The adapter only passes activities from the configured
 * tenant whose Bot Framework token it verified, so the ids are Microsoft's.
 */
export function teamsDirectoryUser(activity: unknown) {
  const objectId = asRecord(asRecord(activity)?.from)?.aadObjectId;
  const tenantId = teamsActivityTenant(activity);
  return typeof objectId === "string" && objectId && tenantId
    ? { tenantId, objectId: objectId.toLowerCase() }
    : undefined;
}

/** The tenant a Bot Framework activity came from, lowercased. */
export function teamsActivityTenant(activity: unknown): string | undefined {
  const record = asRecord(activity);
  const tenant =
    asRecord(record?.conversation)?.tenantId ?? asRecord(asRecord(record?.channelData)?.tenant)?.id;
  return typeof tenant === "string" && tenant ? tenant.toLowerCase() : undefined;
}

/** Group chat topic or channel name; Teams leaves both off most messages. */
function teamsChannelName(raw: unknown): string | null {
  const activity = asRecord(raw);
  const name =
    asRecord(activity?.conversation)?.name ??
    asRecord(asRecord(activity?.channelData)?.channel)?.name;
  return typeof name === "string" && name ? name : null;
}

/**
 * Every post in a Teams channel is its own thread (`;messageid=<root>`), but
 * the channel is the conversation people approve their bots for, so threads
 * collapse onto it and replies post to the channel.
 */
function teamsChannelThreadId(threadId: string, raw: unknown): string {
  const conversation = asRecord(asRecord(raw)?.conversation);
  const id = conversation?.id;
  if (typeof id !== "string" || !/;messageid=\d+$/.test(id)) return threadId;
  const [prefix, , serviceUrl, type] = threadId.split(":");
  if (prefix !== "teams" || !serviceUrl) return threadId;
  const channel = Buffer.from(id.replace(/;messageid=\d+$/, "")).toString("base64url");
  return [prefix, channel, serviceUrl, ...(type ? [type] : [])].join(":");
}

/**
 * Teams answers 403 once the bot is removed from the conversation or the app
 * is blocked, 401 when its credentials are wrong, and 404 for a conversation
 * that no longer exists; none of these clears up on retry.
 */
function isTeamsUnreachable(error: unknown): boolean {
  if (!(error instanceof Error)) return false;
  if (error.name === "PermissionError" || error.name === "AuthenticationError") return true;
  return error.name === "NetworkError" && error.message.startsWith("Resource not found");
}

function asRecord(value: unknown): Record<string, unknown> | undefined {
  return typeof value === "object" && value !== null
    ? (value as Record<string, unknown>)
    : undefined;
}

// The adapter logs every activity at info; keep only its warnings and errors.
const teamsLogger: Logger = {
  child: () => teamsLogger,
  debug: () => undefined,
  info: () => undefined,
  warn: (message, bindings) => getLogger().warn(`teams: ${message}`, logBindings(bindings)),
  error: (message, bindings) => getLogger().error(`teams: ${message}`, logBindings(bindings)),
};

function logBindings(value: unknown): Record<string, unknown> | undefined {
  return asRecord(value);
}

/**
 * The Chat SDK Teams adapter, loaded with the Microsoft Teams SDK only once
 * used so unconfigured processes never pay for them. Activities from any
 * tenant but the configured one, or with none, are refused before they reach
 * the adapter; the adapter then verifies Microsoft's JWT on the rest.
 */
export class LazyTeamsAdapter implements Adapter<TeamsThreadId, unknown> {
  readonly name = "teams";
  private loaded: Promise<TeamsAdapter> | undefined;
  private implementation: TeamsAdapter | undefined;
  constructor(readonly config: TeamsCredentials) {}

  get userName() {
    return this.implementation?.userName ?? "teams";
  }
  get botUserId() {
    return this.implementation?.botUserId;
  }
  private load() {
    this.loaded ??= (async () => {
      const { createTeamsAdapter } = await import("@chat-adapter/teams");
      // Every credential is explicit so the adapter never falls back to process.env.
      const adapter = createTeamsAdapter({
        appId: this.config.appId,
        appPassword: this.config.appPassword,
        appTenantId: this.config.tenantId,
        appType: "SingleTenant",
        logger: teamsLogger,
      });
      boundInboundMarkdown(adapter);
      this.implementation = adapter;
      return adapter;
    })().catch((error: unknown) => {
      this.loaded = undefined;
      throw error;
    });
    return this.loaded;
  }
  private get ready(): TeamsAdapter {
    if (!this.implementation) throw new Error("Teams adapter is not initialized");
    return this.implementation;
  }
  async initialize(chat: ChatInstance) {
    await (await this.load()).initialize(chat);
  }
  async handleWebhook(request: Request, options?: WebhookOptions) {
    const body = await request.text();
    let activity: unknown;
    try {
      activity = JSON.parse(body);
    } catch {
      return new Response("Invalid JSON", { status: 400 });
    }
    // Refusing here only saves work: a matching tenant still needs a valid token.
    if (teamsActivityTenant(activity) !== this.config.tenantId.toLowerCase()) {
      return new Response("Forbidden", { status: 403 });
    }
    const forwarded = new Request(request.url, {
      method: request.method,
      headers: request.headers,
      body,
    });
    return (await this.load()).handleWebhook(forwarded, options);
  }
  async postMessage(threadId: string, message: AdapterPostableMessage) {
    return (await this.load()).postMessage(threadId, message);
  }
  async editMessage(threadId: string, messageId: string, message: AdapterPostableMessage) {
    return (await this.load()).editMessage(threadId, messageId, message);
  }
  async deleteMessage(threadId: string, messageId: string) {
    await (await this.load()).deleteMessage(threadId, messageId);
  }
  async addReaction(threadId: string, messageId: string, emoji: EmojiValue | string) {
    await (await this.load()).addReaction(threadId, messageId, emoji);
  }
  async removeReaction(threadId: string, messageId: string, emoji: EmojiValue | string) {
    await (await this.load()).removeReaction(threadId, messageId, emoji);
  }
  async startTyping(threadId: string) {
    await (await this.load()).startTyping(threadId);
  }
  async fetchMessages(threadId: string, options?: FetchOptions) {
    return (await this.load()).fetchMessages(threadId, options);
  }
  async fetchThread(threadId: string) {
    return (await this.load()).fetchThread(threadId);
  }
  async openDM(userId: string) {
    return (await this.load()).openDM(userId);
  }
  parseMessage(raw: unknown) {
    return this.ready.parseMessage(raw);
  }
  renderFormatted(content: FormattedContent) {
    return this.implementation?.renderFormatted(content) ?? stringifyMarkdown(content);
  }
  encodeThreadId(platformData: TeamsThreadId) {
    return this.ready.encodeThreadId(platformData);
  }
  decodeThreadId(threadId: string) {
    return this.ready.decodeThreadId(threadId);
  }
  channelIdFromThreadId(threadId: string) {
    return this.ready.channelIdFromThreadId(threadId);
  }
  isDM(threadId: string) {
    return this.implementation?.isDM(threadId) ?? false;
  }
}
