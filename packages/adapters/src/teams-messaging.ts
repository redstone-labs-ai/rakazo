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
import type { MessagingMembershipChange, MessagingPlatform } from "./chat-sdk-surface.js";
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
 * Messages carry no roster, so the members come from the Bot Connector's member
 * list. Teams only delivers group messages that @mention the bot unless the app
 * manifest grants the ChatMessage.Read.Chat and ChannelMessage.Read.Group
 * resource-specific permissions.
 */
export function teamsPlatform(credentials: {
  teamsAppId?: string | undefined;
  teamsAppPassword?: string | undefined;
  teamsAppTenantId?: string | undefined;
}): MessagingPlatform | undefined {
  const { teamsAppId, teamsAppPassword, teamsAppTenantId } = credentials;
  if (!teamsAppId || !teamsAppPassword || !teamsAppTenantId) return undefined;
  const adapter = new LazyTeamsAdapter({
    appId: teamsAppId,
    appPassword: teamsAppPassword,
    tenantId: teamsAppTenantId,
  });
  return {
    provider: "teams",
    capabilities: { direct: true, groups: true, receipt: "typing", files: true },
    adapter,
    members: (threadId) => adapter.listMembers(threadId),
    membershipChanges: (raw) => adapter.membershipChange(raw),
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

/** The conversation type the adapter derives a thread id from, as it reads an activity. */
function teamsConversationType(
  activity: Record<string, unknown>,
): "personal" | "groupChat" | "channel" | undefined {
  const conversation = asRecord(activity.conversation);
  const explicit = conversation?.conversationType;
  if (explicit === "personal" || explicit === "groupChat" || explicit === "channel")
    return explicit;
  if (conversation?.isGroup === false) return "personal";
  if (conversation?.isGroup === true) {
    return asRecord(asRecord(activity.channelData)?.team)?.id ? "channel" : "groupChat";
  }
  return undefined;
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

/** Bot Connector member list pages: 500 is the largest page Teams serves. */
const MEMBER_PAGE_SIZE = 500;
/** Stop paging past this many pages (100,000 members) rather than loop on a bad token. */
const MEMBER_PAGE_LIMIT = 200;

/** The parts of the Teams SDK's Bot Connector client the member list uses. */
type MemberPage = {
  members: Array<{ id: string }>;
  continuationToken?: string;
};
type ConnectorApp = {
  apiFor(serviceUrl: string): {
    conversations: {
      getPagedMembers(
        conversationId: string,
        pageSize?: number,
        continuationToken?: string,
      ): Promise<MemberPage>;
    };
  };
};

/**
 * Only the HTTP status of a failed connector call: the SDK's error carries the request,
 * bot token included, so it never reaches a log.
 */
function memberListError(error: unknown): Error {
  const response = asRecord(asRecord(error)?.response);
  const status = response?.status;
  return new Error(
    typeof status === "number"
      ? `Teams member list request failed with status ${status}`
      : "Teams member list request failed",
  );
}

function accountIds(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return value.flatMap((account) => {
    const id = asRecord(account)?.id;
    return typeof id === "string" && id ? [id] : [];
  });
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
  private isOwnBot(accountId: string) {
    const appId = this.config.appId.toLowerCase();
    const id = accountId.toLowerCase();
    return id === appId || id.endsWith(`:${appId}`);
  }
  /**
   * Everyone in the conversation but this bot, from the Bot Connector's paged member list at
   * the conversation's own service URL, with the adapter's bot token.
   */
  async listMembers(threadId: string): Promise<string[]> {
    const adapter = await this.load();
    const { conversationId, serviceUrl } = adapter.decodeThreadId(threadId);
    // The connector client lives on the adapter's Teams app, which the adapter keeps protected.
    const app = (adapter as unknown as { app: ConnectorApp }).app;
    const conversations = app.apiFor(serviceUrl).conversations;
    const members: string[] = [];
    let continuationToken: string | undefined;
    for (let page = 0; page < MEMBER_PAGE_LIMIT; page += 1) {
      let result: MemberPage;
      try {
        result = await conversations.getPagedMembers(
          conversationId,
          MEMBER_PAGE_SIZE,
          continuationToken,
        );
      } catch (error) {
        throw memberListError(error);
      }
      members.push(...accountIds(result.members).filter((id) => !this.isOwnBot(id)));
      continuationToken = result.continuationToken || undefined;
      if (!continuationToken) return members;
    }
    throw new Error("Teams member list has too many pages");
  }
  /**
   * People added to or removed from a group chat or team, from a conversationUpdate activity.
   * A 1:1 chat has no members to track.
   */
  membershipChange(raw: unknown): MessagingMembershipChange | undefined {
    const activity = asRecord(raw);
    if (activity?.type !== "conversationUpdate" || !this.implementation) return undefined;
    const conversationType = teamsConversationType(activity);
    const conversationId = asRecord(activity.conversation)?.id;
    const { serviceUrl, id } = activity;
    if (
      conversationType === "personal" ||
      typeof conversationId !== "string" ||
      typeof serviceUrl !== "string" ||
      typeof id !== "string"
    )
      return undefined;
    const added = accountIds(activity.membersAdded);
    const removedIds = accountIds(activity.membersRemoved);
    if (added.length === 0 && removedIds.length === 0) return undefined;
    const removed = removedIds.filter((member) => !this.isOwnBot(member));
    return {
      threadId: teamsChannelThreadId(
        this.implementation.encodeThreadId({ conversationId, serviceUrl, conversationType }),
        raw,
      ),
      handle: id,
      added: added.filter((member) => !this.isOwnBot(member)),
      removed,
      botRemoved: removed.length < removedIds.length,
    };
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
