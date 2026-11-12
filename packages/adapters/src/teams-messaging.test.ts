import type { IncomingMessage, Server } from "node:http";
import { createServer } from "node:http";
import type { AddressInfo } from "node:net";
import type * as TeamsAdapterModule from "@chat-adapter/teams";
import type { AdapterContext, MessagingInboundMessage } from "@rakazo/adapter-kit";
import { createLogger, createTestSink, installLogger } from "@rakazo/logging";
import { afterAll, beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import {
  ChatSdkMessagingSurface,
  MessagingConversationGoneError,
  MessagingSendError,
} from "./chat-sdk-surface.js";
import { messagingEnvFromProcess, messagingPlatformsFromEnv } from "./messaging-platforms.js";
import type { AgentTeamsAppRow } from "./teams-agent-app-platforms.js";
import {
  refreshAgentTeamsPlatforms,
  setActiveTeamsAgentAppRouter,
  TeamsAgentAppRouter,
} from "./teams-agent-app-platforms.js";
import {
  isTeamsBotSender,
  LazyTeamsAdapter,
  SHARED_TEAMS_BOT_APP_ID,
  teamsConversationIdOfThread,
  teamsPlatform,
  teamsProviderForAgent,
} from "./teams-messaging.js";

const teams = vi.hoisted(() => {
  const part = (value: unknown) => Buffer.from(JSON.stringify(value)).toString("base64url");
  // The Teams SDK reads the expiry out of the token, so it has to be a JWT.
  const token = [
    part({ alg: "none", typ: "JWT" }),
    part({ aud: "https://api.botframework.com", exp: 4_102_444_800 }),
    "unsigned",
  ].join(".");
  return { configs: [] as Record<string, unknown>[], token };
});

// Microsoft's JWT check needs its signing keys from the network, and the
// token for outbound calls comes from Entra; swap only those two seams.
vi.mock("@chat-adapter/teams", async (importOriginal) => {
  const actual = await importOriginal<typeof TeamsAdapterModule>();
  return {
    ...actual,
    createTeamsAdapter: (config: TeamsAdapterModule.TeamsAdapterConfig) => {
      teams.configs.push({ ...config });
      return actual.createTeamsAdapter({
        ...config,
        webhookVerifier: () => true,
        token: () => teams.token,
      });
    },
  };
});

const APP_ID = "00000000-0000-0000-0000-00000000a99";
const TENANT = "11111111-2222-3333-4444-555555555555";
const BOT = `28:${APP_ID}`;

type ConnectorRequest = { method: string; path: string; auth: string; body: unknown };

/** The Bot Connector service an activity's serviceUrl points at. */
const connector = {
  requests: [] as ConnectorRequest[],
  forbidden: new Set<string>(),
  /** Member list pages per conversation; a conversation without any answers 404. */
  members: new Map<string, Array<Array<{ id: string; name: string; userRole?: string }>>>(),
  failingMembers: new Set<string>(),
  server: undefined as Server | undefined,
  serviceUrl: "",
};

async function readBody(request: IncomingMessage): Promise<unknown> {
  const chunks: Buffer[] = [];
  for await (const chunk of request) chunks.push(chunk as Buffer);
  const text = Buffer.concat(chunks).toString("utf8");
  return text ? JSON.parse(text) : null;
}

beforeAll(async () => {
  connector.server = createServer(async (request, response) => {
    const path = decodeURIComponent(request.url ?? "");
    const body = await readBody(request);
    connector.requests.push({
      method: request.method ?? "",
      path,
      auth: request.headers.authorization ?? "",
      body,
    });
    const conversation = /^\/v3\/conversations\/([^/]+)/.exec(path)?.[1] ?? "";
    if (connector.forbidden.has(conversation)) {
      response.writeHead(403, { "content-type": "application/json" });
      response.end(JSON.stringify({ error: { code: "BotNotInConversationRoster" } }));
      return;
    }
    const pagedMembers = /^\/v3\/conversations\/[^/]+\/pagedmembers(?:\?|$)/i.test(path);
    if (pagedMembers && connector.failingMembers.has(conversation)) {
      response.writeHead(500, { "content-type": "application/json" });
      response.end(JSON.stringify({ error: { code: "ServiceError" } }));
      return;
    }
    const pages = connector.members.get(conversation);
    if (pagedMembers && pages) {
      const token = new URL(path, "http://connector.test").searchParams.get("continuationToken");
      const index = token ? Number(token.replace("page-", "")) : 0;
      response.writeHead(200, { "content-type": "application/json" });
      response.end(
        JSON.stringify({
          members: pages[index] ?? [],
          ...(index + 1 < pages.length ? { continuationToken: `page-${index + 1}` } : {}),
        }),
      );
      return;
    }
    if (path.startsWith("/v3/attachments/")) {
      response.writeHead(200, { "content-type": "image/png" });
      response.end(Buffer.from("png-bytes"));
      return;
    }
    if (path.includes("/members/")) {
      response.writeHead(200, { "content-type": "application/json" });
      response.end(
        JSON.stringify({ id: "29:ada", name: "Ada Lovelace", email: "ada@example.test" }),
      );
      return;
    }
    if (request.method === "POST" && path.endsWith("/activities")) {
      response.writeHead(200, { "content-type": "application/json" });
      response.end(JSON.stringify({ id: `sent-${connector.requests.length}` }));
      return;
    }
    response.writeHead(404);
    response.end();
  });
  await new Promise<void>((resolve) => connector.server!.listen(0, "127.0.0.1", resolve));
  const { port } = connector.server.address() as AddressInfo;
  connector.serviceUrl = `http://127.0.0.1:${port}/`;
});

afterAll(async () => {
  await new Promise((resolve) => connector.server?.close(resolve));
});

beforeEach(() => {
  connector.requests.length = 0;
  connector.forbidden.clear();
  connector.members.clear();
  connector.failingMembers.clear();
});

const context: AdapterContext = {
  operationId: "op-test",
  traceId: "trace-test",
  spaceId: "space-test",
  userId: "user-test",
  signal: new AbortController().signal,
};

function teamsEnv() {
  return messagingEnvFromProcess({
    TEAMS_APP_ID: ` ${APP_ID} `,
    TEAMS_APP_PASSWORD: " app-secret ",
    TEAMS_APP_TENANT_ID: ` ${TENANT} `,
  });
}

function createTeamsSurface() {
  const surface = new ChatSdkMessagingSurface(messagingPlatformsFromEnv(teamsEnv()));
  const events: MessagingInboundMessage[] = [];
  surface.onInbound(async (event) => {
    if (event.type === "message") events.push(event);
  });
  return { surface, events };
}

let activitySeq = 0;

function activity(overrides: {
  conversation: Record<string, unknown>;
  text?: string;
  from?: Record<string, unknown>;
  entities?: unknown[];
  attachments?: unknown[];
  channelData?: Record<string, unknown>;
  tenant?: string;
}) {
  const tenant = overrides.tenant ?? TENANT;
  activitySeq += 1;
  return {
    type: "message",
    id: `17000000000${activitySeq}`,
    timestamp: "2026-11-04T12:00:00.000Z",
    channelId: "msteams",
    serviceUrl: connector.serviceUrl,
    from: overrides.from ?? { id: "29:ada", name: "Ada Lovelace", aadObjectId: "aad-ada" },
    recipient: { id: BOT, name: "Rakazo" },
    conversation: { tenantId: tenant, ...overrides.conversation },
    channelData: { tenant: { id: tenant }, ...overrides.channelData },
    text: overrides.text ?? "",
    entities: overrides.entities ?? [],
    attachments: overrides.attachments ?? [],
  };
}

const personal = { id: "a:1-personal-chat", conversationType: "personal" };
const groupChat = { id: "19:group-chat@thread.v2", conversationType: "groupChat", isGroup: true };
const mentionEntity = {
  type: "mention",
  text: "<at>Rakazo</at>",
  mentioned: { id: BOT, name: "Rakazo" },
};

function webhook(body: unknown): Request {
  return new Request("https://rakazo.test/api/v1/messaging/webhook/teams", {
    method: "POST",
    headers: { "content-type": "application/json", authorization: "Bearer from-teams" },
    body: JSON.stringify(body),
  });
}

async function deliver(body: unknown) {
  const { surface, events } = createTeamsSurface();
  const response = await surface.handleWebhook("teams", webhook(body));
  return { surface, events, status: response?.status };
}

function activityPosts() {
  return connector.requests.filter(
    (request) => request.method === "POST" && request.path.endsWith("/activities"),
  );
}

describe("Teams platform", () => {
  it("mounts only with app id, password and tenant", () => {
    expect(messagingPlatformsFromEnv({ teamsAppId: APP_ID, teamsAppPassword: "secret" })).toEqual(
      [],
    );
    const [platform] = messagingPlatformsFromEnv(teamsEnv());
    expect(platform?.provider).toBe("teams");
    expect(platform?.capabilities).toEqual({
      direct: true,
      groups: true,
      receipt: "typing",
      files: true,
      lines: true,
    });
    expect(platform?.adapter).toBeInstanceOf(LazyTeamsAdapter);
    expect((platform!.adapter as LazyTeamsAdapter).config).toEqual({
      appId: APP_ID,
      appPassword: "app-secret",
      tenantId: TENANT,
    });
  });

  it("loads the Teams SDK on first use, as a single-tenant app", async () => {
    const before = teams.configs.length;
    createTeamsSurface();
    expect(teams.configs).toHaveLength(before);

    await deliver(activity({ conversation: personal, text: "hi" }));
    expect(teams.configs).toHaveLength(before + 1);
    expect(teams.configs.at(-1)).toMatchObject({
      appId: APP_ID,
      appPassword: "app-secret",
      appTenantId: TENANT,
      appType: "SingleTenant",
    });
  });

  it("turns a 1:1 message into a direct inbound event", async () => {
    const { events, status } = await deliver(
      activity({ conversation: personal, text: "What is on my calendar?" }),
    );
    expect(status).toBe(200);
    expect(events).toHaveLength(1);
    expect(events[0]).toMatchObject({
      provider: "teams",
      isDirect: true,
      from: "29:ada",
      fromLabel: "Ada Lovelace",
      content: "What is on my calendar?",
      channelName: null,
      participants: [],
    });
    expect(events[0]?.threadId.startsWith("teams:")).toBe(true);
  });

  it("names the sender's directory account in 1:1 and group chats", async () => {
    const from = { id: "29:ada", name: "Ada Lovelace", aadObjectId: "AAD-OBJECT-ADA" };
    const tenant = TENANT.toUpperCase();
    const direct = await deliver(activity({ conversation: personal, text: "hi", from, tenant }));
    expect(direct.events[0]?.directoryUser).toEqual({
      tenantId: TENANT,
      objectId: "aad-object-ada",
    });

    const group = await deliver(
      activity({
        conversation: groupChat,
        text: "<at>Rakazo</at> hi",
        entities: [mentionEntity],
        from,
      }),
    );
    expect(group.events[0]?.directoryUser).toEqual({
      tenantId: TENANT,
      objectId: "aad-object-ada",
    });

    const guest = await deliver(
      activity({ conversation: personal, text: "hi", from: { id: "29:guest", name: "Guest" } }),
    );
    expect(guest.events[0]?.directoryUser).toBeUndefined();
  });

  it("turns a group chat @mention into a group event named after the chat", async () => {
    connector.members.set(groupChat.id, [
      [
        { id: "29:ada", name: "Ada Lovelace" },
        { id: BOT, name: "Rakazo" },
      ],
    ]);
    const { events } = await deliver(
      activity({
        conversation: { ...groupChat, name: "Ops leads" },
        text: "<at>Rakazo</at> what is due today?",
        entities: [mentionEntity],
      }),
    );
    expect(events).toHaveLength(1);
    expect(events[0]).toMatchObject({
      provider: "teams",
      isDirect: false,
      mentionsBot: true,
      channelName: "Ops leads",
      participants: ["29:ada"],
    });
    expect(events[0]?.content).toContain("what is due today?");
    expect(events[0]).not.toHaveProperty("workspaceId");
  });

  it("delivers group messages that do not mention the bot when Teams sends them", async () => {
    const { events } = await deliver(
      activity({ conversation: groupChat, text: "Rent roll is uploaded" }),
    );
    expect(events).toHaveLength(1);
    expect(events[0]).toMatchObject({ isDirect: false, content: "Rent roll is uploaded" });
    expect(events[0]).not.toHaveProperty("mentionsBot");
  });

  it("collapses channel threads onto the channel and posts replies there", async () => {
    const post = (root: string) =>
      activity({
        conversation: {
          id: `19:general@thread.tacv2;messageid=${root}`,
          conversationType: "channel",
          isGroup: true,
        },
        channelData: {
          team: { id: "19:team@thread.tacv2" },
          channel: { id: "19:general@thread.tacv2", name: "General" },
        },
        text: "<at>Rakazo</at> summarize this",
        entities: [mentionEntity],
      });
    const first = await deliver(post("1700000000100"));
    const second = await deliver(post("1700000000200"));
    expect(first.events[0]).toMatchObject({ channelName: "General", isDirect: false });
    expect(second.events[0]?.threadId).toBe(first.events[0]?.threadId);

    await first.surface.sendToThread(
      { threadId: first.events[0]!.threadId, body: "Summary" },
      context,
    );
    expect(activityPosts().at(-1)?.path).toBe(
      "/v3/conversations/19:general@thread.tacv2/activities",
    );
  });

  it("refuses activities from another tenant or with no tenant", async () => {
    const other = await deliver(
      activity({
        conversation: personal,
        text: "hello",
        tenant: "99999999-0000-0000-0000-000000000000",
      }),
    );
    expect(other.status).toBe(403);
    expect(other.events).toHaveLength(0);

    const { tenantId: _tenantId, ...conversation } = activity({
      conversation: personal,
    }).conversation;
    const missing = await deliver({
      ...activity({ conversation: personal, text: "hello" }),
      conversation,
      channelData: {},
    });
    expect(missing.status).toBe(403);
    expect(missing.events).toHaveLength(0);
    expect(connector.requests).toHaveLength(0);
  });

  it("accepts the tenant in any letter case", async () => {
    const { events, status } = await deliver(
      activity({ conversation: personal, text: "hi", tenant: TENANT.toUpperCase() }),
    );
    expect(status).toBe(200);
    expect(events).toHaveLength(1);
  });

  it("posts replies to the conversation's own service URL as markdown", async () => {
    const { surface, events } = await deliver(activity({ conversation: personal, text: "hi" }));
    const sent = await surface.sendToThread(
      { threadId: events[0]!.threadId, body: "**Done.** Three items are due." },
      context,
    );
    expect(activityPosts()[0]).toMatchObject({
      path: `/v3/conversations/${personal.id}/activities`,
      auth: `Bearer ${teams.token}`,
      body: { type: "message", text: "**Done.** Three items are due.", textFormat: "markdown" },
    });
    expect(sent.handle).toMatch(/^sent-/);

    await surface.acknowledge(events[0]!.threadId, events[0]!.handle, context);
    expect(activityPosts().at(-1)?.body).toMatchObject({ type: "typing" });
  });

  it("gives up on a conversation the bot was removed from", async () => {
    const { surface, events } = await deliver(
      activity({ conversation: groupChat, text: "<at>Rakazo</at> hi", entities: [mentionEntity] }),
    );
    connector.forbidden.add(groupChat.id);
    const failure = await surface
      .sendToThread({ threadId: events[0]!.threadId, body: "hello" }, context)
      .catch((error: unknown) => error);
    expect(failure).toBeInstanceOf(MessagingSendError);
    expect((failure as MessagingSendError).permanent).toBe(true);
  });

  it("hands 1:1 file attachments to the inbound sink with the bot's credentials", async () => {
    const imageUrl = `${connector.serviceUrl}v3/attachments/att-1/views/original`;
    const { events } = await deliver(
      activity({
        conversation: personal,
        text: "here is the receipt",
        attachments: [
          { contentType: "image/png", contentUrl: imageUrl, name: "receipt.png" },
          {
            contentType: "application/vnd.microsoft.teams.file.download.info",
            name: "lease.pdf",
            content: { downloadUrl: "https://files.example.com/lease.pdf", fileType: "pdf" },
          },
          { contentType: "text/html", content: "<p>here is the receipt</p>" },
        ],
      }),
    );
    expect(events).toHaveLength(1);
    const attachments = events[0]?.attachments ?? [];
    expect(attachments.map(({ name, mimeType }) => ({ name, mimeType }))).toEqual([
      { name: "receipt.png", mimeType: "image/png" },
      { name: "lease.pdf", mimeType: "application/pdf" },
    ]);
    const bytes = await attachments[0]!.fetch();
    expect(Buffer.from(bytes).toString()).toBe("png-bytes");
    expect(
      connector.requests.find((request) => request.path.startsWith("/v3/attachments/")),
    ).toMatchObject({ method: "GET", auth: `Bearer ${teams.token}` });
  });
});

describe("Teams group members", () => {
  const mention = (overrides: { conversation?: Record<string, unknown>; text?: string } = {}) =>
    activity({
      conversation: overrides.conversation ?? groupChat,
      text: overrides.text ?? "<at>Rakazo</at> hi",
      entities: [mentionEntity],
    });
  const memberReads = () =>
    connector.requests.filter((request) => /\/pagedmembers/i.test(request.path));
  const conversationUpdate = (members: {
    membersAdded?: Array<{ id: string }>;
    membersRemoved?: Array<{ id: string }>;
  }) => {
    activitySeq += 1;
    return {
      type: "conversationUpdate",
      id: `f:update-${activitySeq}`,
      channelId: "msteams",
      serviceUrl: connector.serviceUrl,
      from: { id: "29:ada", name: "Ada Lovelace", aadObjectId: "aad-ada" },
      recipient: { id: BOT, name: "Rakazo" },
      conversation: { tenantId: TENANT, ...groupChat },
      channelData: { tenant: { id: TENANT } },
      ...members,
    };
  };

  it("reads every page of the member list and leaves the bot out", async () => {
    connector.members.set(groupChat.id, [
      [
        { id: "29:ada", name: "Ada Lovelace", userRole: "user" },
        { id: BOT, name: "Rakazo" },
      ],
      [{ id: "29:grace", name: "Grace Hopper", userRole: "user" }],
      [{ id: "29:guest", name: "Guest", userRole: "guest" }],
    ]);
    const { events } = await deliver(mention());

    expect(events[0]?.participants).toEqual(["29:ada", "29:grace", "29:guest"]);
    expect(events[0]).not.toHaveProperty("participantsUnavailable");
    expect(memberReads().map((request) => request.path)).toEqual([
      `/v3/conversations/${groupChat.id}/pagedMembers?pageSize=500`,
      `/v3/conversations/${groupChat.id}/pagedMembers?pageSize=500&continuationToken=page-1`,
      `/v3/conversations/${groupChat.id}/pagedMembers?pageSize=500&continuationToken=page-2`,
    ]);
    expect(memberReads().every((request) => request.auth === `Bearer ${teams.token}`)).toBe(true);
  });

  it("marks the roster unavailable, without the request in the log, when the list fails", async () => {
    connector.failingMembers.add(groupChat.id);
    const sink = createTestSink();
    installLogger(createLogger({ service: "test", sinks: [sink] }));
    try {
      const { events, status } = await deliver(mention());
      expect(status).toBe(200);
      expect(events[0]).toMatchObject({ participants: [], participantsUnavailable: true });
      const logged = sink.events.filter((event) =>
        event.message.includes("member list unavailable"),
      );
      expect(logged).toEqual([
        expect.objectContaining({
          reason: "Teams member list request failed with status 500",
        }),
      ]);
      expect(JSON.stringify(sink.events)).not.toContain(teams.token);
    } finally {
      installLogger(createLogger({ service: "test", level: "off", sinks: [] }));
    }
  });

  it("leaves 1:1 messages without a member list read", async () => {
    await deliver(activity({ conversation: personal, text: "hi" }));
    expect(memberReads()).toHaveLength(0);
  });

  it("reuses one read for a minute, and reads again after members change", async () => {
    connector.members.set(groupChat.id, [[{ id: "29:ada", name: "Ada Lovelace" }]]);
    const { surface, events } = createTeamsSurface();
    await surface.handleWebhook("teams", webhook(mention()));
    await surface.handleWebhook("teams", webhook(mention({ text: "<at>Rakazo</at> again" })));
    expect(memberReads()).toHaveLength(1);
    expect(await surface.conversationMembers(events[0]!.threadId, context)).toEqual(["29:ada"]);
    expect(memberReads()).toHaveLength(1);

    connector.members.set(groupChat.id, [
      [
        { id: "29:ada", name: "Ada Lovelace" },
        { id: "29:grace", name: "Grace Hopper" },
      ],
    ]);
    const added = await surface.handleWebhook(
      "teams",
      webhook(conversationUpdate({ membersAdded: [{ id: "29:grace" }] })),
    );
    expect(added?.status).toBe(200);
    await surface.handleWebhook("teams", webhook(mention({ text: "<at>Rakazo</at> third" })));
    expect(memberReads()).toHaveLength(2);
    expect(events.at(-1)?.participants).toEqual(["29:ada", "29:grace"]);
  });

  it("reports people removed from the chat as departures of its conversation", async () => {
    connector.members.set(groupChat.id, [[{ id: "29:ada", name: "Ada Lovelace" }]]);
    const { surface, events } = createTeamsSurface();
    await surface.handleWebhook("teams", webhook(mention()));
    await surface.handleWebhook(
      "teams",
      webhook(conversationUpdate({ membersRemoved: [{ id: "29:grace" }] })),
    );
    expect(events).toHaveLength(2);
    expect(events[1]).toMatchObject({
      isDirect: false,
      threadId: events[0]!.threadId,
      content: "",
      departures: { addresses: ["29:grace"], botLeft: false },
    });

    await surface.handleWebhook(
      "teams",
      webhook(conversationUpdate({ membersRemoved: [{ id: BOT }] })),
    );
    expect(events[2]).toMatchObject({
      threadId: events[0]!.threadId,
      departures: { addresses: [], botLeft: true },
    });
  });

  it("tracks no members of a 1:1 chat", async () => {
    const { surface, events } = createTeamsSurface();
    await surface.handleWebhook(
      "teams",
      webhook({
        ...conversationUpdate({ membersRemoved: [{ id: BOT }] }),
        conversation: { tenantId: TENANT, ...personal },
      }),
    );
    expect(events).toHaveLength(0);
  });

  it("marks group messages as addressed only by an @mention of this bot", async () => {
    const mentioned = await deliver(mention({ text: "<at>Rakazo</at> ABCD-2345" }));
    expect(mentioned.events[0]).toMatchObject({
      mentionsBot: true,
      groupMentionRequired: true,
      contentWithoutMention: "ABCD-2345",
    });
    const otherMention = await deliver(
      activity({
        conversation: groupChat,
        text: "<at>Grace</at> can you look?",
        entities: [{ type: "mention", text: "<at>Grace</at>", mentioned: { id: "29:grace" } }],
      }),
    );
    expect(otherMention.events[0]).toMatchObject({ groupMentionRequired: true });
    expect(otherMention.events[0]).not.toHaveProperty("mentionsBot");
    expect(otherMention.events[0]).not.toHaveProperty("contentWithoutMention");
    const direct = await deliver(activity({ conversation: personal, text: "hi" }));
    expect(direct.events[0]).not.toHaveProperty("groupMentionRequired");
  });

  it("tells a conversation the bot was removed from apart from a failed member list", async () => {
    connector.members.set(groupChat.id, [[{ id: "29:ada", name: "Ada Lovelace" }]]);
    const { surface, events } = await deliver(mention());
    connector.forbidden.add(groupChat.id);
    const surfaceWithFreshCache = createTeamsSurface().surface;
    const failure = await surfaceWithFreshCache
      .conversationMembers(events[0]!.threadId, context)
      .catch((error: unknown) => error);
    expect(failure).toBeInstanceOf(MessagingConversationGoneError);
    expect(await surface.conversationMembers(events[0]!.threadId, context)).toEqual(["29:ada"]);
  });

  it("reports people added to the chat as arrivals", async () => {
    connector.members.set(groupChat.id, [[{ id: "29:ada", name: "Ada Lovelace" }]]);
    const { surface, events } = createTeamsSurface();
    await surface.handleWebhook("teams", webhook(mention()));
    await surface.handleWebhook(
      "teams",
      webhook(conversationUpdate({ membersAdded: [{ id: "29:grace" }] })),
    );
    expect(events[1]).toMatchObject({ threadId: events[0]!.threadId, arrivals: ["29:grace"] });
    expect(events[1]).not.toHaveProperty("departures");
  });

  it("ignores member changes the adapter refused", async () => {
    const { surface, events } = createTeamsSurface();
    const response = await surface.handleWebhook(
      "teams",
      webhook({
        ...conversationUpdate({ membersRemoved: [{ id: "29:grace" }] }),
        conversation: { tenantId: "99999999-0000-0000-0000-000000000000", ...groupChat },
        channelData: {},
      }),
    );
    expect(response?.status).toBe(403);
    expect(events).toHaveLength(0);
  });
});

describe("Teams agent bots", () => {
  const AGENT_APP_ID = "00000000-0000-0000-0000-0000000000b7";
  const AGENT_BOT = `28:${AGENT_APP_ID}`;

  function agentEnv(extra: Record<string, string> = {}) {
    return messagingEnvFromProcess({
      TEAMS_APP_ID: APP_ID,
      TEAMS_APP_PASSWORD: "app-secret",
      TEAMS_APP_TENANT_ID: TENANT,
      TEAMS_AGENT_BOTS: JSON.stringify({
        agentLeasing: { appId: AGENT_APP_ID, passwordEnv: "TEAMS_APP_PASSWORD_LEASING" },
        agentNoSecret: { appId: "00000000-0000-0000-0000-0000000000c1", passwordEnv: "MISSING" },
      }),
      TEAMS_APP_PASSWORD_LEASING: "leasing-secret",
      ...extra,
    });
  }

  function agentSurface() {
    const surface = new ChatSdkMessagingSurface(messagingPlatformsFromEnv(agentEnv()));
    const events: MessagingInboundMessage[] = [];
    surface.onInbound(async (event) => {
      if (event.type === "message") events.push(event);
    });
    return { surface, events };
  }

  it("mounts one provider per agent bot, with the password from its named variable", () => {
    const platforms = messagingPlatformsFromEnv(agentEnv());
    expect(platforms.map((platform) => platform.provider)).toEqual(["teams", "teams-agentleasing"]);
    const agent = platforms[1]!;
    expect(agent.agentId).toBe("agentLeasing");
    expect((agent.adapter as LazyTeamsAdapter).config).toEqual({
      appId: AGENT_APP_ID,
      appPassword: "leasing-secret",
      tenantId: TENANT,
    });
    expect(teamsProviderForAgent("agentLeasing", agentEnv().teamsAgentBots ?? [])).toBe(
      "teams-agentleasing",
    );
    expect(teamsProviderForAgent("agentNoSecret", agentEnv().teamsAgentBots ?? [])).toBe("teams");
    expect(SHARED_TEAMS_BOT_APP_ID).toBe("8ccaca5c-5964-4594-b8c8-01d6c27e7ef8");
  });

  it("routes each activity to the agent whose bot it was sent to and keeps the chats apart", async () => {
    connector.members.set(groupChat.id, [[{ id: "29:ada", name: "Ada Lovelace" }]]);
    const { surface, events } = agentSurface();
    const toAgent = {
      ...activity({ conversation: groupChat, text: "lease renewals are in" }),
      recipient: { id: AGENT_BOT, name: "Leasing" },
    };
    // Both bots post to the same URL: the recipient picks the agent.
    await surface.handleWebhook("teams", webhook(toAgent));
    await surface.handleWebhook(
      "teams",
      webhook(activity({ conversation: groupChat, text: "rent roll is up" })),
    );
    expect(events).toHaveLength(2);
    const [agentEvent, sharedEvent] = events;
    expect(agentEvent).toMatchObject({ provider: "teams-agentleasing", agentId: "agentLeasing" });
    expect(sharedEvent?.provider).toBe("teams");
    expect(sharedEvent).not.toHaveProperty("agentId");
    expect(agentEvent?.threadId.startsWith("teams-agentleasing:")).toBe(true);
    expect(agentEvent?.threadId).not.toBe(sharedEvent?.threadId);
    expect(teamsConversationIdOfThread(agentEvent!.threadId)).toBe(groupChat.id);
    expect(teamsConversationIdOfThread(sharedEvent!.threadId)).toBe(groupChat.id);

    // Each agent replies as its own bot.
    expect(teams.configs.map((config) => config.appId)).toContain(AGENT_APP_ID);
    const posts = activityPosts().length;
    await surface.sendToThread({ threadId: agentEvent!.threadId, body: "Noted." }, context);
    expect(activityPosts()).toHaveLength(posts + 1);
  });
});

describe("Teams bot loop guard", () => {
  const credentials = { teamsAppId: "app-a", teamsAppPassword: "pw", teamsAppTenantId: "t" };
  it("treats a 28: sender as a bot", () => {
    expect(isTeamsBotSender({ from: { id: "28:app-b" } })).toBe(true);
    expect(isTeamsBotSender({ from: { id: "29:user", aadObjectId: "u1" } })).toBe(false);
    expect(isTeamsBotSender({})).toBe(false);
  });
  it("drops group messages from another agent bot", () => {
    const platform = teamsPlatform(credentials, { provider: "teams-a", agentId: "a" });
    expect(platform?.acceptsGroup?.({ from: { id: "28:app-b" } })).toBe(false);
    expect(platform?.acceptsGroup?.({ from: { id: "29:user" } })).toBe(true);
  });
});

function appIdOf(surface: ChatSdkMessagingSurface, provider: string): string | undefined {
  return configOf(surface.mountedPlatform(provider))?.appId;
}

function configOf(platform: { adapter: unknown } | undefined) {
  return platform?.adapter instanceof LazyTeamsAdapter ? platform.adapter.config : undefined;
}

describe("Teams agent apps (runtime)", () => {
  const TONY_APP_ID = "00000000-0000-0000-0000-0000000000d1";
  const TONY_BOT = `28:${TONY_APP_ID}`;
  const ENV_LEASING_APP_ID = "00000000-0000-0000-0000-0000000000b7";
  const chatA = { id: "19:chat-a@thread.v2", conversationType: "groupChat", isGroup: true };
  const chatB = { id: "19:chat-b@thread.v2", conversationType: "groupChat", isGroup: true };

  function harness(rows: AgentTeamsAppRow[], secrets: Record<string, string>) {
    const env = messagingEnvFromProcess({
      TEAMS_APP_ID: APP_ID,
      TEAMS_APP_PASSWORD: "app-secret",
      TEAMS_APP_TENANT_ID: TENANT,
      TEAMS_AGENT_BOTS: JSON.stringify({
        botLeasing: { appId: ENV_LEASING_APP_ID, passwordEnv: "TEAMS_APP_PASSWORD_LEASING" },
      }),
      TEAMS_APP_PASSWORD_LEASING: "leasing-secret",
    });
    const surface = new ChatSdkMessagingSurface(messagingPlatformsFromEnv(env));
    const events: MessagingInboundMessage[] = [];
    surface.onInbound(async (event) => {
      if (event.type === "message") events.push(event);
    });
    const installs = { recorded: [] as string[][], removed: [] as string[][] };
    const state = { rows };
    const router = new TeamsAgentAppRouter({
      host: surface,
      tenantId: TENANT,
      envBots: env.teamsAgentBots ?? [],
      loadApps: async () => state.rows,
      readSecret: async (ref) => secrets[ref],
      installs: {
        recordChatInstall: async (appId, chatId, installId) => {
          installs.recorded.push([appId, chatId, installId ?? ""]);
        },
        markChatRemoved: async (appId, chatId) => {
          installs.removed.push([appId, chatId]);
        },
      },
      intervalMs: 0,
    });
    return { surface, events, installs, router, state };
  }

  const tonyRow: AgentTeamsAppRow = {
    id: "app-row-tony",
    botId: "botTony",
    status: "published",
    entraAppId: TONY_APP_ID,
    secretRef: "teams-app-bottony",
  };

  function toTony(
    conversation: Record<string, unknown>,
    text: string,
    from?: Record<string, unknown>,
  ) {
    return {
      ...activity({ conversation, text, ...(from ? { from } : {}) }),
      recipient: { id: TONY_BOT, name: "Tony" },
      entities: [
        { type: "mention", text: "<at>Tony</at>", mentioned: { id: TONY_BOT, name: "Tony" } },
      ],
    };
  }

  it("mounts published and visible apps at runtime and unmounts them without a restart", async () => {
    const { surface, router, state } = harness(
      [
        tonyRow,
        {
          ...tonyRow,
          id: "r2",
          botId: "botDraft",
          status: "registered",
          entraAppId: "00000000-0000-0000-0000-0000000000e1",
        },
      ],
      { "teams-app-bottony": "tony-secret" },
    );
    expect(surface.mountedPlatform("teams-bottony")).toBeUndefined();
    setActiveTeamsAgentAppRouter(router);
    await refreshAgentTeamsPlatforms();
    const tony = surface.mountedPlatform("teams-bottony");
    expect(tony?.agentId).toBe("botTony");
    expect(configOf(tony)).toEqual({
      appId: TONY_APP_ID,
      appPassword: "tony-secret",
      tenantId: TENANT,
    });
    expect(surface.mountedPlatform("teams-botdraft")).toBeUndefined();
    state.rows = [{ ...tonyRow, status: "removed" }];
    await refreshAgentTeamsPlatforms();
    expect(surface.mountedPlatform("teams-bottony")).toBeUndefined();
    setActiveTeamsAgentAppRouter(undefined);
  });

  it("skips an app whose secret cannot be read", async () => {
    const { surface, router } = harness([tonyRow], {});
    await router.refresh();
    expect(surface.mountedPlatform("teams-bottony")).toBeUndefined();
  });

  it("lets a database row win over the env bot for the same agent, and restores the env bot after", async () => {
    const leasingRow = {
      ...tonyRow,
      id: "row-leasing",
      botId: "botLeasing",
      entraAppId: "00000000-0000-0000-0000-0000000000f1",
      secretRef: "s-leasing",
    };
    const { surface, router, state } = harness([leasingRow], { "s-leasing": "db-secret" });
    expect(appIdOf(surface, "teams-botleasing")).toBe(ENV_LEASING_APP_ID);
    await router.refresh();
    expect(appIdOf(surface, "teams-botleasing")).toBe("00000000-0000-0000-0000-0000000000f1");
    state.rows = [];
    await router.refresh();
    expect(appIdOf(surface, "teams-botleasing")).toBe(ENV_LEASING_APP_ID);
  });

  it("a failed mount restores the env bot and does not abort the rest of the pass", async () => {
    const leasingRow = {
      ...tonyRow,
      id: "row-leasing",
      botId: "botLeasing",
      entraAppId: "00000000-0000-0000-0000-0000000000f1",
      secretRef: "s-leasing",
    };
    const { surface, router } = harness([leasingRow, tonyRow], {
      "s-leasing": "db-secret",
      "teams-app-bottony": "tony-secret",
    });
    const realMount = surface.mountPlatform.bind(surface);
    let failNext = true;
    const mount = vi.spyOn(surface, "mountPlatform").mockImplementation(async (platform) => {
      if (
        failNext &&
        platform.provider === "teams-botleasing" &&
        platform.routeKey !== ENV_LEASING_APP_ID
      ) {
        failNext = false;
        // Like the real host: the previous platform is dropped, then initialize throws.
        await surface.unmountPlatform(platform.provider);
        throw new Error("initialize failed");
      }
      await realMount(platform);
    });

    await router.refresh();
    // The env bot is back, and Tony (later in the same pass) still mounted.
    expect(appIdOf(surface, "teams-botleasing")).toBe(ENV_LEASING_APP_ID);
    expect(appIdOf(surface, "teams-bottony")).toBe(TONY_APP_ID);
    expect(router.mounted().has("teams-botleasing")).toBe(false);

    // The next pass retries and the database row wins.
    await router.refresh();
    expect(appIdOf(surface, "teams-botleasing")).toBe("00000000-0000-0000-0000-0000000000f1");
    mount.mockRestore();
  });

  it("never mounts two platforms on one app id", async () => {
    const clash = { ...tonyRow, entraAppId: APP_ID.toUpperCase() };
    const { surface, router } = harness([clash], { "teams-app-bottony": "x" });
    await router.refresh();
    expect(surface.mountedPlatform("teams-bottony")).toBeUndefined();
    expect(
      surface
        .platforms()
        .map((p) => p.provider)
        .sort(),
    ).toEqual(["teams", "teams-botleasing"]);
  });

  it("keeps one agent's two group chats apart across interleaved messages, replying in each as the agent", async () => {
    connector.members.set(chatA.id, [[{ id: "29:ada", name: "Ada Lovelace" }]]);
    connector.members.set(chatB.id, [[{ id: "29:bob", name: "Bob" }]]);
    const { surface, events, installs, router } = harness([tonyRow], {
      "teams-app-bottony": "tony-secret",
    });
    await router.refresh();
    const bob = { id: "29:bob", name: "Bob", aadObjectId: "aad-bob" };
    // Both arrive at the shared bot's webhook URL; the recipient picks Tony's app.
    for (const body of [
      toTony(chatA, "A1"),
      toTony(chatB, "B1", bob),
      toTony(chatA, "A2"),
      toTony(chatB, "B2", bob),
    ]) {
      expect((await surface.handleWebhook("teams", webhook(body)))?.status).toBeLessThan(300);
    }
    expect(events.map((event) => event.content.replace(/@?Tony/g, "").trim())).toEqual([
      "A1",
      "B1",
      "A2",
      "B2",
    ]);
    for (const event of events)
      expect(event).toMatchObject({
        provider: "teams-bottony",
        agentId: "botTony",
        isDirect: false,
        mentionsBot: true,
      });
    const threadA = events[0]!.threadId;
    const threadB = events[1]!.threadId;
    expect(threadA).not.toBe(threadB);
    expect(events[2]!.threadId).toBe(threadA);
    expect(events[3]!.threadId).toBe(threadB);
    expect(teamsConversationIdOfThread(threadA)).toBe(chatA.id);
    expect(teamsConversationIdOfThread(threadB)).toBe(chatB.id);
    expect(events[1]!.participants).toEqual(["29:bob"]);

    // Install recorded once per chat, on the first activity.
    expect(installs.recorded.map(([row, chat]) => `${row}|${chat}`)).toEqual([
      `app-row-tony|${chatA.id}`,
      `app-row-tony|${chatB.id}`,
    ]);

    // Interleaved replies land in their own chats, sent by Tony's app.
    expect(appIdOf(surface, "teams-bottony")).toBe(TONY_APP_ID);
    const before = activityPosts().length;
    await surface.sendToThread({ threadId: threadB, body: "reply B" }, context);
    await surface.sendToThread({ threadId: threadA, body: "reply A" }, context);
    const posts = activityPosts().slice(before);
    expect(posts.map((post) => post.path)).toEqual([
      `/v3/conversations/${chatB.id}/activities`,
      `/v3/conversations/${chatA.id}/activities`,
    ]);
    expect(teams.configs.some((config) => config.appId === TONY_APP_ID)).toBe(true);
  });

  it("records installs from installationUpdate and marks removed chats", async () => {
    const { surface, installs, router } = harness([tonyRow], {
      "teams-app-bottony": "tony-secret",
    });
    await router.refresh();
    const install = (action: string) => ({
      ...activity({ conversation: chatA }),
      type: "installationUpdate",
      action,
      recipient: { id: TONY_BOT, name: "Tony" },
    });
    expect((await surface.handleWebhook("teams", webhook(install("add"))))?.status).toBeLessThan(
      300,
    );
    expect(installs.recorded.map(([, chat]) => chat)).toEqual([chatA.id]);
    expect((await surface.handleWebhook("teams", webhook(install("remove"))))?.status).toBeLessThan(
      300,
    );
    expect(installs.removed).toEqual([["app-row-tony", chatA.id]]);
  });

  it("ignores group messages from other bots, the shared bot and other agent apps included", async () => {
    connector.members.set(chatA.id, [[{ id: "29:ada", name: "Ada Lovelace" }]]);
    const { surface, events, router } = harness([tonyRow], { "teams-app-bottony": "tony-secret" });
    await router.refresh();
    for (const sender of [
      `28:${APP_ID}`,
      `28:${ENV_LEASING_APP_ID}`,
      "28:00000000-0000-0000-0000-0000000000ff",
    ]) {
      await surface.handleWebhook(
        "teams",
        webhook(toTony(chatA, "hi", { id: sender, name: "Bot" })),
      );
    }
    expect(events).toEqual([]);
  });
});
