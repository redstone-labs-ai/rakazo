import type { IncomingMessage, Server } from "node:http";
import { createServer } from "node:http";
import type { AddressInfo } from "node:net";
import type * as TeamsAdapterModule from "@chat-adapter/teams";
import type { AdapterContext, MessagingInboundMessage } from "@rakazo/adapter-kit";
import { afterAll, beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import { ChatSdkMessagingSurface, MessagingSendError } from "./chat-sdk-surface.js";
import { messagingEnvFromProcess, messagingPlatformsFromEnv } from "./messaging-platforms.js";
import { LazyTeamsAdapter } from "./teams-messaging.js";

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

  it("names the 1:1 sender's directory account, and no group sender's", async () => {
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
    expect(group.events[0]).toBeDefined();
    expect(group.events[0]?.directoryUser).toBeUndefined();

    const guest = await deliver(
      activity({ conversation: personal, text: "hi", from: { id: "29:guest", name: "Guest" } }),
    );
    expect(guest.events[0]?.directoryUser).toBeUndefined();
  });

  it("turns a group chat @mention into a group event named after the chat", async () => {
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
      participants: [],
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
