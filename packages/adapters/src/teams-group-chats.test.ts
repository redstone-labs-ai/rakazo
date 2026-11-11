import { toolRequiresApproval } from "@rakazo/core";
import { describe, expect, it } from "vitest";
import { selectBuiltinToolsForRun } from "./run-tool-selection.js";
import { TEAMS_CREATE_GROUP_CHAT_TOOL, teamsTools } from "./run-tools-teams.js";
import {
  TeamsGroupChatError,
  TeamsGroupChats,
  teamsGroupChatsFromEnv,
  teamsGroupThreadId,
} from "./teams-group-chats.js";
import { SHARED_TEAMS_BOT_APP_ID } from "./teams-messaging.js";

const ALPHA_APP = "aaaaaaaa-0000-0000-0000-000000000001";
const BETA_APP = "bbbbbbbb-0000-0000-0000-000000000002";
const SERVICE_URL = "https://smba.example.test/teams/";

interface Recorded {
  url: string;
  token?: string;
  body: string;
}

/**
 * A fake Microsoft identity, Graph and Bot Connector. A token is `tok:<appId>:<scope>`, so every
 * call shows which bot made it. `failures` maps a URL suffix to an error response.
 */
function fakeMicrosoft(
  options: { failures?: Record<string, { status: number; code: string; message: string }> } = {},
) {
  const calls: Recorded[] = [];
  let chats = 0;
  let tokenRequests = 0;
  const fetch = async (url: string, init?: RequestInit): Promise<Response> => {
    const body = String(init?.body ?? "");
    const auth = new Headers(init?.headers).get("authorization") ?? undefined;
    calls.push({ url, body, ...(auth ? { token: auth.replace(/^Bearer /, "") } : {}) });
    // Yield, so concurrent creates interleave.
    await new Promise((resolve) => setTimeout(resolve, Math.random() * 5));
    if (new URL(url).hostname === "login.microsoftonline.com") {
      tokenRequests += 1;
      const form = new URLSearchParams(body);
      const appId = form.get("client_id");
      if (form.get("client_secret") !== `secret-${appId}`)
        return Response.json({ error: "invalid_client" }, { status: 401 });
      return Response.json({
        access_token: `tok:${appId}:${form.get("scope")}`,
        expires_in: 3600,
      });
    }
    for (const [suffix, failure] of Object.entries(options.failures ?? {})) {
      if (url.endsWith(suffix))
        return Response.json(
          { error: { code: failure.code, message: failure.message } },
          { status: failure.status },
        );
    }
    if (url.endsWith("/chats")) {
      chats += 1;
      const appId = auth?.split(":")[1];
      return Response.json({ id: `19:chat-${chats}-${appId}@thread.v2` }, { status: 201 });
    }
    if (url.endsWith("/installedApps")) return new Response(null, { status: 201 });
    if (url.endsWith("/activities")) return Response.json({ id: `msg-${calls.length}` });
    return new Response("not found", { status: 404 });
  };
  return { fetch, calls, tokenRequests: () => tokenRequests };
}

function service(
  fake: ReturnType<typeof fakeMicrosoft>,
  overrides: Partial<ConstructorParameters<typeof TeamsGroupChats>[0]> = {},
) {
  return new TeamsGroupChats({
    tenantId: "tenant-1",
    shared: {
      appId: SHARED_TEAMS_BOT_APP_ID,
      appPassword: `secret-${SHARED_TEAMS_BOT_APP_ID}`,
      catalogAppId: "catalog-shared",
    },
    agents: new Map([
      ["alpha", { appId: ALPHA_APP, appPassword: `secret-${ALPHA_APP}`, catalogAppId: "cat-a" }],
      ["beta", { appId: BETA_APP, appPassword: `secret-${BETA_APP}`, catalogAppId: "cat-b" }],
    ]),
    rscConsentEnabled: false,
    rscPermissions: ["ChatMessage.Read.Chat"],
    serviceUrl: SERVICE_URL,
    fetch: fake.fetch,
    ...overrides,
  });
}

const input = (agentId: string) => ({
  agentId,
  members: ["ann@example.test", "11111111-2222-3333-4444-555555555555"],
  topic: "Move-in",
  firstMessage: "Hello",
});

const tokensOf = (calls: Recorded[], match: string) =>
  calls.filter((call) => call.url.includes(match)).map((call) => call.token);

describe("TeamsGroupChats", () => {
  it("creates, installs and posts with the agent's own bot", async () => {
    const fake = fakeMicrosoft();
    const created = await service(fake).create(input("alpha"));
    expect(created).toMatchObject({
      chatId: `19:chat-1-${ALPHA_APP}@thread.v2`,
      provider: "teams-alpha",
      appId: ALPHA_APP,
      installed: true,
      threadId: teamsGroupThreadId(created.chatId, SERVICE_URL),
    });
    expect(tokensOf(fake.calls, "/chats")).toEqual([
      `tok:${ALPHA_APP}:https://graph.microsoft.com/.default`,
      `tok:${ALPHA_APP}:https://graph.microsoft.com/.default`,
    ]);
    expect(tokensOf(fake.calls, "/activities")).toEqual([
      `tok:${ALPHA_APP}:https://api.botframework.com/.default`,
    ]);
    const chat = JSON.parse(fake.calls.find((call) => call.url.endsWith("/chats"))?.body ?? "{}");
    expect(chat).toMatchObject({ chatType: "group", topic: "Move-in" });
    expect(chat.members.map((m: Record<string, string>) => m["user@odata.bind"])).toEqual([
      "https://graph.microsoft.com/v1.0/users('ann%40example.test')",
      "https://graph.microsoft.com/v1.0/users('11111111-2222-3333-4444-555555555555')",
    ]);
    const activity = fake.calls.find((call) => call.url.endsWith("/activities"));
    expect(activity?.url).toBe(
      `${SERVICE_URL}v3/conversations/${encodeURIComponent(created.chatId)}/activities`,
    );
  });

  it("falls back to the shared bot for an agent without its own", async () => {
    const fake = fakeMicrosoft();
    const created = await service(fake).create(input("gamma"));
    expect(created).toMatchObject({ provider: "teams", appId: SHARED_TEAMS_BOT_APP_ID });
    const install = fake.calls.find((call) => call.url.endsWith("/installedApps"));
    expect(JSON.parse(install?.body ?? "{}")["teamsApp@odata.bind"]).toBe(
      "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/catalog-shared",
    );
  });

  it("refuses when no bot serves the agent", async () => {
    const fake = fakeMicrosoft();
    const chats = service(fake, { shared: undefined });
    await expect(chats.create(input("gamma"))).rejects.toBeInstanceOf(TeamsGroupChatError);
    expect(fake.calls).toHaveLength(0);
  });

  it("sends RSC consent only when the flag is on", async () => {
    const off = fakeMicrosoft();
    await service(off).create(input("alpha"));
    const offBody = JSON.parse(
      off.calls.find((call) => call.url.endsWith("/installedApps"))?.body ?? "{}",
    );
    expect(offBody).not.toHaveProperty("consentedPermissionSet");
    expect(offBody["teamsApp@odata.bind"]).toContain("/teamsApps/cat-a");

    const on = fakeMicrosoft();
    await service(on, {
      rscConsentEnabled: true,
      rscPermissions: ["ChatMessage.Read.Chat", "ChatMember.Read.Chat"],
    }).create(input("alpha"));
    const onBody = JSON.parse(
      on.calls.find((call) => call.url.endsWith("/installedApps"))?.body ?? "{}",
    );
    expect(onBody.consentedPermissionSet).toEqual({
      resourceSpecificPermissions: [
        { permissionValue: "ChatMessage.Read.Chat", permissionType: "Application" },
        { permissionValue: "ChatMember.Read.Chat", permissionType: "Application" },
      ],
    });
  });

  it("skips the install when the bot has no catalog id", async () => {
    const fake = fakeMicrosoft();
    const created = await service(fake, {
      agents: new Map([["alpha", { appId: ALPHA_APP, appPassword: `secret-${ALPHA_APP}` }]]),
    }).create(input("alpha"));
    expect(created.installed).toBe(false);
    expect(fake.calls.some((call) => call.url.endsWith("/installedApps"))).toBe(false);
  });

  it("keeps each agent's credentials apart when two agents create at once", async () => {
    const fake = fakeMicrosoft();
    const chats = service(fake);
    const results = await Promise.all([
      chats.create(input("alpha")),
      chats.create(input("beta")),
      chats.create(input("alpha")),
      chats.create(input("beta")),
    ]);
    expect(results.map((result) => result.appId)).toEqual([
      ALPHA_APP,
      BETA_APP,
      ALPHA_APP,
      BETA_APP,
    ]);
    expect(results.map((result) => result.provider)).toEqual([
      "teams-alpha",
      "teams-beta",
      "teams-alpha",
      "teams-beta",
    ]);
    // Every chat id names the bot whose token created it, and no call used another bot's token.
    for (const result of results) expect(result.chatId).toContain(result.appId);
    for (const call of fake.calls.filter((c) => c.token)) {
      const chatId = decodeURIComponent(call.url.split("/chats/")[1]?.split("/")[0] ?? "");
      const conversation = decodeURIComponent(
        call.url.split("/conversations/")[1]?.split("/")[0] ?? "",
      );
      const target = chatId || conversation;
      if (target) expect(target).toContain(String(call.token?.split(":")[1]));
    }
    // One token per (bot, scope): two bots times Graph and Bot Connector.
    expect(fake.tokenRequests()).toBe(4);
  });

  it("reuses a cached token until it nearly expires", async () => {
    const fake = fakeMicrosoft();
    let now = 0;
    const chats = service(fake, { now: () => now });
    await chats.create(input("alpha"));
    await chats.create(input("alpha"));
    expect(fake.tokenRequests()).toBe(2);
    now = 3_600_000;
    await chats.create(input("alpha"));
    expect(fake.tokenRequests()).toBe(4);
  });

  it("surfaces ResourceSpecificPermissionsMismatch once, without a retry", async () => {
    const fake = fakeMicrosoft({
      failures: {
        "/installedApps": {
          status: 400,
          code: "ResourceSpecificPermissionsMismatch",
          message: "Consented permissions do not match the app manifest.",
        },
      },
    });
    const error = await service(fake, { rscConsentEnabled: true })
      .create(input("alpha"))
      .catch((caught: unknown) => caught);
    expect(error).toBeInstanceOf(TeamsGroupChatError);
    expect(error).toMatchObject({
      step: "install",
      status: 400,
      code: "ResourceSpecificPermissionsMismatch",
    });
    expect(String((error as Error).message)).toContain("ResourceSpecificPermissionsMismatch");
    expect((error as TeamsGroupChatError).chat).toMatchObject({
      chatId: `19:chat-1-${ALPHA_APP}@thread.v2`,
      provider: "teams-alpha",
    });
    expect(fake.calls.filter((call) => call.url.endsWith("/installedApps"))).toHaveLength(1);
    expect(fake.calls.some((call) => call.url.endsWith("/activities"))).toBe(false);
  });

  it("surfaces a chat create failure with no chat, and a bad secret", async () => {
    const forbidden = fakeMicrosoft({
      failures: { "/chats": { status: 403, code: "Forbidden", message: "Missing Chat.Create" } },
    });
    await expect(service(forbidden).create(input("alpha"))).rejects.toMatchObject({
      step: "create",
      status: 403,
      code: "Forbidden",
      chat: undefined,
    });
    expect(forbidden.calls.filter((call) => call.url.endsWith("/chats"))).toHaveLength(1);

    const badSecret = fakeMicrosoft();
    const chats = service(badSecret, {
      agents: new Map([["alpha", { appId: ALPHA_APP, appPassword: "wrong" }]]),
    });
    await expect(chats.create(input("alpha"))).rejects.toMatchObject({
      step: "token",
      code: "invalid_client",
    });
    // A failed token is not cached: the next call asks again.
    await expect(chats.create(input("alpha"))).rejects.toMatchObject({ step: "token" });
    expect(badSecret.tokenRequests()).toBe(2);
  });
});

describe("teamsGroupChatsFromEnv", () => {
  const base = {
    RAKAZO_TEAMS_GROUP_CHATS: "true",
    TEAMS_APP_TENANT_ID: "tenant-1",
    TEAMS_APP_PASSWORD: "shared-secret",
    TEAMS_AGENT_BOTS: JSON.stringify({
      alpha: { appId: ALPHA_APP, passwordEnv: "PW_ALPHA", catalogAppId: "cat-a" },
    }),
    PW_ALPHA: "alpha-secret",
  };

  it("is off without the flag or the tenant", () => {
    expect(
      teamsGroupChatsFromEnv({ ...base, RAKAZO_TEAMS_GROUP_CHATS: undefined }),
    ).toBeUndefined();
    expect(teamsGroupChatsFromEnv({ ...base, TEAMS_APP_TENANT_ID: "" })).toBeUndefined();
  });

  it("reads agent bots, catalog ids and the shared fallback", () => {
    const chats = teamsGroupChatsFromEnv(base);
    expect(chats?.botFor("alpha")).toEqual({
      provider: "teams-alpha",
      bot: { appId: ALPHA_APP, appPassword: "alpha-secret", catalogAppId: "cat-a" },
    });
    expect(chats?.botFor("other")).toEqual({
      provider: "teams",
      bot: { appId: SHARED_TEAMS_BOT_APP_ID, appPassword: "shared-secret" },
    });
  });
});

describe("teams_create_group_chat tool", () => {
  function fakeRun(chats: TeamsGroupChats | undefined, botId = "alpha") {
    const channels: Record<string, unknown>[] = [];
    const ctx = {
      deps: {
        teamsGroupChats: chats,
        prisma: {
          messagingChannel: {
            upsert: async ({ create }: { create: Record<string, unknown> }) => {
              channels.push(create);
              return { id: `channel-${channels.length}` };
            },
          },
        },
      },
      run: { id: "run-1", botId, userId: "user-1", spaceId: "space-1", threadId: "thread-1" },
      secrets: ["sk-secret"],
    };
    const call = (args: Record<string, unknown>) =>
      teamsTools[TEAMS_CREATE_GROUP_CHAT_TOOL]!(
        ctx as never,
        {
          name: TEAMS_CREATE_GROUP_CHAT_TOOL,
          args,
          executionId: "call-1",
          finish: async (result: unknown) => result,
        } as never,
      );
    return { channels, call };
  }

  it("creates the chat as the run's agent and registers it", async () => {
    const fake = fakeMicrosoft();
    const { channels, call } = fakeRun(service(fake));
    const result = await call({
      members: ["ann@example.test", "bob@example.test"],
      topic: "Move-in",
      message: "Key sk-secret ready",
    });
    expect(result).toMatchObject({ ok: true, provider: "teams-alpha", channelId: "channel-1" });
    expect(channels[0]).toMatchObject({ provider: "teams-alpha", name: "Move-in" });
    const activity = fake.calls.find((c) => c.url.endsWith("/activities"));
    expect(JSON.parse(activity?.body ?? "{}").text).toBe("Key [redacted] ready");
  });

  it("registers the chat and returns its id when the first message fails", async () => {
    const fake = fakeMicrosoft({
      failures: {
        "/activities": { status: 403, code: "BotNotInConversationRoster", message: "no" },
      },
    });
    const { channels, call } = fakeRun(service(fake));
    const result = await call({ members: ["a@example.test", "b@example.test"], message: "Hi" });
    expect(result).toMatchObject({
      code: "BotNotInConversationRoster",
      chatId: `19:chat-1-${ALPHA_APP}@thread.v2`,
      provider: "teams-alpha",
      channelId: "channel-1",
      retryable: false,
    });
    expect(channels).toHaveLength(1);
    expect(channels[0]).not.toHaveProperty("introPostedAt");
    expect(fake.calls.filter((c) => c.url.endsWith("/activities"))).toHaveLength(1);
  });

  it("returns an install error with the chat id and registers the chat", async () => {
    const fake = fakeMicrosoft({
      failures: {
        "/installedApps": {
          status: 400,
          code: "ResourceSpecificPermissionsMismatch",
          message: "mismatch",
        },
      },
    });
    const { channels, call } = fakeRun(service(fake));
    const result = await call({ members: ["a@example.test", "b@example.test"], message: "Hi" });
    expect(result).toMatchObject({
      code: "ResourceSpecificPermissionsMismatch",
      chatId: `19:chat-1-${ALPHA_APP}@thread.v2`,
      channelId: "channel-1",
      retryable: false,
      error: expect.stringContaining("ResourceSpecificPermissionsMismatch"),
    });
    expect(channels).toEqual([expect.objectContaining({ provider: "teams-alpha" })]);
  });

  it.each([
    { members: ["a@example.test"], message: "Hi" },
    { members: ["a@example.test", "not a user"], message: "Hi" },
    { members: ["a@example.test", "b@example.test"], message: "  " },
    { members: ["a@example.test", "b@example.test"], message: "Hi", extra: true },
  ])("rejects invalid input: %o", async (args) => {
    const fake = fakeMicrosoft();
    const { call } = fakeRun(service(fake));
    await expect(call(args)).resolves.toHaveProperty("error");
    expect(fake.calls).toHaveLength(0);
  });

  it("says so when the server has no Teams group chats", async () => {
    const { call } = fakeRun(undefined);
    await expect(
      call({ members: ["a@example.test", "b@example.test"], message: "Hi" }),
    ).resolves.toEqual({ error: expect.stringContaining("not enabled") });
  });

  it("needs approval and is offered only to owner runs on enabled servers", () => {
    expect(toolRequiresApproval(TEAMS_CREATE_GROUP_CHAT_TOOL, false)).toBe(true);
    const offered = (trigger: string, enabled: boolean) =>
      selectBuiltinToolsForRun({
        graphicalToolsAllowed: false,
        groupId: null,
        trigger,
        semanticMemoryEnabled: false,
        messagingChannelRun: false,
        teamsGroupChatsEnabled: enabled,
      }).some((tool) => tool.name === TEAMS_CREATE_GROUP_CHAT_TOOL);
    expect(offered("user", true)).toBe(true);
    expect(offered("user", false)).toBe(false);
    for (const trigger of ["routine", "webhook", "bot_message"])
      expect(offered(trigger, true), trigger).toBe(false);
  });
});
