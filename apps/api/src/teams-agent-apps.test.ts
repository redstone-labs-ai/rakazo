import { memoryAgentTeamsAppStore, visibilityStatus } from "@rakazo/adapters/teams-agent-apps";
import type { Actor } from "@rakazo/contracts";
import { Hono } from "hono";
import { describe, expect, it, vi } from "vitest";
import type { AgentTeamsAppRecord, TeamsAgentAppsDeps } from "./teams-agent-apps.js";
import { mountTeamsAgentAppRoutes, storeFindApp, teamsAppUserMessage } from "./teams-agent-apps.js";

const owner: Actor = {
  userId: "user-owner",
  spaceId: "space-1",
  email: "owner@example.com",
  isDeploymentOwner: false,
};
const stranger: Actor = { ...owner, userId: "user-other", email: "other@example.com" };
const now = new Date("2026-11-11T12:00:00.000Z");

function record(overrides: Partial<AgentTeamsAppRecord> = {}): AgentTeamsAppRecord {
  return {
    status: "published",
    publishedAt: new Date("2026-11-11T11:00:00.000Z"),
    firstInstalledAt: null,
    lastError: null,
    chats: [
      { chatId: "chat-a", installedAt: new Date("2026-11-11T11:30:00.000Z"), removedAt: null },
      {
        chatId: "chat-gone",
        installedAt: new Date("2026-11-11T11:10:00.000Z"),
        removedAt: new Date("2026-11-11T11:20:00.000Z"),
      },
    ],
    ...overrides,
  };
}

function setup(options: {
  actor?: Actor | null;
  enabled?: boolean;
  app?: AgentTeamsAppRecord | null;
}) {
  const start = vi.fn<TeamsAgentAppsDeps["start"]>(async () => undefined);
  const app = new Hono();
  mountTeamsAgentAppRoutes(app, {
    enabled: options.enabled ?? true,
    sessionActor: async () => (options.actor === undefined ? owner : options.actor),
    ownsBot: async (actor, botId) => actor.userId === owner.userId && botId === "bot-1",
    findApp: async () => options.app ?? null,
    visibilityStatus: (found) => ({
      step: found.firstInstalledAt ? "in_chats" : "visible",
      label: "Visible in @mentions",
      estimated: !found.firstInstalledAt,
    }),
    start,
    now: () => now,
  });
  const call = (method: string, path = "/api/bots/bot-1/teams-app") =>
    app.request(path, { method });
  return { call, start };
}

describe("teams agent app routes", () => {
  it("requires a session", async () => {
    const { call } = setup({ actor: null });
    expect((await call("GET")).status).toBe(401);
    expect((await call("POST")).status).toBe(401);
  });

  it("hides another user's bot behind 404 for every method", async () => {
    const { call, start } = setup({ actor: stranger, app: record() });
    for (const [method, path] of [
      ["GET", undefined],
      ["POST", undefined],
      ["POST", "/api/bots/bot-1/teams-app/retry"],
      ["DELETE", undefined],
    ] as const) {
      expect((await call(method, path)).status).toBe(404);
    }
    expect(start).not.toHaveBeenCalled();
  });

  it("404s an unknown bot for the owner too", async () => {
    const { call } = setup({});
    expect((await call("GET", "/api/bots/bot-nope/teams-app")).status).toBe(404);
  });

  it("reports disabled and refuses every change when the flag is off", async () => {
    const { call, start } = setup({ enabled: false, app: record() });
    const response = await call("GET");
    expect(response.status).toBe(200);
    expect(await response.json()).toEqual({ enabled: false, app: null });
    expect((await call("POST")).status).toBe(404);
    expect((await call("POST", "/api/bots/bot-1/teams-app/retry")).status).toBe(404);
    expect((await call("DELETE")).status).toBe(404);
    expect(start).not.toHaveBeenCalled();
  });

  it("returns null app when the bot has none", async () => {
    const { call } = setup({});
    expect(await (await call("GET")).json()).toEqual({ enabled: true, app: null });
  });

  it("returns step, estimate and only the chats the agent is still in", async () => {
    const { call } = setup({ app: record() });
    expect(await (await call("GET")).json()).toEqual({
      enabled: true,
      app: {
        status: "published",
        step: "visible",
        label: "Visible in @mentions",
        estimated: true,
        chatCount: 1,
        chats: [{ chatId: "chat-a", installedAt: "2026-11-11T11:30:00.000Z" }],
        lastError: null,
        userMessage: null,
        userMessageKind: null,
        publishedAt: "2026-11-11T11:00:00.000Z",
        firstInstalledAt: null,
      },
    });
  });

  it("treats a removed app as none", async () => {
    const { call } = setup({ app: record({ status: "removed" }) });
    expect((await (await call("GET")).json()).app).toBeNull();
  });

  it("queues provisioning with 202 and refuses when one already exists", async () => {
    const fresh = setup({});
    expect((await fresh.call("POST")).status).toBe(202);
    expect(fresh.start).toHaveBeenCalledWith("provision", "bot-1");

    const existing = setup({ app: record() });
    expect((await existing.call("POST")).status).toBe(409);
    expect(existing.start).not.toHaveBeenCalled();
  });

  it("retries only a failed app", async () => {
    const failed = setup({ app: record({ status: "failed", lastError: "boom" }) });
    expect((await failed.call("POST", "/api/bots/bot-1/teams-app/retry")).status).toBe(202);
    expect(failed.start).toHaveBeenCalledWith("retry", "bot-1");

    const healthy = setup({ app: record() });
    expect((await healthy.call("POST", "/api/bots/bot-1/teams-app/retry")).status).toBe(409);
  });

  it("queues removal with 202 and refuses when there is nothing to remove", async () => {
    const present = setup({ app: record() });
    expect((await present.call("DELETE")).status).toBe(202);
    expect(present.start).toHaveBeenCalledWith("deprovision", "bot-1");

    expect((await setup({}).call("DELETE")).status).toBe(409);
    expect((await setup({ app: record({ status: "removing" }) }).call("DELETE")).status).toBe(409);
  });
});

describe("teams agent app routes over the provisioner", () => {
  it("reads the store's row and active chats, and the provisioner's step", async () => {
    const store = memoryAgentTeamsAppStore(() => now);
    const row = await store.ensure("bot-1");
    await store.update(row.id, {
      status: "visible",
      publishedAt: new Date("2026-11-11T11:00:00.000Z"),
      firstInstalledAt: new Date("2026-11-11T11:30:00.000Z"),
    });
    await store.upsertChat({
      agentAppId: row.id,
      chatId: "chat-a",
      installId: null,
      installedAt: new Date("2026-11-11T11:30:00.000Z"),
    });
    const gone = await store.upsertChat({
      agentAppId: row.id,
      chatId: "chat-gone",
      installId: null,
      installedAt: new Date("2026-11-11T11:10:00.000Z"),
    });
    await store.markChatRemoved(gone.id, now);
    const app = new Hono();
    mountTeamsAgentAppRoutes(app, {
      enabled: true,
      sessionActor: async () => owner,
      ownsBot: async () => true,
      findApp: storeFindApp(store),
      visibilityStatus,
      start: async () => undefined,
      now: () => now,
    });
    const body = await (await app.request("/api/bots/bot-1/teams-app")).json();
    expect(body.app).toMatchObject({
      status: "visible",
      step: "in_chats",
      chatCount: 1,
      chats: [{ chatId: "chat-a" }],
    });
  });

  it("reports a failed provision as the failed step with its stored error", async () => {
    const store = memoryAgentTeamsAppStore(() => now);
    const row = await store.ensure("bot-1");
    await store.update(row.id, { status: "failed", lastError: "create Entra app failed (403)" });
    const found = await storeFindApp(store)("bot-1");
    expect(found?.status).toBe("failed");
    expect(found && visibilityStatus(found, now, 0).step).toBe("failed");
    expect(await storeFindApp(store)("bot-other")).toBeNull();
  });
});

describe("teamsAppUserMessage", () => {
  it.each([
    [
      "read app role assignments failed (404 Request_ResourceNotFound): Resource 'sp-1' does not exist or one of its queried reference-property objects are not present.",
      "replicating",
    ],
    ["grant Chat.Create failed (400 Request_ResourceNotFound): nope", "replicating"],
    ["catalog token failed: AADSTS700082 The refresh token has expired", "catalog_auth"],
    ["publish Teams catalog app failed (403 Forbidden): denied", "catalog_auth"],
    ["read Teams catalog app failed (401 InvalidAuthenticationToken): x", "catalog_auth"],
    [
      "create Entra app failed (403 Authorization_RequestDenied): Insufficient privileges",
      "forbidden",
    ],
    ["create Azure Bot failed (403 AuthorizationFailed): no", "forbidden"],
    ["app role assignment refused: provisioner is not an owner of the Entra app", "forbidden"],
    ["create Azure Bot failed (409 Conflict): taken", "unknown"],
  ] as const)("maps %s", (lastError, kind) => {
    const mapped = teamsAppUserMessage(lastError);
    expect(mapped?.kind).toBe(kind);
    expect(mapped?.message).not.toContain(lastError.slice(0, 20));
  });

  it("is null without an error, and plain English otherwise", () => {
    expect(teamsAppUserMessage(null)).toBeNull();
    expect(teamsAppUserMessage("Request_ResourceNotFound")?.message).toBe(
      "Microsoft is still setting up this agent's app. This usually takes under a minute — Retry.",
    );
    expect(teamsAppUserMessage("weird")?.message).toBe(
      "Something went wrong setting up Teams. Retry, or ask an admin.",
    );
  });

  it("returns userMessage from GET next to the redacted detail", async () => {
    const { call } = setup({
      app: record({
        status: "failed",
        lastError: "create Entra app failed (403 Authorization_RequestDenied): x",
      }),
    });
    const body = await (await call("GET")).json();
    expect(body.app.userMessage).toBe("Bot doesn't have permission to do this step. Ask an admin.");
    expect(body.app.userMessageKind).toBe("forbidden");
    expect(body.app.lastError).toContain("Authorization_RequestDenied");
  });
});
