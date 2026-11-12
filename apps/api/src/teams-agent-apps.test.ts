import type { Actor } from "@rakazo/contracts";
import { Hono } from "hono";
import { describe, expect, it, vi } from "vitest";
import type { AgentTeamsAppRecord, TeamsAgentAppsDeps } from "./teams-agent-apps.js";
import { mountTeamsAgentAppRoutes } from "./teams-agent-apps.js";

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
