import { inflateRawSync } from "node:zlib";
import { describe, expect, it, vi } from "vitest";
import { deprovisionAgentTeamsApp } from "./deprovision.js";
import type { AgentTeamsDeprovisionDeps, AgentTeamsSecretStore } from "./deps.js";
import type { FetchLike } from "./http.js";
import { buildAgentTeamsManifest, initialPng } from "./manifest.js";
import { markChatRemoved, provisionAgentTeamsApp, recordChatInstall } from "./provision.js";
import { memoryAgentTeamsAppStore } from "./store.js";
import { visibilityStatus } from "./visibility.js";

const APP_ID = "11111111-2222-3333-4444-555555555555";
const OBJECT_ID = "obj-1";
const SP_ID = "sp-1";
const GRAPH_SP_ID = "graph-sp";
const CATALOG_ID = "catalog-tony";
const SHARED_CATALOG_ID = "catalog-shared";

interface Call {
  method: string;
  url: string;
  body: unknown;
  auth: string;
}

/** A scripted fake of Graph, ARM and Bot Connector. `failOnce` fails the first matching call. */
function fakeMicrosoft(
  options: { failOnce?: (call: Call) => boolean; sharedInstalledIn?: string[] } = {},
) {
  const calls: Call[] = [];
  let failed = false;
  const assignedRoles = new Set<string>();
  let spCreated = false;
  let catalogPublished = false;
  const sharedInstalled = new Set(options.sharedInstalledIn ?? []);
  const json = (status: number, body?: unknown) =>
    new Response(body === undefined ? null : JSON.stringify(body), { status });
  const fetchImpl: FetchLike = async (url, init) => {
    const method = init?.method ?? "GET";
    const raw = init?.body;
    const body =
      typeof raw === "string"
        ? (safeParse(raw) ?? raw)
        : raw instanceof Uint8Array
          ? raw
          : undefined;
    const headers = new Headers(init?.headers);
    const call: Call = { method, url, body, auth: headers.get("authorization") ?? "" };
    calls.push(call);
    if (!failed && options.failOnce?.(call)) {
      failed = true;
      return json(503, { error: { code: "ServiceUnavailable", message: "try later" } });
    }
    if (url.includes("/applications(uniqueName=")) {
      return json(201, { id: OBJECT_ID, appId: APP_ID });
    }
    if (url.includes(`/servicePrincipals(appId='${APP_ID}')`)) {
      return spCreated ? json(200, { id: SP_ID }) : json(404, { error: { code: "NotFound" } });
    }
    if (url.endsWith("/servicePrincipals") && method === "POST") {
      spCreated = true;
      return json(201, { id: SP_ID });
    }
    if (url.includes("/addPassword")) return json(200, { secretText: "s3cret-value" });
    if (url.includes("Microsoft.BotService")) return json(method === "DELETE" ? 200 : 200, {});
    if (url.includes("servicePrincipals(appId='00000003")) {
      return json(200, {
        id: GRAPH_SP_ID,
        appRoles: [
          { id: "role-chat-create", value: "Chat.Create" },
          { id: "role-install", value: "TeamsAppInstallation.ReadWriteAndConsentSelfForChat.All" },
          { id: "role-other", value: "User.Read.All" },
        ],
      });
    }
    if (url.endsWith(`/servicePrincipals/${SP_ID}/appRoleAssignments`)) {
      if (method === "GET") {
        return json(200, {
          value: [...assignedRoles].map((appRoleId) => ({ appRoleId, resourceId: GRAPH_SP_ID })),
        });
      }
      const record = body as { appRoleId: string };
      assignedRoles.add(record.appRoleId);
      return json(201, {});
    }
    if (url.includes("/appCatalogs/teamsApps?$filter=")) {
      return json(200, { value: catalogPublished ? [{ id: CATALOG_ID }] : [] });
    }
    if (url.endsWith("/appCatalogs/teamsApps") && method === "POST") {
      catalogPublished = true;
      return json(201, { id: CATALOG_ID });
    }
    const chatInstall = url.match(/\/chats\/([^/]+)\/installedApps(.*)$/);
    if (chatInstall) {
      const chatId = decodeURIComponent(chatInstall[1] ?? "");
      if (method === "GET") {
        const wantsShared = decodeURIComponent(url).includes(SHARED_CATALOG_ID);
        return json(200, {
          value:
            wantsShared && sharedInstalled.has(chatId) ? [{ id: `shared-install-${chatId}` }] : [],
        });
      }
      if (method === "POST") {
        sharedInstalled.add(chatId);
        return json(201, {});
      }
      if (method === "DELETE") return json(204);
    }
    if (url.includes("/v3/conversations/")) return json(201, { id: "activity-1" });
    if (url.includes("/appCatalogs/teamsApps/") && method === "DELETE") return json(204);
    if (url.includes("/applications/") && method === "DELETE") return json(204);
    if (url.includes("login.microsoftonline.com")) {
      return json(200, { access_token: "agent-token", expires_in: 3600 });
    }
    return json(404, { error: { code: "Unexpected", message: `${method} ${url}` } });
  };
  return { fetch: vi.fn(fetchImpl), calls, assignedRoles };
}

function safeParse(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
}

function fakeSecrets() {
  const values = new Map<string, string>();
  const store: AgentTeamsSecretStore = {
    set: vi.fn(async (name: string, value: string) => {
      values.set(name, value);
    }),
    get: vi.fn(async (name: string) => values.get(name) ?? ""),
    delete: vi.fn(async (name: string) => {
      values.delete(name);
    }),
  };
  return { store, values };
}

function setup(ms = fakeMicrosoft()) {
  const store = memoryAgentTeamsAppStore(() => new Date("2026-11-11T12:00:00Z"));
  const secrets = fakeSecrets();
  const onPublished = vi.fn(async () => undefined);
  const deps: AgentTeamsDeprovisionDeps = {
    store,
    secrets: secrets.store,
    fetch: ms.fetch,
    config: {
      tenantId: "tenant-1",
      subscriptionId: "sub-1",
      resourceGroup: "rg-bots",
      messagingEndpoint: "https://api.example.test/api/webhooks/teams",
      developer: {
        name: "Redstone",
        websiteUrl: "https://example.test",
        privacyUrl: "https://example.test/privacy",
        termsOfUseUrl: "https://example.test/terms",
      },
    },
    provisionerToken: async (scope) => `provisioner:${scope}`,
    catalogToken: async () => "delegated-admin",
    now: () => new Date("2026-11-11T12:00:00Z"),
    onPublished,
    sharedBot: {
      catalogAppId: SHARED_CATALOG_ID,
      token: async (scope) => `shared:${scope}`,
      serviceUrl: "https://smba.example.test/teams",
    },
    agentToken: () => async (scope) => `agent:${scope}`,
  };
  return { ms, store, secrets, deps, onPublished };
}

const INPUT = { botId: "bot_tony", name: "Tony", description: "Leasing agent for Streetcar Flats" };

describe("provisionAgentTeamsApp", () => {
  it("runs every step and ends published", async () => {
    const { ms, deps, secrets, onPublished } = setup();
    const app = await provisionAgentTeamsApp(INPUT, deps);

    expect(app).toMatchObject({
      status: "published",
      entraAppId: APP_ID,
      entraObjectId: OBJECT_ID,
      secretRef: "teams-agent-bottony",
      azureBotName: "rakazo-agent-bot_tony",
      catalogAppId: CATALOG_ID,
      manifestVersion: "1.0.0",
      lastError: null,
    });
    expect(app.publishedAt).toEqual(new Date("2026-11-11T12:00:00Z"));
    expect(secrets.values.get("teams-agent-bottony")).toBe("s3cret-value");
    expect(onPublished).toHaveBeenCalledOnce();

    const entra = ms.calls.find((call) => call.url.includes("uniqueName"));
    expect(entra?.body).toEqual({ displayName: "Tony", signInAudience: "AzureADMyOrg" });

    const bot = ms.calls.find(
      (call) => call.method === "PUT" && call.url.includes("botServices/rakazo"),
    );
    expect(bot?.auth).toBe("Bearer provisioner:https://management.azure.com/.default");
    expect(bot?.body).toMatchObject({
      kind: "azurebot",
      sku: { name: "F0" },
      properties: {
        endpoint: "https://api.example.test/api/webhooks/teams",
        msaAppId: APP_ID,
        msaAppType: "SingleTenant",
        msaAppTenantId: "tenant-1",
      },
    });
    expect(ms.calls.some((call) => call.url.includes("/channels/MsTeamsChannel"))).toBe(true);
    expect([...ms.assignedRoles].sort()).toEqual(["role-chat-create", "role-install"]);

    const publish = ms.calls.find(
      (call) => call.method === "POST" && call.url.endsWith("/appCatalogs/teamsApps"),
    );
    expect(publish?.auth).toBe("Bearer delegated-admin");
    const files = unzip(publish?.body as Uint8Array);
    const manifest = JSON.parse(files.get("manifest.json")?.toString() ?? "{}");
    expect(manifest.bots[0].botId).toBe(APP_ID);
    expect(manifest.webApplicationInfo.id).toBe(APP_ID);
    expect(manifest.name.short).toBe("Tony");
    expect(manifest.authorization.permissions.resourceSpecific).toEqual([
      { name: "ChatMessage.Read.Chat", type: "Application" },
    ]);
    expect(files.get("color.png")?.subarray(1, 4).toString()).toBe("PNG");
    expect(files.get("outline.png")?.subarray(1, 4).toString()).toBe("PNG");
    // No secret ever reaches a stored error or the row.
    expect(JSON.stringify(app)).not.toContain("s3cret-value");
  });

  it("is a no-op once published", async () => {
    const { ms, deps } = setup();
    await provisionAgentTeamsApp(INPUT, deps);
    const before = ms.calls.length;
    await provisionAgentTeamsApp(INPUT, deps);
    expect(ms.calls.length).toBe(before);
  });

  it("resumes after a mid-way failure without redoing finished steps", async () => {
    const ms = fakeMicrosoft({
      failOnce: (call) => call.method === "PUT" && call.url.includes("/channels/MsTeamsChannel"),
    });
    const { deps, store, secrets } = setup(ms);

    await expect(provisionAgentTeamsApp(INPUT, deps)).rejects.toThrow(
      /enable Teams channel failed \(503/,
    );
    const failed = await store.findByBotId("bot_tony");
    expect(failed).toMatchObject({ status: "failed", entraAppId: APP_ID, azureBotName: null });
    expect(failed?.secretRef).toBe("teams-agent-bottony");
    expect(failed?.lastError).toContain("ServiceUnavailable");
    expect(failed?.lastError).not.toContain("s3cret");

    const app = await provisionAgentTeamsApp(INPUT, deps);
    expect(app.status).toBe("published");
    expect(app.lastError).toBeNull();
    const count = (match: (call: Call) => boolean) => ms.calls.filter(match).length;
    expect(count((call) => call.url.includes("uniqueName"))).toBe(1);
    expect(count((call) => call.url.includes("/addPassword"))).toBe(1);
    expect(secrets.store.set).toHaveBeenCalledTimes(1);
    expect(count((call) => call.url.endsWith("/appCatalogs/teamsApps"))).toBe(1);
  });
});

describe("recordChatInstall and markChatRemoved", () => {
  it("sets firstInstalledAt once and moves published to visible", async () => {
    const { deps, store } = setup();
    const app = await provisionAgentTeamsApp(INPUT, deps);
    await recordChatInstall(app.id, "19:a@thread.v2", "inst-a", {
      store,
      now: () => new Date("2026-11-11T12:10:00Z"),
    });
    await recordChatInstall(app.id, "19:b@thread.v2", "inst-b", {
      store,
      now: () => new Date("2026-11-11T13:00:00Z"),
    });
    const after = await store.findById(app.id);
    expect(after?.status).toBe("visible");
    expect(after?.firstInstalledAt).toEqual(new Date("2026-11-11T12:10:00Z"));
    expect(await store.chatCount(app.id)).toBe(2);
    await markChatRemoved(app.id, "19:a@thread.v2", { store });
    expect(await store.chatCount(app.id)).toBe(1);
  });
});

describe("deprovisionAgentTeamsApp", () => {
  it("hands each of two chats to the shared Bot, then deletes every resource", async () => {
    const ms = fakeMicrosoft({ sharedInstalledIn: ["19:b@thread.v2"] });
    const { deps, store, secrets } = setup(ms);
    const app = await provisionAgentTeamsApp(INPUT, deps);
    await recordChatInstall(app.id, "19:a@thread.v2", "inst-a", { store });
    await recordChatInstall(app.id, "19:b@thread.v2", "inst-b", { store });
    ms.calls.length = 0;

    const removed = await deprovisionAgentTeamsApp("bot_tony", deps, { agentName: "Tony" });

    expect(removed).toMatchObject({
      status: "removed",
      catalogAppId: null,
      azureBotName: null,
      entraObjectId: null,
      secretRef: null,
    });
    expect(await store.chatCount(app.id)).toBe(0);
    expect(secrets.values.size).toBe(0);

    const steps = ms.calls.map((call) => `${call.method} ${decodeURIComponent(call.url)}`);
    const sharedInstalls = ms.calls.filter(
      (call) => call.method === "POST" && call.url.includes("/installedApps"),
    );
    // Chat a lacked the shared Bot: installed, with RSC consent, as the shared bot. Chat b had it.
    expect(sharedInstalls).toHaveLength(1);
    expect(sharedInstalls[0]?.url).toContain(encodeURIComponent("19:a@thread.v2"));
    expect(sharedInstalls[0]?.auth).toBe("Bearer shared:https://graph.microsoft.com/.default");
    expect(sharedInstalls[0]?.body).toMatchObject({
      "teamsApp@odata.bind": expect.stringContaining(SHARED_CATALOG_ID),
      consentedPermissionSet: {
        resourceSpecificPermissions: [
          { permissionValue: "ChatMessage.Read.Chat", permissionType: "Application" },
        ],
      },
    });
    const handoffs = ms.calls.filter((call) => call.url.includes("/v3/conversations/"));
    expect(handoffs).toHaveLength(2);
    expect(handoffs[0]?.body).toEqual({
      type: "message",
      text: "Tony was removed. @Bot is here instead.",
    });
    expect(handoffs[0]?.auth).toBe("Bearer shared:https://api.botframework.com/.default");
    const uninstalls = ms.calls.filter(
      (call) => call.method === "DELETE" && call.url.includes("/installedApps/"),
    );
    expect(uninstalls.map((call) => call.url.split("/").pop())).toEqual(["inst-a", "inst-b"]);
    expect(uninstalls[0]?.auth).toBe("Bearer agent:https://graph.microsoft.com/.default");

    // Per chat: shared install (if needed) → handoff → uninstall; then the resources, in order.
    const index = (needle: string) => steps.findIndex((step) => step.includes(needle));
    expect(index("v3/conversations/19:a")).toBeGreaterThan(
      index("POST https://graph.microsoft.com/v1.0/chats/19:a"),
    );
    expect(index("DELETE https://graph.microsoft.com/v1.0/chats/19:a")).toBeGreaterThan(
      index("v3/conversations/19:a"),
    );
    expect(
      index(`DELETE https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${CATALOG_ID}`),
    ).toBeGreaterThan(index("DELETE https://graph.microsoft.com/v1.0/chats/19:b"));
    expect(index("DELETE https://management.azure.com")).toBeGreaterThan(
      index(`DELETE https://graph.microsoft.com/v1.0/appCatalogs`),
    );
    expect(
      index(`DELETE https://graph.microsoft.com/v1.0/applications/${OBJECT_ID}`),
    ).toBeGreaterThan(index("DELETE https://management.azure.com"));

    // Idempotent: a second run does nothing.
    const before = ms.calls.length;
    await deprovisionAgentTeamsApp("bot_tony", deps, { agentName: "Tony" });
    expect(ms.calls.length).toBe(before);
  });

  it("keeps the row removing with lastError and finishes on retry", async () => {
    const ms = fakeMicrosoft({
      failOnce: (call) => call.method === "DELETE" && call.url.includes("Microsoft.BotService"),
    });
    const { deps, store } = setup(ms);
    const app = await provisionAgentTeamsApp(INPUT, deps);
    await recordChatInstall(app.id, "19:a@thread.v2", "inst-a", { store });

    await expect(deprovisionAgentTeamsApp("bot_tony", deps)).rejects.toThrow(/delete Azure Bot/);
    expect(await store.findById(app.id)).toMatchObject({ status: "removing", catalogAppId: null });
    const removed = await deprovisionAgentTeamsApp("bot_tony", deps);
    expect(removed?.status).toBe("removed");
    // The chat was handed off once, not again on retry.
    expect(ms.calls.filter((call) => call.url.includes("/v3/conversations/"))).toHaveLength(1);
  });

  it("does nothing for a bot without an agent app", async () => {
    const { deps, ms } = setup();
    expect(await deprovisionAgentTeamsApp("bot_none", deps)).toBeNull();
    expect(ms.calls).toHaveLength(0);
  });
});

describe("visibilityStatus", () => {
  const published = new Date("2026-11-11T12:00:00Z");
  const base = { status: "published", publishedAt: published, firstInstalledAt: null };

  it("is published until five minutes pass", () => {
    expect(visibilityStatus(base, new Date("2026-11-11T12:04:59Z"))).toMatchObject({
      step: "published",
      estimated: false,
    });
  });

  it("estimates visible after five minutes", () => {
    expect(visibilityStatus(base, new Date("2026-11-11T12:05:00Z"))).toMatchObject({
      step: "visible",
      estimated: true,
      chatCount: 0,
    });
  });

  it("confirms visible on a first install, and in_chats with chats", () => {
    const installed = { ...base, status: "visible", firstInstalledAt: published };
    expect(visibilityStatus(installed, published)).toMatchObject({
      step: "visible",
      estimated: false,
    });
    expect(visibilityStatus(installed, published, 3)).toMatchObject({
      step: "in_chats",
      estimated: false,
      chatCount: 3,
    });
  });

  it("passes early, removal and failure steps through", () => {
    for (const status of ["creating", "registered", "removing", "removed", "failed"]) {
      const result = visibilityStatus(
        { status, publishedAt: null, firstInstalledAt: null },
        published,
      );
      expect(result.step).toBe(status);
      expect(result.estimated).toBe(false);
      expect(result.label).not.toBe("");
    }
  });
});

describe("manifest", () => {
  it("clips long names and writes a valid PNG", () => {
    const manifest = buildAgentTeamsManifest({
      appId: APP_ID,
      name: "A very long agent name that exceeds thirty characters",
      description: "",
      version: "1.0.0",
      developer: {
        name: "R",
        websiteUrl: "https://x",
        privacyUrl: "https://x",
        termsOfUseUrl: "https://x",
      },
    });
    expect(manifest.name.short.length).toBeLessThanOrEqual(30);
    expect(manifest.description.short).toContain("Rakazo agent");
    const png = initialPng("tony", 32, { foreground: [255, 255, 255] });
    expect(png.subarray(0, 8)).toEqual(
      Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    );
    expect(png.readUInt32BE(16)).toBe(32);
  });
});

/** Reads the local file entries of a zip this module wrote. */
function unzip(data: Uint8Array): Map<string, Buffer> {
  const buffer = Buffer.from(data);
  const files = new Map<string, Buffer>();
  let offset = 0;
  while (buffer.readUInt32LE(offset) === 0x04034b50) {
    const size = buffer.readUInt32LE(offset + 18);
    const nameLength = buffer.readUInt16LE(offset + 26);
    const name = buffer.subarray(offset + 30, offset + 30 + nameLength).toString();
    const start = offset + 30 + nameLength;
    files.set(name, inflateRawSync(buffer.subarray(start, start + size)));
    offset = start + size;
  }
  return files;
}
