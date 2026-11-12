import { readdirSync, readFileSync, statSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import type { AgentAppSecurityEvent, GuardDeps } from "./app-role-guard.js";
import {
  AGENT_APP_TAG,
  AGENT_GRAPH_APP_ROLES,
  AppRoleGuardRefusal,
  agentEntraAppBody,
  assignAgentGraphAppRoles,
  checkEntraAppShape,
  checkOwnership,
  checkRoleRequest,
  checkServicePrincipal,
  checkTargetAppId,
  resolveGraphRoles,
} from "./app-role-guard.js";
import type { FetchLike } from "./http.js";
import { MICROSOFT_GRAPH_APP_ID } from "./http.js";
import { memoryAgentTeamsAppStore } from "./store.js";

const GRAPH = MICROSOFT_GRAPH_APP_ID;
const AGENT_APP_ID = "11111111-2222-3333-4444-555555555555";
const AGENT_OBJECT_ID = "agent-obj";
const AGENT_SP_ID = "agent-sp";
const GRAPH_SP_ID = "graph-sp";
const PROVISIONER_CLIENT_ID = "aaaaaaaa-0000-0000-0000-00000000000a";
const PROVISIONER_OID = "provisioner-oid";
const SHARED_BOT_APP_ID = "bbbbbbbb-0000-0000-0000-00000000000b";
const PROTECTED = {
  provisionerClientId: PROVISIONER_CLIENT_ID,
  provisionerObjectId: PROVISIONER_OID,
  sharedBotAppId: SHARED_BOT_APP_ID,
};
const ROLE_CHAT = "Chat.Create";
const ROLE_INSTALL = "TeamsAppInstallation.ReadWriteAndConsentSelfForChat.All";

describe("checkRoleRequest (allowlist)", () => {
  it.each(AGENT_GRAPH_APP_ROLES)("allows the application role %s on Graph", (role) => {
    expect(checkRoleRequest({ kind: "application", resourceAppId: GRAPH, role })).toEqual({
      ok: true,
    });
  });

  it("accepts Graph's app id in any case", () => {
    expect(
      checkRoleRequest({ kind: "application", resourceAppId: GRAPH.toUpperCase(), role: ROLE_CHAT })
        .ok,
    ).toBe(true);
  });

  it.each([
    "Application.ReadWrite.All",
    "Application.ReadWrite.OwnedBy",
    "AppRoleAssignment.ReadWrite.All",
    "RoleManagement.ReadWrite.Directory",
    "Directory.ReadWrite.All",
    "Chat.ReadWrite.All",
    "TeamsAppInstallation.ReadWriteForChat.All",
    "chat.create",
    "Chat.Create ",
    " Chat.Create",
    "Chat.Create.All",
    "",
  ])("refuses the role %j", (role) => {
    const result = checkRoleRequest({ kind: "application", resourceAppId: GRAPH, role });
    expect(result.ok).toBe(false);
  });

  it.each([
    "00000002-0000-0ff1-ce00-000000000000", // Exchange Online
    "00000003-0000-0ff1-ce00-000000000000", // SharePoint
    "797f4846-ba00-4fd7-ba43-dac1f8f63013", // Azure Service Management
    "00000003-0000-0000-c000-000000000001",
    "",
  ])("refuses an allowed role name on another resource (%s)", (resourceAppId) => {
    expect(checkRoleRequest({ kind: "application", resourceAppId, role: ROLE_CHAT }).ok).toBe(
      false,
    );
  });

  it("refuses delegated (oauth2 permission grant) requests, even for allowed roles", () => {
    for (const role of AGENT_GRAPH_APP_ROLES) {
      expect(checkRoleRequest({ kind: "delegated", resourceAppId: GRAPH, role })).toEqual({
        ok: false,
        reason: "only application roles may be granted",
      });
    }
  });
});

describe("checkTargetAppId", () => {
  const record = { entraAppId: AGENT_APP_ID, entraObjectId: AGENT_OBJECT_ID, status: "registered" };

  it.each(["creating", "registered"])("allows a recorded agent app while %s", (status) => {
    expect(checkTargetAppId({ ...record, status }, PROTECTED).ok).toBe(true);
  });

  it.each(["published", "visible", "removing", "removed", "failed", "weird"])(
    "refuses while %s",
    (status) => {
      expect(checkTargetAppId({ ...record, status }, PROTECTED).ok).toBe(false);
    },
  );

  it("refuses a row with no recorded Entra app", () => {
    expect(checkTargetAppId({ ...record, entraAppId: null }, PROTECTED).ok).toBe(false);
    expect(checkTargetAppId({ ...record, entraAppId: "  " }, PROTECTED).ok).toBe(false);
    expect(checkTargetAppId({ ...record, entraObjectId: null }, PROTECTED).ok).toBe(false);
  });

  it("refuses the provisioner itself, in any case", () => {
    const result = checkTargetAppId(
      { ...record, entraAppId: PROVISIONER_CLIENT_ID.toUpperCase() },
      PROTECTED,
    );
    expect(result).toEqual({ ok: false, reason: "target is the provisioner" });
  });

  it("refuses the shared Bot app", () => {
    expect(checkTargetAppId({ ...record, entraAppId: SHARED_BOT_APP_ID }, PROTECTED)).toEqual({
      ok: false,
      reason: "target is the shared Bot app",
    });
  });

  it("refuses Microsoft Graph itself", () => {
    expect(checkTargetAppId({ ...record, entraAppId: GRAPH }, PROTECTED).ok).toBe(false);
  });
});

describe("agentEntraAppBody and checkEntraAppShape", () => {
  const expected = { appId: AGENT_APP_ID, objectId: AGENT_OBJECT_ID };
  const created = { id: AGENT_OBJECT_ID, appId: AGENT_APP_ID, ...agentEntraAppBody("Tony") };

  it("creates single-tenant, tagged, prefixed apps that request nothing", () => {
    expect(agentEntraAppBody("  Tony \n Leasing ")).toEqual({
      displayName: "Rakazo agent: Tony Leasing",
      signInAudience: "AzureADMyOrg",
      tags: [AGENT_APP_TAG],
      requiredResourceAccess: [],
    });
    expect(agentEntraAppBody("x".repeat(500)).displayName).toHaveLength(120);
    expect(agentEntraAppBody("").displayName).toBe("Rakazo agent: agent");
  });

  it("accepts what this flow creates", () => {
    expect(checkEntraAppShape(created, expected).ok).toBe(true);
    expect(checkEntraAppShape({ ...created, requiredResourceAccess: undefined }, expected).ok).toBe(
      true,
    );
  });

  it.each([
    ["another app id", { appId: "99999999-0000-0000-0000-000000000000" }],
    ["no app id", { appId: "" }],
    ["another object id", { id: "other-obj" }],
    ["multi-tenant", { signInAudience: "AzureADMultipleOrgs" }],
    ["personal accounts", { signInAudience: "AzureADandPersonalMicrosoftAccount" }],
    ["no tag", { tags: [] }],
    ["tags not a list", { tags: AGENT_APP_TAG }],
    ["no prefix", { displayName: "Tony" }],
    [
      "requested API access",
      { requiredResourceAccess: [{ resourceAppId: GRAPH, resourceAccess: [] }] },
    ],
    ["malformed API access", { requiredResourceAccess: { resourceAppId: GRAPH } }],
  ])("refuses %s", (_label, change) => {
    expect(checkEntraAppShape({ ...created, ...change }, expected).ok).toBe(false);
  });
});

describe("checkOwnership", () => {
  it("passes when Graph lists the provisioner as owner", () => {
    expect(checkOwnership(["someone", PROVISIONER_OID.toUpperCase()], PROVISIONER_OID).ok).toBe(
      true,
    );
  });
  it("refuses an app the provisioner does not own", () => {
    expect(checkOwnership(["someone-else"], PROVISIONER_OID)).toEqual({
      ok: false,
      reason: "provisioner is not an owner of the Entra app",
    });
    expect(checkOwnership([], PROVISIONER_OID).ok).toBe(false);
  });
  it("refuses when the provisioner's own id is unknown", () => {
    expect(checkOwnership([""], "").ok).toBe(false);
  });
});

describe("checkServicePrincipal", () => {
  it("passes for the recorded app's own service principal", () => {
    expect(
      checkServicePrincipal({ id: AGENT_SP_ID, appId: AGENT_APP_ID }, AGENT_APP_ID, PROTECTED).ok,
    ).toBe(true);
  });
  it("refuses a missing principal, another app's, or the provisioner's", () => {
    expect(checkServicePrincipal({ id: "", appId: "" }, AGENT_APP_ID, PROTECTED).ok).toBe(false);
    expect(
      checkServicePrincipal({ id: AGENT_SP_ID, appId: "other" }, AGENT_APP_ID, PROTECTED).ok,
    ).toBe(false);
    expect(
      checkServicePrincipal({ id: PROVISIONER_OID, appId: AGENT_APP_ID }, AGENT_APP_ID, PROTECTED),
    ).toEqual({ ok: false, reason: "target is the provisioner" });
  });
});

function graphSp(overrides: Record<string, unknown> = {}) {
  const application = ["Application"];
  return {
    id: GRAPH_SP_ID,
    appId: GRAPH,
    appRoles: [
      { id: "r-chat", value: ROLE_CHAT, allowedMemberTypes: application, isEnabled: true },
      { id: "r-install", value: ROLE_INSTALL, allowedMemberTypes: application, isEnabled: true },
      { id: "r-dir", value: "Directory.ReadWrite.All", allowedMemberTypes: application },
    ],
    ...overrides,
  };
}

describe("resolveGraphRoles", () => {
  it("resolves exactly the two allowed roles by value", () => {
    const resolved = resolveGraphRoles(graphSp());
    expect(resolved).toEqual({
      ok: true,
      resourceId: GRAPH_SP_ID,
      roles: new Map([
        [ROLE_CHAT, "r-chat"],
        [ROLE_INSTALL, "r-install"],
      ]),
    });
  });

  it("refuses a resource that is not Graph's service principal", () => {
    expect(resolveGraphRoles(graphSp({ appId: "spoofed" })).ok).toBe(false);
    expect(resolveGraphRoles(graphSp({ id: "" })).ok).toBe(false);
    expect(resolveGraphRoles(undefined).ok).toBe(false);
  });

  it("ignores delegated-only and disabled roles with an allowed value", () => {
    const delegatedOnly = graphSp({
      appRoles: [
        { id: "r-chat", value: ROLE_CHAT, allowedMemberTypes: ["User"] },
        { id: "r-install", value: ROLE_INSTALL, allowedMemberTypes: ["Application"] },
      ],
    });
    expect(resolveGraphRoles(delegatedOnly).ok).toBe(false);
    const disabled = graphSp({
      appRoles: [
        { id: "r-chat", value: ROLE_CHAT, allowedMemberTypes: ["Application"], isEnabled: false },
        { id: "r-install", value: ROLE_INSTALL, allowedMemberTypes: ["Application"] },
      ],
    });
    expect(resolveGraphRoles(disabled).ok).toBe(false);
  });
});

interface Call {
  method: string;
  url: string;
  body: unknown;
  auth: string;
}

/** A fake Graph for the guard. Each knob breaks one fact the guard checks. */
function fakeGraph(
  options: {
    owners?: string[];
    entraApp?: Record<string, unknown>;
    sp?: Record<string, unknown> | null;
    graphSp?: Record<string, unknown>;
    assigned?: string[];
  } = {},
) {
  const calls: Call[] = [];
  const assigned = new Set(options.assigned ?? []);
  const json = (status: number, body?: unknown) =>
    new Response(body === undefined ? null : JSON.stringify(body), { status });
  const fetchImpl: FetchLike = async (url, init) => {
    const method = init?.method ?? "GET";
    const body: { appRoleId?: string } | undefined =
      typeof init?.body === "string" ? JSON.parse(init.body) : undefined;
    calls.push({
      method,
      url,
      body,
      auth: new Headers(init?.headers).get("authorization") ?? "",
    });
    if (url.includes(`/applications/${AGENT_OBJECT_ID}/owners`)) {
      return json(200, { value: (options.owners ?? [PROVISIONER_OID]).map((id) => ({ id })) });
    }
    if (url.includes(`/applications/${AGENT_OBJECT_ID}?`)) {
      return json(200, {
        id: AGENT_OBJECT_ID,
        appId: AGENT_APP_ID,
        ...agentEntraAppBody("Tony"),
        ...options.entraApp,
      });
    }
    if (url.includes(`/servicePrincipals(appId='${AGENT_APP_ID}')`)) {
      if (options.sp === null) return json(404, { error: { code: "NotFound" } });
      return json(200, { id: AGENT_SP_ID, appId: AGENT_APP_ID, ...options.sp });
    }
    if (url.includes(`/servicePrincipals(appId='${GRAPH}')`)) {
      return json(200, options.graphSp ?? graphSp());
    }
    if (url.endsWith("/appRoleAssignments")) {
      if (method === "GET") {
        return json(200, {
          value: [...assigned].map((appRoleId) => ({ appRoleId, resourceId: GRAPH_SP_ID })),
        });
      }
      assigned.add(body?.appRoleId ?? "");
      return json(201, {});
    }
    return json(404, { error: { code: "Unexpected", message: `${method} ${url}` } });
  };
  return { fetch: fetchImpl, calls, assigned };
}

async function guardSetup(
  graph = fakeGraph(),
  row: { entraAppId?: string | null; status?: string } = {},
) {
  const store = memoryAgentTeamsAppStore();
  const app = await store.ensure("bot_tony");
  await store.update(app.id, {
    entraAppId: row.entraAppId === undefined ? AGENT_APP_ID : row.entraAppId,
    entraObjectId: AGENT_OBJECT_ID,
    status: row.status ?? "registered",
  });
  const events: AgentAppSecurityEvent[] = [];
  const deps: GuardDeps = {
    store,
    fetch: graph.fetch,
    provisionerToken: async (scope) => `secret-token-for:${scope}`,
    identity: { clientId: PROVISIONER_CLIENT_ID, objectId: async () => PROVISIONER_OID },
    config: {
      tenantId: "tenant",
      subscriptionId: "sub",
      resourceGroup: "rg",
      messagingEndpoint: "https://example.test/teams",
      developer: {
        name: "Redstone",
        websiteUrl: "https://example.test",
        privacyUrl: "https://example.test/privacy",
        termsOfUseUrl: "https://example.test/terms",
      },
      sharedBotAppId: SHARED_BOT_APP_ID,
    },
    securityEvent: async (event) => {
      events.push(event);
    },
  };
  return { deps, events, graph };
}

const posts = (calls: Call[]) => calls.filter((call) => call.method === "POST");

describe("assignAgentGraphAppRoles", () => {
  it("grants exactly the two Graph roles to the agent's own service principal", async () => {
    const { deps, events, graph } = await guardSetup();
    await assignAgentGraphAppRoles(deps, "bot_tony");
    expect(posts(graph.calls).map((call) => [call.url, call.body])).toEqual([
      [
        `https://graph.microsoft.com/v1.0/servicePrincipals/${AGENT_SP_ID}/appRoleAssignments`,
        { principalId: AGENT_SP_ID, resourceId: GRAPH_SP_ID, appRoleId: "r-chat" },
      ],
      [
        `https://graph.microsoft.com/v1.0/servicePrincipals/${AGENT_SP_ID}/appRoleAssignments`,
        { principalId: AGENT_SP_ID, resourceId: GRAPH_SP_ID, appRoleId: "r-install" },
      ],
    ]);
    expect(events).toEqual([
      {
        type: "teams.agent_app.role_assigned",
        metadata: {
          botId: "bot_tony",
          entraAppId: AGENT_APP_ID,
          servicePrincipalId: AGENT_SP_ID,
          role: ROLE_CHAT,
          appRoleId: "r-chat",
          resourceId: GRAPH_SP_ID,
        },
      },
      expect.objectContaining({
        type: "teams.agent_app.role_assigned",
        metadata: expect.objectContaining({ role: ROLE_INSTALL }),
      }),
    ]);
    expect(JSON.stringify(events)).not.toContain("secret-token");
  });

  it("checks ownership and identity before the first grant", async () => {
    const { deps, graph } = await guardSetup();
    await assignAgentGraphAppRoles(deps, "bot_tony");
    const order = graph.calls.map((call) => `${call.method} ${call.url}`);
    const first = (needle: string) => order.findIndex((line) => line.includes(needle));
    const firstPost = order.findIndex((line) => line.startsWith("POST"));
    expect(first("/owners")).toBeGreaterThanOrEqual(0);
    expect(first("/owners")).toBeLessThan(firstPost);
    expect(first(`/applications/${AGENT_OBJECT_ID}?`)).toBeLessThan(firstPost);
    expect(first(`servicePrincipals(appId='${GRAPH}')`)).toBeLessThan(firstPost);
  });

  it("is idempotent: roles already held are not granted again", async () => {
    const { deps, events, graph } = await guardSetup(
      fakeGraph({ assigned: ["r-chat", "r-install"] }),
    );
    await assignAgentGraphAppRoles(deps, "bot_tony");
    expect(posts(graph.calls)).toHaveLength(0);
    expect(events).toHaveLength(0);
  });

  it.each([
    ["the provisioner", { entraAppId: PROVISIONER_CLIENT_ID }],
    ["the shared Bot app", { entraAppId: SHARED_BOT_APP_ID }],
    ["Microsoft Graph", { entraAppId: GRAPH }],
    ["a row with no Entra app", { entraAppId: null }],
    ["a published app", { status: "published" }],
    ["an app being removed", { status: "removing" }],
  ])("refuses %s before any Graph call", async (_label, row) => {
    const { deps, events, graph } = await guardSetup(fakeGraph(), row);
    await expect(assignAgentGraphAppRoles(deps, "bot_tony")).rejects.toBeInstanceOf(
      AppRoleGuardRefusal,
    );
    expect(graph.calls).toHaveLength(0);
    expect(events).toEqual([expect.objectContaining({ type: "teams.agent_app.role_refused" })]);
  });

  it("refuses a bot with no recorded agent app, before any Graph call", async () => {
    const { deps, events, graph } = await guardSetup();
    await expect(assignAgentGraphAppRoles(deps, "bot_other")).rejects.toThrow(/no agent app/);
    expect(graph.calls).toHaveLength(0);
    expect(events[0]?.type).toBe("teams.agent_app.role_refused");
  });

  it("refuses when the provisioner's object id is unknown", async () => {
    const { deps, graph } = await guardSetup();
    deps.identity = { clientId: PROVISIONER_CLIENT_ID, objectId: async () => "" };
    await expect(assignAgentGraphAppRoles(deps, "bot_tony")).rejects.toThrow(/object id/);
    expect(graph.calls).toHaveLength(0);
  });

  it.each([
    ["the provisioner does not own the app", fakeGraph({ owners: ["someone-else"] }), /owner/],
    ["the app lacks the tag", fakeGraph({ entraApp: { tags: [] } }), /tag/],
    [
      "the app is multi-tenant",
      fakeGraph({ entraApp: { signInAudience: "AzureADMultipleOrgs" } }),
      /single-tenant/,
    ],
    [
      "the app requests API access",
      fakeGraph({ entraApp: { requiredResourceAccess: [{ resourceAppId: GRAPH }] } }),
      /API access/,
    ],
    ["the app id differs from the record", fakeGraph({ entraApp: { appId: "other" } }), /match/],
    ["the service principal is missing", fakeGraph({ sp: null }), /not found/],
    [
      "the service principal is another app's",
      fakeGraph({ sp: { appId: "other" } }),
      /another app/,
    ],
    [
      "the service principal is the provisioner",
      fakeGraph({ sp: { id: PROVISIONER_OID } }),
      /provisioner/,
    ],
    [
      "the resource is not Graph",
      fakeGraph({ graphSp: graphSp({ appId: "spoofed" }) }),
      /Microsoft Graph/,
    ],
    [
      "Graph lacks an application role",
      fakeGraph({ graphSp: graphSp({ appRoles: [] }) }),
      /no application role/,
    ],
  ])("refuses when %s, granting nothing", async (_label, graph, message) => {
    const { deps, events } = await guardSetup(graph);
    await expect(assignAgentGraphAppRoles(deps, "bot_tony")).rejects.toThrow(message);
    expect(posts(graph.calls)).toHaveLength(0);
    expect(events).toHaveLength(1);
    expect(events[0]?.type).toBe("teams.agent_app.role_refused");
    expect(JSON.stringify(events)).not.toContain("secret-token");
  });
});

/** Every non-test TypeScript source under packages/ and apps/. */
function sources(dir: string): string[] {
  return readdirSync(dir).flatMap((entry) => {
    if (
      entry === "node_modules" ||
      entry === "dist" ||
      entry === "generated" ||
      entry.startsWith(".")
    ) {
      return [];
    }
    const file = path.join(dir, entry);
    if (statSync(file).isDirectory()) return sources(file);
    if (!/\.(?:ts|tsx|mts)$/.test(entry) || /\.d\.ts$/.test(entry)) return [];
    if (/\.(?:test|spec)\./.test(entry)) return [];
    return [file];
  });
}

describe("the guard is the only way to grant app roles", () => {
  const here = path.dirname(fileURLToPath(import.meta.url));
  const repoRoot = path.resolve(here, "../../../..");
  const files = ["packages", "apps", "infra"].flatMap((dir) => sources(path.join(repoRoot, dir)));
  const guardFile = path.join(here, "app-role-guard.ts");

  it("scans the repository's sources", () => {
    expect(files).toContain(guardFile);
    expect(files.length).toBeGreaterThan(100);
  });

  it("names the app role assignment endpoints in no other module", () => {
    const offenders = files.filter(
      (file) =>
        file !== guardFile && /appRoleAssign(?:edTo|ments)/i.test(readFileSync(file, "utf8")),
    );
    expect(offenders).toEqual([]);
  });

  it("never creates oauth2 (delegated) permission grants anywhere", () => {
    const offenders = files.filter((file) =>
      /oauth2PermissionGrants/i.test(readFileSync(file, "utf8")),
    );
    expect(offenders).toEqual([]);
  });

  it("has exactly one POST in the guard, after every check", () => {
    const text = readFileSync(guardFile, "utf8");
    expect(text.match(/method: "POST"/g)).toHaveLength(1);
    const post = text.indexOf('method: "POST"');
    for (const check of [
      "checkTargetAppId(",
      "checkRoleRequest({",
      "checkEntraAppShape(",
      "checkOwnership(",
      "checkServicePrincipal(",
      "resolveGraphRoles(",
    ]) {
      expect(text.lastIndexOf(check)).toBeLessThan(post);
    }
  });
});
