import type { PrismaClient } from "@rakazo/db";
import { describe, expect, it, vi } from "vitest";
import { refreshTokenCatalogToken, teamsAgentAppsFromEnv } from "./env.js";
import type { FetchLike } from "./http.js";
import { managedIdentityTokens, scopeResource, tokenObjectId } from "./http.js";

const ENDPOINT = "http://localhost:42356/msi/token";

function jwt(claims: Record<string, unknown>): string {
  const part = (value: unknown) => Buffer.from(JSON.stringify(value)).toString("base64url");
  return `${part({ alg: "none" })}.${part(claims)}.sig`;
}

function identityEndpoint(options: { expiresOn?: (now: number) => number; status?: number } = {}) {
  const calls: { url: URL; header: string | null }[] = [];
  let clock = Date.parse("2026-11-11T12:00:00Z");
  let issued = 0;
  const fetchImpl: FetchLike = async (url, init) => {
    calls.push({ url: new URL(url), header: new Headers(init?.headers).get("x-identity-header") });
    if (options.status && options.status !== 200) {
      return new Response(JSON.stringify({ error: "invalid_request" }), { status: options.status });
    }
    issued += 1;
    const expiresOn = options.expiresOn?.(clock) ?? Math.floor(clock / 1000) + 3600;
    return new Response(
      JSON.stringify({
        access_token: jwt({ oid: "mi-oid", n: issued }),
        expires_on: String(expiresOn),
        resource: new URL(url).searchParams.get("resource"),
      }),
      { status: 200 },
    );
  };
  return {
    fetch: vi.fn(fetchImpl),
    calls,
    now: () => clock,
    advance: (ms: number) => {
      clock += ms;
    },
  };
}

describe("managedIdentityTokens", () => {
  it("asks IDENTITY_ENDPOINT with the header, api version, resource and client id", async () => {
    const endpoint = identityEndpoint();
    const tokens = managedIdentityTokens({
      endpoint: ENDPOINT,
      header: "header-value",
      clientId: "mi-client",
      fetch: endpoint.fetch,
      now: endpoint.now,
    });
    await tokens("https://graph.microsoft.com/.default");
    const [call] = endpoint.calls;
    expect(call?.header).toBe("header-value");
    expect(`${call?.url.origin}${call?.url.pathname}`).toBe(ENDPOINT);
    expect(Object.fromEntries(call?.url.searchParams ?? [])).toEqual({
      "api-version": "2019-08-01",
      resource: "https://graph.microsoft.com",
      client_id: "mi-client",
    });
  });

  it.each([
    ["https://graph.microsoft.com/.default", "https://graph.microsoft.com"],
    ["https://management.azure.com/.default", "https://management.azure.com"],
    ["https://vault.azure.net/.default", "https://vault.azure.net"],
  ])("serves %s", async (scope, resource) => {
    const endpoint = identityEndpoint();
    const tokens = managedIdentityTokens({
      endpoint: ENDPOINT,
      header: "h",
      clientId: "c",
      fetch: endpoint.fetch,
    });
    await tokens(scope);
    expect(endpoint.calls[0]?.url.searchParams.get("resource")).toBe(resource);
    expect(scopeResource(scope)).toBe(resource);
  });

  it("refuses any other resource without calling the endpoint", async () => {
    const endpoint = identityEndpoint();
    const tokens = managedIdentityTokens({
      endpoint: ENDPOINT,
      header: "h",
      clientId: "c",
      fetch: endpoint.fetch,
    });
    await expect(tokens("https://api.botframework.com/.default")).rejects.toThrow(/not issued/);
    await expect(tokens("https://storage.azure.com/.default")).rejects.toThrow(/not issued/);
    expect(endpoint.fetch).not.toHaveBeenCalled();
  });

  it("caches per resource until five minutes before expiry", async () => {
    const endpoint = identityEndpoint();
    const tokens = managedIdentityTokens({
      endpoint: ENDPOINT,
      header: "h",
      clientId: "c",
      fetch: endpoint.fetch,
      now: endpoint.now,
    });
    const first = await tokens("https://graph.microsoft.com/.default");
    expect(await tokens("https://graph.microsoft.com/.default")).toBe(first);
    await tokens("https://vault.azure.net/.default");
    expect(endpoint.fetch).toHaveBeenCalledTimes(2);

    endpoint.advance(54 * 60_000); // 6 minutes left: still cached
    expect(await tokens("https://graph.microsoft.com/.default")).toBe(first);
    expect(endpoint.fetch).toHaveBeenCalledTimes(2);

    endpoint.advance(2 * 60_000); // 4 minutes left: renewed
    expect(await tokens("https://graph.microsoft.com/.default")).not.toBe(first);
    expect(endpoint.fetch).toHaveBeenCalledTimes(3);
  });

  it("caches nothing on failure and reports no token", async () => {
    const endpoint = identityEndpoint({ status: 400 });
    const tokens = managedIdentityTokens({
      endpoint: ENDPOINT,
      header: "h",
      clientId: "c",
      fetch: endpoint.fetch,
    });
    await expect(tokens("https://graph.microsoft.com/.default")).rejects.toThrow(
      /managed identity token request failed \(400 invalid_request\)/,
    );
    await expect(tokens("https://graph.microsoft.com/.default")).rejects.toThrow();
    expect(endpoint.fetch).toHaveBeenCalledTimes(2);
  });
});

describe("tokenObjectId", () => {
  it("reads the oid claim", () => {
    expect(tokenObjectId(jwt({ oid: "abc" }))).toBe("abc");
  });
  it("refuses tokens without one", () => {
    expect(() => tokenObjectId("opaque")).toThrow(/JWT/);
    expect(() => tokenObjectId(jwt({ sub: "x" }))).toThrow(/oid/);
    expect(() => tokenObjectId("a.!!!.c")).toThrow(/oid/);
  });
});

describe("refreshTokenCatalogToken", () => {
  it("refreshes as a public client: no client secret is sent", async () => {
    const bodies: URLSearchParams[] = [];
    const fetchImpl: FetchLike = async (_url, init) => {
      bodies.push(new URLSearchParams(String(init?.body)));
      return new Response(JSON.stringify({ access_token: "admin", refresh_token: "rt-2" }), {
        status: 200,
      });
    };
    const values = new Map([["catalog-rt", "rt-1"]]);
    const token = refreshTokenCatalogToken({
      tenantId: "t",
      clientId: "catalog-client",
      refreshTokenSecret: "catalog-rt",
      secrets: {
        get: async (name) => values.get(name) ?? "",
        set: async (name, value) => {
          values.set(name, value);
        },
        delete: async () => undefined,
      },
      fetch: fetchImpl,
    });
    expect(await token()).toBe("admin");
    expect(bodies[0]?.get("client_id")).toBe("catalog-client");
    expect(bodies[0]?.has("client_secret")).toBe(false);
    expect(values.get("catalog-rt")).toBe("rt-2");
  });
});

describe("teamsAgentAppsFromEnv", () => {
  const prisma = {} as unknown as PrismaClient;
  const env = {
    TEAMS_AGENT_APPS_ENABLED: "true",
    TEAMS_APP_TENANT_ID: "tenant",
    TEAMS_AGENT_APPS_IDENTITY_CLIENT_ID: "mi-client",
    IDENTITY_ENDPOINT: ENDPOINT,
    IDENTITY_HEADER: "header-value",
    TEAMS_AGENT_APPS_SUBSCRIPTION_ID: "sub",
    TEAMS_AGENT_APPS_RESOURCE_GROUP: "rg",
    TEAMS_AGENT_APPS_KEY_VAULT_URL: "https://vault.example.test/",
    TEAMS_AGENT_APPS_CATALOG_CLIENT_ID: "catalog-client",
    TEAMS_AGENT_APPS_CATALOG_REFRESH_TOKEN_SECRET: "catalog-rt",
    TEAMS_MESSAGING_ENDPOINT: "https://api.example.test/teams",
    TEAMS_APP_ID: "SHARED-APP",
    TEAMS_APP_PASSWORD: "shared-pw",
    TEAMS_APP_CATALOG_ID: "shared-catalog",
  };

  it("is off unless TEAMS_AGENT_APPS_ENABLED is exactly true", () => {
    expect(
      teamsAgentAppsFromEnv({ ...env, TEAMS_AGENT_APPS_ENABLED: "1" }, prisma),
    ).toBeUndefined();
    expect(
      teamsAgentAppsFromEnv({ ...env, TEAMS_AGENT_APPS_ENABLED: undefined }, prisma),
    ).toBeUndefined();
  });

  it("needs no provisioner client id or secret", () => {
    expect(teamsAgentAppsFromEnv(env, prisma)).toBeDefined();
  });

  it.each(["TEAMS_AGENT_APPS_IDENTITY_CLIENT_ID", "IDENTITY_ENDPOINT", "IDENTITY_HEADER"])(
    "is off without %s",
    (key) => {
      expect(teamsAgentAppsFromEnv({ ...env, [key]: " " }, prisma)).toBeUndefined();
    },
  );

  it("authenticates the provisioner as the managed identity and learns its object id", async () => {
    const endpoint = identityEndpoint();
    const deps = teamsAgentAppsFromEnv(env, prisma, { fetch: endpoint.fetch });
    expect(deps?.identity.clientId).toBe("mi-client");
    expect(deps?.config.sharedBotAppId).toBe("shared-app");
    expect(await deps?.identity.objectId()).toBe("mi-oid");
    expect(endpoint.calls[0]?.header).toBe("header-value");
    expect(endpoint.calls[0]?.url.searchParams.get("client_id")).toBe("mi-client");
  });
});
