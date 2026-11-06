import { describe, expect, it, vi } from "vitest";
import type { McpTokenExchangeConfig } from "./mcp-token-exchange.js";
import { McpTokenExchange } from "./mcp-token-exchange.js";

const entry = {
  issuer: "https://auth.example.test/api/auth",
  tokenEndpoint: "https://auth.example.test/api/auth/oauth2/token",
  clientId: "deployment-client",
  clientSecretEnv: "RAKAZO_MCP_TOKEN_EXCHANGE_SECRET_SERVICE",
  subjectScope: "api://service/MCP.Access",
};

describe("McpTokenExchange", () => {
  const config: McpTokenExchangeConfig = {
    issuer: entry.issuer,
    tokenEndpoint: entry.tokenEndpoint,
    clientId: entry.clientId,
    clientSecret: "secret",
    subjectScope: entry.subjectScope,
    roleField: "service_role",
  };
  const granted = () =>
    Response.json({
      access_token: "issued",
      issued_token_type: "urn:ietf:params:oauth:token-type:access_token",
      token_type: "Bearer",
      expires_in: 300,
      refresh_token: "issued-refresh",
      scope: "offline_access read",
      service_role: "staff",
    });

  it("holds a client only for its own issuer, however the issuer is spelled", () => {
    const exchange = new McpTokenExchange([config], { subjectToken: async () => "sso" });
    expect(exchange.clientFor(`${entry.issuer}/`)).toEqual({
      clientId: "deployment-client",
      clientSecret: "secret",
      authMethod: "client_secret_post",
    });
    expect(exchange.clientFor("https://other.example.test")).toBeUndefined();
  });

  it("exchanges the person's SSO token for one resource and reports their role", async () => {
    const fetch = vi.fn(async (_input: string | URL | Request, _init?: RequestInit) => granted());
    const subjectToken = vi.fn(async () => "sso-token");
    const onRole = vi.fn(async () => undefined);
    const exchange = new McpTokenExchange([config], { subjectToken, onRole, fetch });

    const tokens = await exchange.issue({
      userId: "user-1",
      authorizationServerUrl: entry.issuer,
      resource: "https://mcp.example.test/mcp",
    });

    expect(tokens).toEqual({
      access_token: "issued",
      token_type: "Bearer",
      expires_in: 300,
      refresh_token: "issued-refresh",
      scope: "offline_access read",
    });
    expect(subjectToken).toHaveBeenCalledWith("user-1", "api://service/MCP.Access");
    expect(onRole).toHaveBeenCalledWith("user-1", entry.issuer, "staff");
    const [url, init] = fetch.mock.calls[0]!;
    expect(url).toBe(entry.tokenEndpoint);
    // One resource, no scope (the server's defaults include a refresh token), no audience.
    expect(Object.fromEntries(new URLSearchParams(String(init?.body)))).toEqual({
      grant_type: "urn:ietf:params:oauth:grant-type:token-exchange",
      subject_token: "sso-token",
      subject_token_type: "urn:ietf:params:oauth:token-type:access_token",
      client_id: "deployment-client",
      client_secret: "secret",
      resource: "https://mcp.example.test/mcp",
    });
  });

  it("issues nothing without an SSO token or a client for the issuer", async () => {
    const fetch = vi.fn(async () => granted());
    const none = new McpTokenExchange([config], { subjectToken: async () => undefined, fetch });
    expect(await none.issue({ userId: "u", authorizationServerUrl: entry.issuer })).toBe(undefined);
    expect(
      await none.issue({ userId: "u", authorizationServerUrl: "https://other.example.test" }),
    ).toBe(undefined);
    expect(fetch).not.toHaveBeenCalled();
  });

  it("waits out a rate limit, and reports a refusal by its protocol error", async () => {
    const sleep = vi.fn(async () => undefined);
    const fetch = vi
      .fn()
      .mockResolvedValueOnce(new Response("", { status: 429, headers: { "retry-after": "3" } }))
      .mockResolvedValueOnce(granted())
      .mockResolvedValueOnce(
        Response.json(
          { error: "invalid_grant", error_description: "unknown user" },
          { status: 400 },
        ),
      );
    const exchange = new McpTokenExchange([config], {
      subjectToken: async () => "sso",
      fetch,
      sleep,
    });
    const input = { userId: "u", authorizationServerUrl: entry.issuer };
    expect((await exchange.issue(input))?.access_token).toBe("issued");
    expect(sleep).toHaveBeenCalledWith(3000);
    await expect(exchange.issue(input)).rejects.toThrow(
      "MCP token exchange refused: invalid_grant: unknown user",
    );
  });

  it("sends requests to one token endpoint one at a time", async () => {
    let active = 0;
    let most = 0;
    const fetch = vi.fn(async () => {
      active += 1;
      most = Math.max(most, active);
      await new Promise((resolve) => setTimeout(resolve, 5));
      active -= 1;
      return granted();
    });
    const exchange = new McpTokenExchange([config], { subjectToken: async () => "sso", fetch });
    await Promise.all(
      [1, 2, 3].map(() => exchange.issue({ userId: "u", authorizationServerUrl: entry.issuer })),
    );
    expect(fetch).toHaveBeenCalledTimes(3);
    expect(most).toBe(1);
  });
});
