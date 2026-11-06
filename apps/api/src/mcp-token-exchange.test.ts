import { describe, expect, it, vi } from "vitest";
import { mcpTokenExchangeFromEnv, renewIssuedMcpGrants } from "./mcp-token-exchange.js";

const entry = {
  issuer: "https://auth.example.test/api/auth",
  tokenEndpoint: "https://auth.example.test/api/auth/oauth2/token",
  clientId: "deployment-client",
  clientSecretEnv: "RAKAZO_MCP_TOKEN_EXCHANGE_SECRET_SERVICE",
  subjectScope: "api://service/MCP.Access",
};

describe("mcpTokenExchangeFromEnv", () => {
  it("reads each client secret from the variable the entry names", () => {
    expect(
      mcpTokenExchangeFromEnv({
        RAKAZO_MCP_TOKEN_EXCHANGE: JSON.stringify([{ ...entry, roleField: "service_role" }]),
        RAKAZO_MCP_TOKEN_EXCHANGE_SECRET_SERVICE: " secret ",
      }),
    ).toEqual([
      {
        issuer: entry.issuer,
        tokenEndpoint: entry.tokenEndpoint,
        clientId: entry.clientId,
        clientSecret: "secret",
        subjectScope: entry.subjectScope,
        roleField: "service_role",
      },
    ]);
  });

  it("logs and leaves out what cannot work", () => {
    const warn = vi.fn();
    const configs = mcpTokenExchangeFromEnv(
      {
        RAKAZO_MCP_TOKEN_EXCHANGE: JSON.stringify([
          entry,
          { ...entry, clientSecretEnv: "BETTER_AUTH_SECRET" },
          { ...entry, tokenEndpoint: "http://auth.example.test/token" },
          { ...entry, subjectScope: undefined },
          { ...entry, clientSecretEnv: "RAKAZO_MCP_TOKEN_EXCHANGE_SECRET_MISSING" },
        ]),
        RAKAZO_MCP_TOKEN_EXCHANGE_SECRET_SERVICE: "secret",
        BETTER_AUTH_SECRET: "never-sent",
      },
      warn,
    );
    expect(configs).toHaveLength(1);
    expect(warn).toHaveBeenCalledTimes(4);
    expect(warn).toHaveBeenCalledWith(
      expect.stringContaining("must start with RAKAZO_MCP_TOKEN_EXCHANGE_"),
    );
    expect(mcpTokenExchangeFromEnv({ RAKAZO_MCP_TOKEN_EXCHANGE: "{" }, warn)).toEqual([]);
    expect(mcpTokenExchangeFromEnv({}, warn)).toEqual([]);
  });
});

describe("renewIssuedMcpGrants", () => {
  const hour = 60 * 60_000;
  const now = Date.UTC(2026, 10, 6, 12);
  const server = (id: string, userId: string) => ({
    id,
    userId,
    spaceId: `space-${userId}`,
    endpoint: `https://${id}.example.test/mcp`,
    secretId: `secret-${id}`,
  });

  function setup(grants: Record<string, { issuedAt?: number; connected: boolean } | undefined>) {
    const servers = [
      server("a1", "ann"),
      server("a2", "ann"),
      server("b1", "bob"),
      server("c1", "cy"),
    ];
    const prisma = { mcpServer: { findMany: vi.fn(async () => servers) } };
    const mcpOAuth = {
      issuedGrant: vi.fn(async (row: { id: string }) => grants[row.id]),
      renewIssued: vi.fn(async (_row: { id: string }) => true),
    };
    return { deps: { prisma, mcpOAuth } as never, mcpOAuth };
  }

  it("renews all of a person's issued grants once the oldest is due or one was dropped", async () => {
    const { deps, mcpOAuth } = setup({
      // Ann's oldest grant is in the last hour of its day: renew both, so they share a session.
      a1: { issuedAt: now - 23.5 * hour, connected: true },
      a2: { issuedAt: now - 2 * hour, connected: true },
      // Bob's was dropped by a process that cannot issue grants.
      b1: { issuedAt: now - hour, connected: false },
      // Cy's server was not issued by the exchange.
      c1: undefined,
    });

    expect(await renewIssuedMcpGrants(deps, undefined, now)).toBe(3);
    expect(mcpOAuth.renewIssued.mock.calls.map(([row]) => row.id)).toEqual(["a1", "a2", "b1"]);
  });

  it("leaves fresh grants alone", async () => {
    const { deps, mcpOAuth } = setup({
      a1: { issuedAt: now - 22 * hour, connected: true },
      b1: { issuedAt: now - hour, connected: true },
    });
    expect(await renewIssuedMcpGrants(deps, undefined, now)).toBe(0);
    expect(mcpOAuth.renewIssued).not.toHaveBeenCalled();
  });
});
