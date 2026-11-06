import { symmetricDecrypt, symmetricEncrypt } from "better-auth/crypto";
import { describe, expect, it, vi } from "vitest";
import { oidcAccountSubject } from "./oidc.js";
import { createOidcDelegation } from "./oidc-delegation.js";

const issuer = "https://identity.example.test";
const authSecret = "offline-test-secret-with-at-least-32-characters";
const config = { issuer, clientId: "client", clientSecret: "client-secret", authSecret };

const idToken = (claims: Record<string, unknown>) =>
  `e30.${Buffer.from(JSON.stringify(claims)).toString("base64url")}.sig`;

async function account(overrides: Record<string, unknown> = {}) {
  return {
    id: "account-1",
    providerId: "oidc",
    accountId: oidcAccountSubject(issuer, "subject-1"),
    idToken: idToken({ iss: issuer, sub: "subject-1" }),
    refreshToken: await symmetricEncrypt({ key: authSecret, data: "refresh-1" }),
    ...overrides,
  };
}

function provider(token: Record<string, unknown>, status = 200) {
  return vi.fn(async (input: string | URL | Request, _init?: RequestInit) => {
    const url = String(input);
    if (url.endsWith("/.well-known/openid-configuration"))
      return Response.json({ issuer, token_endpoint: `${issuer}/token` });
    return Response.json(token, { status });
  });
}

describe("createOidcDelegation", () => {
  it("redeems the stored refresh token for the scope and keeps the rotated one", async () => {
    const stored = await account();
    const prisma = {
      account: {
        findMany: vi.fn(async () => [stored]),
        updateMany: vi.fn(async (_args: { where: unknown; data: { refreshToken: string } }) => ({
          count: 1,
        })),
      },
    };
    const fetch = provider({ access_token: "delegated", refresh_token: "refresh-2" });
    const delegation = createOidcDelegation(prisma as never, config, fetch);

    expect(await delegation.accessToken("user-1", "api://service/MCP.Access")).toBe("delegated");

    const [, init] = fetch.mock.calls[1]!;
    expect(Object.fromEntries(new URLSearchParams(String(init?.body)))).toEqual({
      grant_type: "refresh_token",
      refresh_token: "refresh-1",
      client_id: "client",
      client_secret: "client-secret",
      scope: "api://service/MCP.Access",
    });
    const update = prisma.account.updateMany.mock.calls[0]![0];
    expect(update.where).toEqual({ id: "account-1", refreshToken: stored.refreshToken });
    expect(await symmetricDecrypt({ key: authSecret, data: update.data.refreshToken })).toBe(
      "refresh-2",
    );
    // Discovery is cached.
    await delegation.accessToken("user-1", "api://service/MCP.Access");
    expect(fetch).toHaveBeenCalledTimes(3);
  });

  it("has nothing without a refresh token from this issuer's sign-in", async () => {
    const fetch = provider({ access_token: "delegated" });
    const prisma = (rows: unknown[]) => ({
      account: { findMany: vi.fn(async () => rows), updateMany: vi.fn() },
    });
    const delegation = (rows: unknown[]) =>
      createOidcDelegation(prisma(rows) as never, config, fetch);
    // Signing in without offline_access stores no refresh token.
    expect(await delegation([]).accessToken("user-1", "scope")).toBeUndefined();
    const replaced = await account({
      idToken: idToken({ iss: "https://old-issuer.example.test", sub: "subject-1" }),
    });
    expect(await delegation([replaced]).accessToken("user-1", "scope")).toBeUndefined();
    expect(fetch).not.toHaveBeenCalled();
  });

  it("reports the provider's refusal by its error code", async () => {
    const prisma = { account: { findMany: vi.fn(async () => [await account()]) } };
    const delegation = createOidcDelegation(
      prisma as never,
      config,
      provider({ error: "invalid_grant", error_description: "AADSTS700082" }, 400),
    );
    await expect(delegation.accessToken("user-1", "scope")).rejects.toThrow(
      "SSO provider refused a delegated token: invalid_grant",
    );
  });
});
