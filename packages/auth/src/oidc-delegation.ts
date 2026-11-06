import type { PrismaClient } from "@rakazo/db";
import { symmetricDecrypt, symmetricEncrypt } from "better-auth/crypto";
import { isCurrentOidcAccount } from "./oidc.js";

export type OidcDelegationConfig = {
  issuer: string;
  clientId: string;
  clientSecret: string;
  /** Better Auth's secret, which encrypts the stored provider tokens. */
  authSecret: string;
};

const REQUEST_TIMEOUT_MS = 15_000;
/** Better Auth's ciphertext; a token stored before encryption was enabled is plain. */
const isEncrypted = (value: string) =>
  value.startsWith("$ba$") || (value.length % 2 === 0 && /^[0-9a-f]+$/i.test(value));

/**
 * Access tokens for other scopes of the person's SSO provider, from the refresh token their
 * last SSO sign-in stored. That needs the `offline_access` scope at sign-in. Runs without a
 * Better Auth instance, so the worker can use it too.
 */
export function createOidcDelegation(
  prisma: Pick<PrismaClient, "account">,
  config: OidcDelegationConfig,
  fetchFn: typeof fetch = fetch,
) {
  let tokenEndpoint: Promise<string> | undefined;
  const discover = async () => {
    const response = await fetchFn(
      `${config.issuer.replace(/\/$/, "")}/.well-known/openid-configuration`,
      { signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS), redirect: "error" },
    );
    const metadata = (await response.json().catch(() => undefined)) as
      | { issuer?: unknown; token_endpoint?: unknown }
      | undefined;
    if (
      !response.ok ||
      metadata?.issuer !== config.issuer ||
      typeof metadata.token_endpoint !== "string"
    )
      throw new Error("SSO provider discovery failed");
    return metadata.token_endpoint;
  };
  return {
    /** Undefined when the person has no usable SSO refresh token. Throws when the provider fails. */
    async accessToken(userId: string, scope: string): Promise<string | undefined> {
      const accounts = await prisma.account.findMany({
        where: { userId, providerId: "oidc", refreshToken: { not: null } },
        orderBy: { updatedAt: "desc" },
      });
      let account: (typeof accounts)[number] | undefined;
      for (const candidate of accounts) {
        if (await isCurrentOidcAccount(candidate, config.issuer, config.authSecret)) {
          account = candidate;
          break;
        }
      }
      const stored = account?.refreshToken;
      if (!account || !stored) return undefined;
      const refreshToken = isEncrypted(stored)
        ? await symmetricDecrypt({ key: config.authSecret, data: stored })
        : stored;
      tokenEndpoint ??= discover().catch((error: unknown) => {
        tokenEndpoint = undefined;
        throw error;
      });
      const response = await fetchFn(await tokenEndpoint, {
        method: "POST",
        headers: {
          "content-type": "application/x-www-form-urlencoded",
          accept: "application/json",
        },
        body: new URLSearchParams({
          grant_type: "refresh_token",
          refresh_token: refreshToken,
          client_id: config.clientId,
          client_secret: config.clientSecret,
          scope,
        }),
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
        redirect: "error",
      });
      const body = (await response.json().catch(() => ({}))) as Record<string, unknown>;
      if (!response.ok || typeof body.access_token !== "string") {
        const code = typeof body.error === "string" ? body.error : `HTTP ${response.status}`;
        throw new Error(`SSO provider refused a delegated token: ${code}`);
      }
      // Keep the newest refresh token, unless a sign-in stored another one meanwhile.
      if (typeof body.refresh_token === "string" && body.refresh_token !== refreshToken)
        await prisma.account.updateMany({
          where: { id: account.id, refreshToken: stored },
          data: {
            refreshToken: await symmetricEncrypt({
              key: config.authSecret,
              data: body.refresh_token,
            }),
          },
        });
      return body.access_token;
    },
  };
}
