import type { McpTokenExchangeConfig } from "@rakazo/adapters";
import { McpTokenExchange, mcpTokenExchangeFromEnv } from "@rakazo/adapters";
import { createOidcDelegation } from "@rakazo/auth/oidc-delegation";
import type { PrismaClient } from "@rakazo/db";
import { getLogger } from "@rakazo/logging";

/**
 * The API's token exchange, so a grant the worker finds expired or refused is issued again
 * during a run instead of asking the person to reconnect. Needs the API's OIDC settings and
 * auth secret; without them the worker leaves such grants to the next sign-in.
 */
export function workerMcpTokenExchange(
  prisma: PrismaClient,
  source: NodeJS.ProcessEnv,
  configs: McpTokenExchangeConfig[] = mcpTokenExchangeFromEnv(source),
): McpTokenExchange | undefined {
  if (!configs.length) return undefined;
  const issuer = source.OIDC_ISSUER?.trim();
  const clientId = source.OIDC_CLIENT_ID?.trim();
  const clientSecret = source.OIDC_CLIENT_SECRET?.trim();
  const authSecret = source.BETTER_AUTH_SECRET?.trim();
  if (!issuer || !clientId || !clientSecret || !authSecret) {
    getLogger().warn(
      "RAKAZO_MCP_TOKEN_EXCHANGE on the worker needs OIDC_ISSUER, OIDC_CLIENT_ID, OIDC_CLIENT_SECRET and BETTER_AUTH_SECRET; runs ask to reconnect instead",
    );
    return undefined;
  }
  const delegation = createOidcDelegation(prisma, { issuer, clientId, clientSecret, authSecret });
  return new McpTokenExchange(configs, {
    subjectToken: (userId, scope) => delegation.accessToken(userId, scope),
  });
}
