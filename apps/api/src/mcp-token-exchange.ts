import type { McpOAuthBroker, McpTokenExchangeConfig } from "@rakazo/adapters";
import { McpTokenExchange } from "@rakazo/adapters";
import type { OidcConfig } from "@rakazo/auth";
import { createOidcDelegation } from "@rakazo/auth/oidc-delegation";
import type { PrismaClient } from "@rakazo/db";
import { runAsSystem, runAsTenant } from "@rakazo/db";
import { getLogger } from "@rakazo/logging";

/**
 * The token exchange a deployment configured, issuing grants for the SSO access token the
 * person's stored refresh token yields. Undefined without both.
 */
export function createMcpTokenExchange(
  prisma: PrismaClient,
  configs: McpTokenExchangeConfig[],
  oidc: Pick<OidcConfig, "issuer" | "clientId" | "clientSecret" | "scopes"> | undefined,
  authSecret: string,
): McpTokenExchange | undefined {
  if (!configs.length) return undefined;
  if (!oidc) {
    getLogger().warn("RAKAZO_MCP_TOKEN_EXCHANGE needs OIDC sign-in; it is not used");
    return undefined;
  }
  if (!oidc.scopes.includes("offline_access"))
    getLogger().warn(
      "RAKAZO_MCP_TOKEN_EXCHANGE needs offline_access in OIDC_SCOPES to keep a refresh token",
    );
  const delegation = createOidcDelegation(prisma, { ...oidc, authSecret });
  return new McpTokenExchange(configs, {
    subjectToken: (userId, scope) => delegation.accessToken(userId, scope),
    onRole: async (userId, issuer, role) => {
      await prisma.serviceRole.upsert({
        where: { userId_issuer: { userId, issuer } },
        create: { userId, issuer, role },
        update: { role },
      });
    },
  });
}

/** Roles services reported for the person when they last issued them a grant. */
export async function serviceRoles(prisma: PrismaClient, userId: string): Promise<string[]> {
  const rows = await prisma.serviceRole.findMany({ where: { userId }, select: { role: true } });
  return [...new Set(rows.map((row) => row.role))];
}

/**
 * Records the directory object the sign-in's ID token names (`tid` and `oid`, lowercased).
 * The newest sign-in owns it, so it never names two users.
 */
export async function recordDirectoryIdentity(
  prisma: PrismaClient,
  userId: string,
  claims: Record<string, unknown> | undefined,
): Promise<void> {
  if (typeof claims?.tid !== "string" || typeof claims.oid !== "string") return;
  // Directory ids are GUIDs, which chat apps may report in either case.
  const tenantId = claims.tid.toLowerCase();
  const objectId = claims.oid.toLowerCase();
  if (!tenantId || !objectId) return;
  await prisma.$transaction(async (tx) => {
    await tx.directoryIdentity.deleteMany({
      where: { tenantId, objectId, userId: { not: userId } },
    });
    await tx.directoryIdentity.upsert({
      where: { userId },
      create: { userId, tenantId, objectId },
      update: { tenantId, objectId },
    });
  });
}

/**
 * Connects each of the person's remote MCP servers that is not connected and whose
 * authorization server the exchange holds a client for. One at a time, since every request
 * goes to the same few authorization servers. Returns the ids connected.
 */
export async function connectIssuedMcpServers(
  deps: { prisma: PrismaClient; mcpOAuth: McpOAuthBroker },
  userId: string,
): Promise<string[]> {
  const servers = await runAsSystem("connect a person's MCP servers after they sign in", () =>
    deps.prisma.mcpServer.findMany({
      where: { userId, enabled: true, endpoint: { not: null } },
      select: { id: true, spaceId: true, endpoint: true, secretId: true },
      orderBy: { createdAt: "asc" },
    }),
  );
  const connected: string[] = [];
  for (const server of servers) {
    const context = { spaceId: server.spaceId, userId };
    try {
      if (await runAsTenant(context, () => deps.mcpOAuth.connectIssued(server, context)))
        connected.push(server.id);
    } catch (error) {
      getLogger().warn("MCP server not connected by token exchange", {
        "mcp.server.id": server.id,
        error: error instanceof Error ? error.message : String(error),
      });
    }
  }
  return connected;
}

/**
 * After a sign-in: connect what the exchange can, then add the default servers a role it just
 * reported admits, and connect those too.
 */
export async function connectIssuedAfterSignIn(
  deps: { prisma: PrismaClient; mcpOAuth: McpOAuthBroker },
  userId: string,
  provisionDefaults: () => Promise<string[]>,
): Promise<void> {
  await connectIssuedMcpServers(deps, userId);
  if ((await provisionDefaults()).length) await connectIssuedMcpServers(deps, userId);
}
