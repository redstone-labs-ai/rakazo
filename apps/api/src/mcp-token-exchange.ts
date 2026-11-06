import type { McpOAuthBroker, McpTokenExchangeConfig } from "@rakazo/adapters";
import { McpTokenExchange } from "@rakazo/adapters";
import type { OidcConfig } from "@rakazo/auth";
import { createOidcDelegation } from "@rakazo/auth";
import type { PrismaClient } from "@rakazo/db";
import { runAsSystem, runAsTenant } from "@rakazo/db";
import { getLogger } from "@rakazo/logging";

/** Secrets are read only from variables with this prefix, so the list cannot send others. */
export const MCP_TOKEN_EXCHANGE_SECRET_PREFIX = "RAKAZO_MCP_TOKEN_EXCHANGE_";

/**
 * RAKAZO_MCP_TOKEN_EXCHANGE: a JSON array of `{ issuer, tokenEndpoint, clientId,
 * clientSecretEnv, subjectScope, roleField? }`. Each client secret is read from the variable
 * `clientSecretEnv` names. An entry that does not parse is logged and left out.
 */
export function mcpTokenExchangeFromEnv(
  source: NodeJS.ProcessEnv,
  warn: (message: string) => void = (message) => getLogger().warn(message),
): McpTokenExchangeConfig[] {
  const raw = source.RAKAZO_MCP_TOKEN_EXCHANGE?.trim();
  if (!raw) return [];
  let list: unknown;
  try {
    list = JSON.parse(raw);
  } catch {
    warn("Ignoring RAKAZO_MCP_TOKEN_EXCHANGE: not valid JSON");
    return [];
  }
  if (!Array.isArray(list)) {
    warn("Ignoring RAKAZO_MCP_TOKEN_EXCHANGE: expected a JSON array");
    return [];
  }
  const configs: McpTokenExchangeConfig[] = [];
  for (const [index, item] of list.entries()) {
    const fields: Record<string, unknown> = item && typeof item === "object" ? item : {};
    const text = (key: string) =>
      typeof fields[key] === "string" && fields[key].trim() ? fields[key].trim() : undefined;
    const skip = (reason: string) =>
      warn(`Ignoring RAKAZO_MCP_TOKEN_EXCHANGE entry ${index}: ${reason}`);
    const issuer = httpsUrl(text("issuer"));
    const tokenEndpoint = httpsUrl(text("tokenEndpoint"));
    const clientId = text("clientId");
    const subjectScope = text("subjectScope");
    const secretEnv = text("clientSecretEnv");
    if (!issuer || !tokenEndpoint) {
      skip("issuer and tokenEndpoint must be HTTPS URLs");
      continue;
    }
    if (!clientId || !subjectScope || !secretEnv) {
      skip("clientId, clientSecretEnv and subjectScope are required");
      continue;
    }
    if (!secretEnv.startsWith(MCP_TOKEN_EXCHANGE_SECRET_PREFIX)) {
      skip(`clientSecretEnv must start with ${MCP_TOKEN_EXCHANGE_SECRET_PREFIX}`);
      continue;
    }
    const clientSecret = source[secretEnv]?.trim();
    if (!clientSecret) {
      skip(`${secretEnv} is not set`);
      continue;
    }
    const roleField = text("roleField");
    configs.push({
      issuer,
      tokenEndpoint,
      clientId,
      clientSecret,
      subjectScope,
      ...(roleField ? { roleField } : {}),
    });
  }
  return configs;
}

function httpsUrl(value: string | undefined): string | undefined {
  if (!value) return undefined;
  try {
    const url = new URL(value);
    return url.protocol === "https:" ? value : undefined;
  } catch {
    return undefined;
  }
}

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

/**
 * How long an issued grant is kept before it is issued again. Services commonly bind issued
 * grants to a day-long session they reuse while it has an hour left, so an hourly sweep that
 * renews after 23 hours lands in that last hour and starts a fresh session.
 */
export const ISSUED_GRANT_RENEW_AFTER_MS = 23 * 60 * 60_000;

/**
 * Keeps issued grants alive for people who never sign in again, such as those who only use a
 * chat app, since only the API can issue them. For each person: when their oldest issued grant
 * is due, or one was dropped (a worker cannot issue a new one), issue all of theirs again,
 * together so they share the new session. Returns how many servers were renewed.
 */
export async function renewIssuedMcpGrants(
  deps: { prisma: PrismaClient; mcpOAuth: McpOAuthBroker },
  signal?: AbortSignal,
  now = Date.now(),
): Promise<number> {
  const servers = await runAsSystem("renew issued MCP grants", () =>
    deps.prisma.mcpServer.findMany({
      where: { enabled: true, endpoint: { not: null }, secretId: { not: null } },
      select: { id: true, userId: true, spaceId: true, endpoint: true, secretId: true },
      orderBy: [{ userId: "asc" }, { createdAt: "asc" }],
    }),
  );
  const byUser = new Map<string, typeof servers>();
  for (const server of servers)
    byUser.set(server.userId, [...(byUser.get(server.userId) ?? []), server]);
  let renewed = 0;
  for (const [userId, owned] of byUser) {
    if (signal?.aborted) break;
    try {
      const issued = [];
      for (const server of owned) {
        const context = { spaceId: server.spaceId, userId };
        const grant = await runAsTenant(context, () => deps.mcpOAuth.issuedGrant(server, context));
        if (grant) issued.push({ server, context, grant });
      }
      const due = issued.some(
        ({ grant }) =>
          !grant.connected || (grant.issuedAt ?? 0) <= now - ISSUED_GRANT_RENEW_AFTER_MS,
      );
      if (!due) continue;
      for (const { server, context } of issued)
        if (await runAsTenant(context, () => deps.mcpOAuth.renewIssued(server, context)))
          renewed += 1;
    } catch (error) {
      getLogger().warn("issued MCP grants not renewed", {
        error: error instanceof Error ? error.message : String(error),
      });
    }
  }
  return renewed;
}
