import type { SecretStore } from "@rakazo/adapter-kit";
import { createAdapterContext } from "@rakazo/adapter-kit";
import { buildMcpCredentialBlob, deleteSecretBestEffort } from "@rakazo/adapters";
import type { McpServerConfigInput as McpServerConfigInputType } from "@rakazo/contracts";
import { McpServerConfigInput } from "@rakazo/contracts";
import type { PrismaClient } from "@rakazo/db";
import { availableInSpace, runAsTenant, saveApprovalRule } from "@rakazo/db";
import { getLogger } from "@rakazo/logging";
import { insertMcpServer } from "./mcp-servers.js";

/** Who an entry is for: anyone holding one of the roles. With no audience, everyone gets it. */
export type DefaultMcpServerAudience = {
  /** Roles the deployment reports for the person (`DefaultMcpServerRoles.roles`). */
  roles?: string[];
  /** Values of the `roles` claim in the person's SSO ID token. */
  oidcRoles?: string[];
};

export type DefaultMcpServer = {
  input: McpServerConfigInputType;
  shareWithAllBots: boolean;
  audience?: DefaultMcpServerAudience;
  /** Tool names on the server that start as "Always allow" for each person it is added for. */
  alwaysAllowTools?: string[];
};

/** An MCP tool name as the server lists it. "__" would read as another server's tool. */
const MCP_TOOL_NAME = /^[A-Za-z0-9](?:[A-Za-z0-9.-]|_(?!_))*$/;

/** The roles an audience is matched against. */
export type DefaultMcpServerRoles = { roles: string[]; oidcRoles: string[] };

/**
 * RAKAZO_DEFAULT_MCP_SERVERS: a JSON array of MCP servers each person gets when they sign in,
 * or `{ "servers": [...] }`. An entry that does not parse is logged and left out.
 */
export function defaultMcpServersFromEnv(
  raw: string | undefined,
  warn: (message: string) => void = (message) => getLogger().warn(message),
): DefaultMcpServer[] {
  const value = raw?.trim();
  if (!value) return [];
  let json: unknown;
  try {
    json = JSON.parse(value);
  } catch {
    warn("Ignoring RAKAZO_DEFAULT_MCP_SERVERS: not valid JSON");
    return [];
  }
  const list = isRecord(json) && "servers" in json ? json.servers : json;
  if (!Array.isArray(list)) {
    warn("Ignoring RAKAZO_DEFAULT_MCP_SERVERS: expected a JSON array");
    return [];
  }
  const servers: DefaultMcpServer[] = [];
  const slugs = new Set<string>();
  for (const [index, item] of list.entries()) {
    const skip = (reason: string) =>
      warn(`Ignoring RAKAZO_DEFAULT_MCP_SERVERS entry ${index}: ${reason}`);
    const fields = isRecord(item) ? item : {};
    if (fields.enabled === false) continue;
    const client = isRecord(fields.oauthClient) ? fields.oauthClient : undefined;
    const input = McpServerConfigInput.safeParse({
      slug: fields.slug,
      name: fields.name,
      description: fields.description,
      enabled: true,
      transport: fields.transport,
      endpoint: fields.endpoint,
      // Only a client id: a client secret would be copied into every person's server.
      ...(client
        ? {
            oauthClient: {
              clientId: client.clientId,
              scope: client.scope,
              authorizationParams: client.authorizationParams,
            },
          }
        : {}),
    });
    if (!input.success) {
      const issue = input.error.issues[0];
      skip(issue ? `${issue.path.join(".") || "value"}: ${issue.message}` : "invalid");
      continue;
    }
    if (input.data.transport === "stdio") {
      skip("only remote servers can be defaults");
      continue;
    }
    const audience = parseAudience(fields.audience);
    if (audience === null) {
      skip("audience must be an object of string arrays");
      continue;
    }
    const alwaysAllowTools = parseToolNames(fields.alwaysAllowTools);
    if (alwaysAllowTools === null) {
      skip("alwaysAllowTools must be an array of tool names");
      continue;
    }
    if (slugs.has(input.data.slug)) {
      skip(`duplicate slug ${input.data.slug}`);
      continue;
    }
    slugs.add(input.data.slug);
    servers.push({
      input: input.data,
      shareWithAllBots: fields.shareWithAllBots !== false,
      ...(audience ? { audience } : {}),
      ...(alwaysAllowTools?.length ? { alwaysAllowTools } : {}),
    });
  }
  return servers;
}

/** Undefined when absent, null when malformed. */
function parseToolNames(value: unknown): string[] | undefined | null {
  if (value === undefined) return undefined;
  if (!Array.isArray(value)) return null;
  const names = value.filter((name): name is string => typeof name === "string");
  if (names.length !== value.length || !names.every((name) => MCP_TOOL_NAME.test(name)))
    return null;
  return [...new Set(names)];
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

/** Undefined when absent, null when malformed. */
function parseAudience(value: unknown): DefaultMcpServerAudience | undefined | null {
  if (value === undefined) return undefined;
  if (!isRecord(value)) return null;
  const audience: DefaultMcpServerAudience = {};
  for (const key of ["roles", "oidcRoles"] as const) {
    const list = value[key];
    if (list === undefined) continue;
    if (!Array.isArray(list) || !list.every((role) => typeof role === "string" && role))
      return null;
    audience[key] = list;
  }
  return audience;
}

/** Whether the person has one of the entry's roles. An entry without an audience is for everyone. */
export function inAudience(server: DefaultMcpServer, held: DefaultMcpServerRoles): boolean {
  const audience = server.audience;
  if (!audience?.roles?.length && !audience?.oidcRoles?.length) return true;
  return Boolean(
    audience.roles?.some((role) => held.roles.includes(role)) ||
      audience.oidcRoles?.some((role) => held.oidcRoles.includes(role)),
  );
}

/** The `roles` claim of an ID token, as strings. */
export function oidcRolesClaim(claims: Record<string, unknown> | undefined): string[] {
  const roles = claims?.roles;
  return Array.isArray(roles) ? roles.filter((role) => typeof role === "string") : [];
}

type Deps = { prisma: PrismaClient; secrets: SecretStore };

/**
 * Adds each default server the person is in the audience for and has not been given before.
 * A server they already have under the same slug counts as given. Once given, a server is never
 * added again, so one the person deleted stays deleted. Its `alwaysAllowTools` get an
 * "Always allow" rule when it is added, never later, so a rule the person changed stays changed.
 * Returns the slugs added.
 */
export async function provisionDefaultMcpServers(
  deps: Deps,
  userId: string,
  servers: DefaultMcpServer[],
  held: DefaultMcpServerRoles,
): Promise<string[]> {
  const eligible = servers.filter((server) => inAudience(server, held));
  if (!eligible.length) return [];
  const given = new Set(
    (
      await deps.prisma.defaultMcpServerOffer.findMany({
        where: { userId, slug: { in: eligible.map((server) => server.input.slug) } },
        select: { slug: true },
      })
    ).map((offer) => offer.slug),
  );
  const pending = eligible.filter((server) => !given.has(server.input.slug));
  if (!pending.length) return [];
  const home = await deps.prisma.spaceMember.findFirst({
    where: { userId, role: "owner" },
    orderBy: { createdAt: "asc" },
    select: { spaceId: true },
  });
  if (!home) return [];
  const actor = { spaceId: home.spaceId, userId };
  const added: string[] = [];
  for (const server of pending) {
    const blob = buildMcpCredentialBlob(server.input);
    const stored = blob
      ? await deps.secrets.put(blob, createAdapterContext(actor, "mcp", "mcp.create"))
      : null;
    const created = await runAsTenant(actor, () =>
      deps.prisma.$transaction(async (tx) => {
        // Claimed first, so a concurrent sign-in adds the server once.
        const claimed = await tx.defaultMcpServerOffer.createMany({
          data: [{ userId, slug: server.input.slug }],
          skipDuplicates: true,
        });
        if (claimed.count === 0) return false;
        const existing = await tx.mcpServer.findFirst({
          where: { ...availableInSpace(actor), slug: server.input.slug },
          select: { id: true },
        });
        if (existing) return false;
        const row = await insertMcpServer(tx, actor, server.input, stored, { everywhere: true });
        if (!server.shareWithAllBots)
          await tx.mcpServer.update({ where: { id: row.id }, data: { allBots: false } });
        for (const tool of server.alwaysAllowTools ?? []) {
          await saveApprovalRule(tx, actor, {
            effect: "always_allow",
            matchKind: "tool",
            matchValue: `mcp__${server.input.slug}__${tool}`,
          });
        }
        return true;
      }),
    ).catch((error: unknown) => {
      // One entry the person cannot take (a name used elsewhere) must not stop the rest.
      getLogger().warn("default MCP server not added", {
        "mcp.server.slug": server.input.slug,
        error: error instanceof Error ? error.message : String(error),
      });
      return false;
    });
    if (created) added.push(server.input.slug);
    else if (stored) await deleteSecretBestEffort(deps.secrets, stored.ciphertext, stored.id);
  }
  return added;
}
