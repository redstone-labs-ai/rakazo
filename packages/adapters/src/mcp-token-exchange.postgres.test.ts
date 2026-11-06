import { randomUUID } from "node:crypto";
import { asSystem, createDb, runAsTenant } from "@rakazo/db";
import { afterAll, describe, expect, it } from "vitest";
import type { OAuthMaterial } from "./mcp-oauth.js";
import { McpOAuthBroker } from "./mcp-oauth.js";
import { McpTokenExchange } from "./mcp-token-exchange.js";
import { EncryptedSecretStore } from "./secrets.js";

const databaseUrl = process.env.DATABASE_URL;
const describePostgres =
  process.env.VERIFY_DATABASE && databaseUrl ? describe.sequential : describe.skip;

const ENDPOINT = "https://mcp.example.test/mcp";
const ISSUER = "https://auth.example.test/api/auth";
const TOKEN_ENDPOINT = `${ISSUER}/oauth2/token`;
const OTHER_ISSUER = "https://vendor.example.test";

describePostgres("MCP grants issued by token exchange (PostgreSQL contract)", () => {
  const db = createDb(databaseUrl!);
  const users: string[] = [];
  const secrets = new EncryptedSecretStore("test-only-encryption-key");

  afterAll(async () => {
    const prisma = asSystem(db.prisma);
    await prisma.organization.deleteMany({ where: { id: { in: users } } });
    await prisma.user.deleteMany({ where: { id: { in: users } } });
    await db.prisma.$disconnect();
    await db.pool.end();
  });

  /** A deployment holding a client at ISSUER, with an authorization server and MCP server behind it. */
  function harness(options: { refresh?: "ok" | "session-ended"; issuer?: string } = {}) {
    const exchanges: URLSearchParams[] = [];
    const refreshes: Array<{ url: string; body: URLSearchParams }> = [];
    let issued = 0;
    const networkFetch = async (input: string | URL | Request, init?: RequestInit) => {
      const url = new URL(input instanceof Request ? input.url : String(input));
      if (url.pathname.startsWith("/.well-known/oauth-protected-resource"))
        return Response.json({
          resource: ENDPOINT,
          authorization_servers: [options.issuer ?? ISSUER],
        });
      if (url.pathname.startsWith("/.well-known/oauth-authorization-server"))
        return Response.json({
          issuer: ISSUER,
          authorization_endpoint: `${ISSUER}/oauth2/authorize`,
          token_endpoint: TOKEN_ENDPOINT,
          response_types_supported: ["code"],
          token_endpoint_auth_methods_supported: ["client_secret_basic", "client_secret_post"],
        });
      if (url.href === TOKEN_ENDPOINT || url.pathname.endsWith("/token")) {
        const raw = init?.body;
        const body = new URLSearchParams(
          raw instanceof ArrayBuffer || ArrayBuffer.isView(raw)
            ? new TextDecoder().decode(raw)
            : String(raw),
        );
        refreshes.push({ url: url.href, body });
        if (options.refresh === "session-ended")
          return Response.json(
            { error: "invalid_grant", error_description: "session ended; exchange again" },
            { status: 400 },
          );
        return Response.json({
          access_token: "refreshed",
          refresh_token: "refresh-next",
          token_type: "Bearer",
          expires_in: 300,
        });
      }
      return new Response("not found", { status: 404 });
    };
    const exchange = new McpTokenExchange(
      [
        {
          issuer: ISSUER,
          tokenEndpoint: TOKEN_ENDPOINT,
          clientId: "deployment-client",
          clientSecret: "deployment-secret",
          subjectScope: "api://service/MCP.Access",
        },
      ],
      {
        subjectToken: async (_userId, scope) => `sso-token-for:${scope}`,
        fetch: async (_input, init) => {
          exchanges.push(new URLSearchParams(String(init?.body)));
          issued += 1;
          return Response.json({
            access_token: `issued-${issued}`,
            refresh_token: `issued-refresh-${issued}`,
            token_type: "Bearer",
            expires_in: 300,
            role: "staff",
          });
        },
      },
    );
    const broker = new McpOAuthBroker(
      db.prisma,
      secrets,
      {
        resolveHostname: async () => [{ address: "203.0.113.10", family: 4 as const }],
        fetch: networkFetch,
      },
      false,
      exchange,
    );
    return { broker, exchanges, refreshes };
  }

  async function seed(material?: OAuthMaterial) {
    const prisma = asSystem(db.prisma);
    const id = randomUUID();
    users.push(id);
    await prisma.user.create({
      data: { id, name: "Exchange", email: `${id}@example.test`, emailVerified: true },
    });
    await prisma.organization.create({
      data: { id, name: "Exchange", slug: id, createdAt: new Date() },
    });
    await prisma.space.create({
      data: { id, organizationId: id, name: "Exchange", isDefault: true },
    });
    let secretId: string | undefined;
    if (material) {
      const stored = await secrets.put(JSON.stringify(material), {
        operationId: "test",
        traceId: "test",
        spaceId: id,
        userId: id,
        signal: new AbortController().signal,
      });
      await prisma.secret.create({
        data: {
          id: stored.id,
          userId: id,
          spaceId: id,
          kind: "mcp",
          ciphertext: stored.ciphertext,
        },
      });
      secretId = stored.id;
    }
    const server = await prisma.mcpServer.create({
      data: {
        spaceId: id,
        userId: id,
        slug: "service",
        name: "Service",
        transport: "streamable_http",
        endpoint: ENDPOINT,
        secretId,
      },
    });
    return { server, context: { spaceId: id, userId: id } };
  }

  async function stored(serverId: string): Promise<OAuthMaterial> {
    const prisma = asSystem(db.prisma);
    const row = await prisma.mcpServer.findUniqueOrThrow({ where: { id: serverId } });
    const secret = await prisma.secret.findUniqueOrThrow({ where: { id: row.secretId! } });
    return JSON.parse(await secrets.load(secret.ciphertext, secret.id));
  }

  /** Makes the stored grant's access token lapse, as it does five minutes after issue. */
  async function expire(serverId: string, context: { spaceId: string; userId: string }) {
    const material = await stored(serverId);
    material.oauth!.obtainedAt = Date.now() - 10 * 60_000;
    const prisma = asSystem(db.prisma);
    const row = await prisma.mcpServer.findUniqueOrThrow({ where: { id: serverId } });
    const next = await secrets.put(JSON.stringify(material), {
      operationId: "test",
      traceId: "test",
      ...context,
      signal: new AbortController().signal,
    });
    await prisma.secret.create({
      data: { id: next.id, ...context, kind: "mcp", ciphertext: next.ciphertext },
    });
    await prisma.mcpServer.update({ where: { id: serverId }, data: { secretId: next.id } });
    await prisma.secret.delete({ where: { id: row.secretId! } });
  }

  it("connects a server with an exchanged grant, keeping the client secret out of it", async () => {
    const { broker, exchanges } = harness();
    const { server, context } = await seed();

    const connected = await runAsTenant(context, () => broker.connectIssued(server, context));

    expect(connected).toBe(true);
    expect(exchanges).toHaveLength(1);
    expect(Object.fromEntries(exchanges[0]!)).toEqual({
      grant_type: "urn:ietf:params:oauth:grant-type:token-exchange",
      subject_token: "sso-token-for:api://service/MCP.Access",
      subject_token_type: "urn:ietf:params:oauth:token-type:access_token",
      client_id: "deployment-client",
      client_secret: "deployment-secret",
      resource: ENDPOINT,
    });
    const current = await asSystem(db.prisma).mcpServer.findUniqueOrThrow({
      where: { id: server.id },
    });
    expect(await runAsTenant(context, () => broker.statusFor(current, context))).toBe("connected");
    const material = await stored(server.id);
    expect(material.oauth?.tokens?.access_token).toBe("issued-1");
    expect(material.oauth?.discoveryState?.authorizationServerMetadata?.token_endpoint).toBe(
      TOKEN_ENDPOINT,
    );
    expect(JSON.stringify(material)).not.toContain("deployment-secret");
    // Connected servers are left alone on the next sign-in.
    expect(await runAsTenant(context, () => broker.connectIssued(current, context))).toBe(false);
  });

  it("refreshes an exchanged grant at the discovered endpoint with the deployment's secret", async () => {
    const { broker, refreshes } = harness();
    const { server, context } = await seed();
    await runAsTenant(context, () => broker.connectIssued(server, context));
    await expire(server.id, context);

    const current = await asSystem(db.prisma).mcpServer.findUniqueOrThrow({
      where: { id: server.id },
    });
    const provider = await runAsTenant(context, () => broker.providerFor(current, context));
    const tokens = await runAsTenant(context, () => provider?.tokens());

    expect(tokens?.access_token).toBe("refreshed");
    expect(refreshes).toHaveLength(1);
    expect(refreshes[0]?.url).toBe(TOKEN_ENDPOINT);
    expect(Object.fromEntries(refreshes[0]!.body)).toMatchObject({
      grant_type: "refresh_token",
      refresh_token: "issued-refresh-1",
      client_id: "deployment-client",
      client_secret: "deployment-secret",
      resource: ENDPOINT,
    });
  });

  it("exchanges again when the refresh says the service session ended", async () => {
    const { broker, exchanges } = harness({ refresh: "session-ended" });
    const { server, context } = await seed();
    await runAsTenant(context, () => broker.connectIssued(server, context));
    await expire(server.id, context);

    const current = await asSystem(db.prisma).mcpServer.findUniqueOrThrow({
      where: { id: server.id },
    });
    const provider = await runAsTenant(context, () => broker.providerFor(current, context));
    const tokens = await runAsTenant(context, () => provider?.tokens());

    expect(tokens?.access_token).toBe("issued-2");
    expect(exchanges).toHaveLength(2);
    expect((await stored(server.id)).oauth?.tokens?.refresh_token).toBe("issued-refresh-2");
  });

  it("exchanges again, instead of refreshing, when the server refuses an unexpired token", async () => {
    const { broker, exchanges, refreshes } = harness();
    const { server, context } = await seed();
    await runAsTenant(context, () => broker.connectIssued(server, context));

    const current = await asSystem(db.prisma).mcpServer.findUniqueOrThrow({
      where: { id: server.id },
    });
    const provider = await runAsTenant(context, () => broker.providerFor(current, context));
    await runAsTenant(context, () => provider!.replaceRejectedToken());

    expect(exchanges).toHaveLength(2);
    expect(refreshes).toHaveLength(0);
    expect(await runAsTenant(context, () => provider?.tokens())).toMatchObject({
      access_token: "issued-2",
    });
    expect((await stored(server.id)).oauth?.tokens?.access_token).toBe("issued-2");
  });

  it("leaves servers of other authorization servers, or with their own client, alone", async () => {
    const other = harness({ issuer: OTHER_ISSUER });
    const first = await seed();
    expect(
      await runAsTenant(first.context, () =>
        other.broker.connectIssued(first.server, first.context),
      ),
    ).toBe(false);
    expect(other.exchanges).toHaveLength(0);

    const { broker, exchanges } = harness();
    const own = await seed({ oauthClient: { client_id: "own-client" } });
    expect(
      await runAsTenant(own.context, () => broker.connectIssued(own.server, own.context)),
    ).toBe(false);
    expect(exchanges).toHaveLength(0);
  });
});
