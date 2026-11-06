import { randomUUID } from "node:crypto";
import { EncryptedSecretStore } from "@rakazo/adapters";
import type { PrismaClient } from "@rakazo/db";
import { asSystem, bootstrapUserSpace, createDb } from "@rakazo/db";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import type { DefaultMcpServerRoles } from "./default-mcp-servers.js";
import { defaultMcpServersFromEnv, provisionDefaultMcpServers } from "./default-mcp-servers.js";

const dbTests =
  process.env.VERIFY_DATABASE && process.env.DATABASE_URL ? describe.sequential : describe.skip;

dbTests("default MCP servers", () => {
  const users: string[] = [];
  const secrets = new EncryptedSecretStore("test-only-encryption-key");
  let db: ReturnType<typeof createDb>;
  let prisma: PrismaClient;
  const servers = defaultMcpServersFromEnv(
    JSON.stringify([
      {
        slug: "docs",
        name: "Docs",
        transport: "streamable_http",
        endpoint: "https://mcp.example.test/mcp",
      },
      {
        slug: "deals",
        name: "Deals",
        transport: "streamable_http",
        endpoint: "https://deals.example.test/mcp",
        oauthClient: { clientId: "deals-client" },
        shareWithAllBots: false,
        audience: { oidcRoles: ["Investments"] },
      },
    ]),
  );
  const nobody: DefaultMcpServerRoles = { roles: [], oidcRoles: [] };

  beforeAll(() => {
    db = createDb(process.env.DATABASE_URL!);
    prisma = asSystem(db.prisma);
  });

  afterAll(async () => {
    await prisma.organization.deleteMany({
      where: { members: { some: { userId: { in: users } } } },
    });
    await prisma.user.deleteMany({ where: { id: { in: users } } });
    await db.prisma.$disconnect();
    await db.pool.end();
  });

  async function newUser() {
    const id = randomUUID();
    users.push(id);
    await prisma.user.create({
      data: { id, name: "Default", email: `${id}@example.test`, emailVerified: true },
    });
    const { spaceId } = await bootstrapUserSpace(
      prisma,
      { id },
      {
        signupsEnabled: "true",
        signupAllowlist: undefined,
      },
    );
    return { userId: id, spaceId };
  }

  // The provisioning code sets its own tenant context, as it does after a real sign-in.
  const provision = (userId: string, held = nobody) =>
    provisionDefaultMcpServers({ prisma: db.prisma, secrets }, userId, servers, held);

  it("adds each server once, shared with every space and bot", async () => {
    const { userId } = await newUser();
    expect(await provision(userId)).toEqual(["docs"]);
    expect(await provision(userId)).toEqual([]);
    const rows = await prisma.mcpServer.findMany({ where: { userId } });
    expect(rows).toEqual([
      expect.objectContaining({ slug: "docs", allSpaces: true, allBots: true, secretId: null }),
    ]);
  });

  it("adds an audience server once the person holds its role, with its client", async () => {
    const { userId } = await newUser();
    await provision(userId);
    expect(await provision(userId, { roles: [], oidcRoles: ["Investments"] })).toEqual(["deals"]);
    const deals = await prisma.mcpServer.findFirstOrThrow({ where: { userId, slug: "deals" } });
    expect(deals).toMatchObject({ allBots: false, allSpaces: true });
    const secret = await prisma.secret.findFirstOrThrow({ where: { id: deals.secretId! } });
    expect(JSON.parse(await secrets.load(secret.ciphertext, secret.id))).toMatchObject({
      oauthClient: { client_id: "deals-client" },
    });
  });

  it("does not bring back a server the person deleted", async () => {
    const { userId } = await newUser();
    await provision(userId);
    await prisma.mcpServer.deleteMany({ where: { userId, slug: "docs" } });
    expect(await provision(userId)).toEqual([]);
    expect(await prisma.mcpServer.count({ where: { userId } })).toBe(0);
  });

  it("counts a server the person already added under the same slug", async () => {
    const { userId, spaceId } = await newUser();
    await prisma.mcpServer.create({
      data: {
        userId,
        spaceId,
        allSpaces: true,
        slug: "docs",
        name: "My docs",
        transport: "streamable_http",
        endpoint: "https://mcp.example.test/mcp",
      },
    });
    expect(await provision(userId)).toEqual([]);
    const rows = await prisma.mcpServer.findMany({ where: { userId } });
    expect(rows.map((row) => row.name)).toEqual(["My docs"]);
  });

  it("adds a server once when two sign-ins race", async () => {
    const { userId } = await newUser();
    const results = await Promise.all([provision(userId), provision(userId), provision(userId)]);
    expect(results.flat()).toEqual(["docs"]);
    expect(await prisma.mcpServer.count({ where: { userId } })).toBe(1);
  });
});
