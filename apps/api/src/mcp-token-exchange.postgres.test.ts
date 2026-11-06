import { randomUUID } from "node:crypto";
import type { PrismaClient } from "@rakazo/db";
import { asSystem, createDb } from "@rakazo/db";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { recordDirectoryIdentity, serviceRoles } from "./mcp-token-exchange.js";

const dbTests =
  process.env.VERIFY_DATABASE && process.env.DATABASE_URL ? describe.sequential : describe.skip;

dbTests("directory identity and service roles", () => {
  const users: string[] = [];
  let db: ReturnType<typeof createDb>;
  let prisma: PrismaClient;

  beforeAll(() => {
    db = createDb(process.env.DATABASE_URL!);
    prisma = asSystem(db.prisma);
  });

  afterAll(async () => {
    await prisma.user.deleteMany({ where: { id: { in: users } } });
    await db.prisma.$disconnect();
    await db.pool.end();
  });

  async function newUser() {
    const id = randomUUID();
    users.push(id);
    await prisma.user.create({
      data: { id, name: "Directory", email: `${id}@example.test`, emailVerified: true },
    });
    return id;
  }

  it("keeps the directory object of the newest sign-in, for one user only", async () => {
    const [first, second] = [await newUser(), await newUser()];
    const objectId = randomUUID();
    await recordDirectoryIdentity(db.prisma, first, { tid: "tenant", oid: objectId });
    await recordDirectoryIdentity(db.prisma, second, { tid: "tenant", oid: objectId });
    expect(await prisma.directoryIdentity.findMany({ where: { objectId } })).toEqual([
      expect.objectContaining({ userId: second, tenantId: "tenant" }),
    ]);
    // A sign-in whose token names no directory object leaves the record alone.
    await recordDirectoryIdentity(db.prisma, second, { sub: "pairwise" });
    expect(await prisma.directoryIdentity.count({ where: { userId: second } })).toBe(1);
  });

  it("lists the roles services reported", async () => {
    const userId = await newUser();
    await prisma.serviceRole.createMany({
      data: [
        { userId, issuer: "https://a.example.test", role: "staff" },
        { userId, issuer: "https://b.example.test", role: "staff" },
      ],
    });
    expect(await serviceRoles(db.prisma, userId)).toEqual(["staff"]);
  });
});
