import { randomUUID } from "node:crypto";
import { afterAll, describe, expect, it } from "vitest";
import { createDb } from "./client.js";
import { linkMessagingIdentityByDirectory } from "./messaging.js";
import { asSystem, runAsSystem } from "./row-security.js";

const databaseUrl = process.env.DATABASE_URL;
const describePostgres =
  process.env.VERIFY_DATABASE && databaseUrl ? describe.sequential : describe.skip;

describePostgres("linkMessagingIdentityByDirectory (PostgreSQL)", () => {
  const db = createDb(databaseUrl!);
  const prisma = asSystem(db.prisma);
  const users: string[] = [];

  afterAll(async () => {
    await prisma.organization.deleteMany({ where: { id: { in: users } } });
    await prisma.user.deleteMany({ where: { id: { in: users } } });
    await db.prisma.$disconnect();
    await db.pool.end();
  });

  async function person() {
    const id = randomUUID();
    users.push(id);
    await prisma.user.create({
      data: { id, name: "Ada", email: `${id}@example.test`, emailVerified: true },
    });
    await prisma.organization.create({
      data: { id, name: "Ada", slug: id, createdAt: new Date() },
    });
    await prisma.space.create({ data: { id, organizationId: id, name: "Ada", isDefault: true } });
    const bot = (name: string) =>
      prisma.bot.create({ data: { spaceId: id, userId: id, name, color: "#3EC5A8" } });
    const first = await bot("First");
    const second = await bot("Second");
    const objectId = randomUUID();
    await prisma.directoryIdentity.create({ data: { userId: id, tenantId: "tenant", objectId } });
    return { userId: id, first, second, objectId };
  }

  // Production links from the inbound messaging webhook.
  const link = (address: string, objectId: string) =>
    runAsSystem("messaging inbound", () =>
      linkMessagingIdentityByDirectory(db.prisma, {
        provider: "teams",
        address,
        dmThreadId: `teams:${address}`,
        directoryUser: { tenantId: "TENANT", objectId: objectId.toUpperCase() },
      }),
    );

  it("links to the earliest bot without a chat app, in any letter case", async () => {
    const { userId, first, second, objectId } = await person();
    await prisma.messagingIdentity.create({
      data: {
        provider: "sms",
        address: `+1${Date.now()}`,
        userId,
        spaceId: userId,
        botId: first.id,
      },
    });

    const linked = await link(`29:${randomUUID()}`, objectId);

    expect(linked).toMatchObject({ userId, spaceId: userId, botId: second.id });
    expect(await link(`29:${randomUUID()}`, objectId)).toBeNull();
  });

  it("links one address once when two messages race", async () => {
    const { objectId } = await person();
    const address = `29:${randomUUID()}`;
    const results = await Promise.all([link(address, objectId), link(address, objectId)]);
    expect(results.filter(Boolean)).toHaveLength(1);
  });

  it("links nobody for an unknown directory account", async () => {
    expect(await link(`29:${randomUUID()}`, randomUUID())).toBeNull();
  });
});
