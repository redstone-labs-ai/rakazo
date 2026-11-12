import { randomUUID } from "node:crypto";
import type { PrismaClient } from "@rakazo/db";
import { createDb } from "@rakazo/db";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { prismaAgentTeamsApps, prismaAgentTeamsChatInstalls } from "./teams-agent-app-platforms.js";
import { prismaAgentTeamsAppStore } from "./teams-agent-apps/store.js";

const databaseUrl = process.env.DATABASE_URL;
const enabled = process.env.VERIFY_DATABASE === "1" && Boolean(databaseUrl);

// The router's install tracking runs on the provisioner's store against the real tables.
describe.skipIf(!enabled).sequential("agent Teams app installs (PostgreSQL)", () => {
  let prisma: PrismaClient;
  let close: () => Promise<void>;

  beforeAll(() => {
    const db = createDb(databaseUrl ?? "");
    prisma = db.prisma;
    close = async () => {
      await prisma.$disconnect();
      await db.pool.end();
    };
  });
  afterAll(async () => close());

  it("records the first install as visible, ends a removed chat, and lists mountable rows", async () => {
    const store = prismaAgentTeamsAppStore(prisma);
    const botId = `bot-${randomUUID()}`;
    const row = await store.ensure(botId);
    await store.update(row.id, {
      status: "published",
      entraAppId: randomUUID(),
      secretRef: `teams-agent-${botId}`,
      publishedAt: new Date(),
    });
    const installs = prismaAgentTeamsChatInstalls(prisma);

    await installs.recordChatInstall(row.id, "19:chat-a@thread.v2", null);
    await installs.recordChatInstall(row.id, "19:chat-a@thread.v2", null);
    await installs.recordChatInstall(row.id, "19:chat-b@thread.v2", null);
    await installs.markChatRemoved(row.id, "19:chat-b@thread.v2");

    const app = await store.findById(row.id);
    expect(app?.status).toBe("visible");
    expect(app?.firstInstalledAt).toBeInstanceOf(Date);
    expect((await store.activeChats(row.id)).map((chat) => chat.chatId)).toEqual([
      "19:chat-a@thread.v2",
    ]);
    const mountable = await prismaAgentTeamsApps(prisma)();
    expect(mountable.find((found) => found.id === row.id)?.status).toBe("visible");
  });
});
