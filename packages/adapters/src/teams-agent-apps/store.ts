import type { PrismaClient } from "@rakazo/db";
import { runAsSystem } from "@rakazo/db";

export const AGENT_TEAMS_APP_STATUSES = [
  "creating",
  "registered",
  "published",
  "visible",
  "removing",
  "removed",
  "failed",
] as const;
export type AgentTeamsAppStatus = (typeof AGENT_TEAMS_APP_STATUSES)[number];

/** One `agent_teams_apps` row. */
export interface AgentTeamsAppRecord {
  id: string;
  botId: string;
  status: string;
  entraAppId: string | null;
  entraObjectId: string | null;
  secretRef: string | null;
  azureBotName: string | null;
  catalogAppId: string | null;
  manifestVersion: string | null;
  publishedAt: Date | null;
  firstInstalledAt: Date | null;
  lastError: string | null;
  createdAt: Date;
  updatedAt: Date;
}

/** One `agent_teams_chats` row. */
export interface AgentTeamsChatRecord {
  id: string;
  agentAppId: string;
  chatId: string;
  installId: string | null;
  installedAt: Date;
  removedAt: Date | null;
}

export type AgentTeamsAppPatch = Partial<
  Omit<AgentTeamsAppRecord, "id" | "botId" | "createdAt" | "updatedAt">
>;

/**
 * Persistence for provisioning. There is deliberately no delete: a removed agent keeps its row
 * with status `removed`, and nothing (bot or group chat removal) cascades into it.
 */
export interface AgentTeamsAppStore {
  findByBotId(botId: string): Promise<AgentTeamsAppRecord | null>;
  findById(id: string): Promise<AgentTeamsAppRecord | null>;
  /** Creates the row in `creating`, or returns the existing one for this bot. */
  ensure(botId: string): Promise<AgentTeamsAppRecord>;
  update(id: string, patch: AgentTeamsAppPatch): Promise<AgentTeamsAppRecord>;
  /** Chats the agent app is still installed in. */
  activeChats(agentAppId: string): Promise<AgentTeamsChatRecord[]>;
  chatCount(agentAppId: string): Promise<number>;
  /** Inserts the chat, or reactivates it with the new install id. */
  upsertChat(input: {
    agentAppId: string;
    chatId: string;
    installId: string | null;
    installedAt: Date;
  }): Promise<AgentTeamsChatRecord>;
  markChatRemoved(id: string, removedAt: Date): Promise<void>;
}

/** The Prisma store. The tables hold no tenant data, so every call runs as system. */
export function prismaAgentTeamsAppStore(prisma: PrismaClient): AgentTeamsAppStore {
  const system = <T>(reason: string, fn: () => Promise<T>) => runAsSystem(reason, fn);
  return {
    findByBotId: (botId) =>
      system("read an agent Teams app", () =>
        prisma.agentTeamsApp.findUnique({ where: { botId } }),
      ),
    findById: (id) =>
      system("read an agent Teams app", () => prisma.agentTeamsApp.findUnique({ where: { id } })),
    ensure: (botId) =>
      system("create an agent Teams app", () =>
        prisma.agentTeamsApp.upsert({
          where: { botId },
          create: { botId, status: "creating" },
          update: {},
        }),
      ),
    update: (id, patch) =>
      system("update an agent Teams app", () =>
        prisma.agentTeamsApp.update({ where: { id }, data: patch }),
      ),
    activeChats: (agentAppId) =>
      system("list an agent Teams app's chats", () =>
        prisma.agentTeamsChat.findMany({
          where: { agentAppId, removedAt: null },
          orderBy: { installedAt: "asc" },
        }),
      ),
    chatCount: (agentAppId) =>
      system("count an agent Teams app's chats", () =>
        prisma.agentTeamsChat.count({ where: { agentAppId, removedAt: null } }),
      ),
    upsertChat: ({ agentAppId, chatId, installId, installedAt }) =>
      system("record an agent Teams app chat install", () =>
        prisma.agentTeamsChat.upsert({
          where: { agentAppId_chatId: { agentAppId, chatId } },
          create: { agentAppId, chatId, installId, installedAt },
          update: { installId, installedAt, removedAt: null },
        }),
      ),
    markChatRemoved: async (id, removedAt) => {
      await system("mark an agent Teams app chat removed", () =>
        prisma.agentTeamsChat.update({ where: { id }, data: { removedAt } }),
      );
    },
  };
}

/** An in-memory store for tests and local runs. */
export function memoryAgentTeamsAppStore(now: () => Date = () => new Date()): AgentTeamsAppStore & {
  apps: Map<string, AgentTeamsAppRecord>;
  chats: Map<string, AgentTeamsChatRecord>;
} {
  const apps = new Map<string, AgentTeamsAppRecord>();
  const chats = new Map<string, AgentTeamsChatRecord>();
  let sequence = 0;
  const nextId = (prefix: string) => `${prefix}_${++sequence}`;
  const byBot = (botId: string) => [...apps.values()].find((app) => app.botId === botId) ?? null;
  return {
    apps,
    chats,
    findByBotId: async (botId) => {
      const app = byBot(botId);
      return app ? { ...app } : null;
    },
    findById: async (id) => {
      const app = apps.get(id);
      return app ? { ...app } : null;
    },
    ensure: async (botId) => {
      const existing = byBot(botId);
      if (existing) return { ...existing };
      const at = now();
      const app: AgentTeamsAppRecord = {
        id: nextId("app"),
        botId,
        status: "creating",
        entraAppId: null,
        entraObjectId: null,
        secretRef: null,
        azureBotName: null,
        catalogAppId: null,
        manifestVersion: null,
        publishedAt: null,
        firstInstalledAt: null,
        lastError: null,
        createdAt: at,
        updatedAt: at,
      };
      apps.set(app.id, app);
      return { ...app };
    },
    update: async (id, patch) => {
      const app = apps.get(id);
      if (!app) throw new Error(`No agent Teams app ${id}`);
      const updated = { ...app, ...patch, updatedAt: now() };
      apps.set(id, updated);
      return { ...updated };
    },
    activeChats: async (agentAppId) =>
      [...chats.values()]
        .filter((chat) => chat.agentAppId === agentAppId && !chat.removedAt)
        .map((chat) => ({ ...chat })),
    chatCount: async (agentAppId) =>
      [...chats.values()].filter((chat) => chat.agentAppId === agentAppId && !chat.removedAt)
        .length,
    upsertChat: async ({ agentAppId, chatId, installId, installedAt }) => {
      const existing = [...chats.values()].find(
        (chat) => chat.agentAppId === agentAppId && chat.chatId === chatId,
      );
      const chat: AgentTeamsChatRecord = existing
        ? { ...existing, installId, installedAt, removedAt: null }
        : { id: nextId("chat"), agentAppId, chatId, installId, installedAt, removedAt: null };
      chats.set(chat.id, chat);
      return { ...chat };
    },
    markChatRemoved: async (id, removedAt) => {
      const chat = chats.get(id);
      if (chat) chats.set(id, { ...chat, removedAt });
    },
  };
}
