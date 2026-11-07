import { afterEach, describe, expect, it, vi } from "vitest";
import { ConnectorRegistry } from "./composio-connector.js";
import { McpConnector } from "./mcp-connector.js";

afterEach(() => vi.unstubAllGlobals());

const NETWORK = {
  fetch: (input: string | URL | Request, init?: RequestInit) => globalThis.fetch(input, init),
  resolveHostname: async () => [{ address: "203.0.113.10", family: 4 }],
};

function server(id: string, name: string, extra: Record<string, unknown> = {}) {
  return {
    id,
    slug: id,
    name,
    transport: "streamable_http",
    endpoint: `https://mcp.example.test/${id}`,
    secretId: null,
    args: [],
    revision: 1,
    ...extra,
  };
}

/** MCP servers by path: each answers `initialize` with its instructions, or fails with 500. */
function mcpServers(servers: Record<string, { instructions?: string; down?: boolean }>) {
  const requests: string[] = [];
  const fetch = vi.fn(async (input: string | URL | Request, init?: RequestInit) => {
    const request = input instanceof Request ? input : new Request(input, init);
    const id = new URL(request.url).pathname.slice(1);
    const behavior = servers[id];
    if (!behavior) throw new Error(`Unexpected request: ${request.url}`);
    if (request.method !== "POST") return new Response(null, { status: 405 });
    if (behavior.down) return new Response("boom", { status: 500 });
    const message = JSON.parse(await request.text()) as { id?: number; method?: string };
    if (message.method) requests.push(`${id} ${message.method}`);
    if (message.method === "initialize") {
      return Response.json({
        jsonrpc: "2.0",
        id: message.id,
        result: {
          protocolVersion: "2025-11-25",
          capabilities: { tools: {} },
          serverInfo: { name: id, version: "1" },
          ...(behavior.instructions ? { instructions: behavior.instructions } : {}),
        },
      });
    }
    if (message.method === "tools/list") {
      return Response.json({
        jsonrpc: "2.0",
        id: message.id,
        result: { tools: [{ name: "echo", inputSchema: { type: "object" } }] },
      });
    }
    return new Response(null, { status: 202 });
  });
  return { fetch, requests };
}

const context = (botId: string) =>
  ({
    operationId: "test",
    traceId: "test",
    spaceId: "w1",
    userId: "u1",
    botId,
    signal: new AbortController().signal,
  }) as const;

describe("MCP server instructions", () => {
  it("come only from servers attached to the bot that are connected, without asking them again", async () => {
    const docs = server("docs", "Docs");
    const down = server("down", "Down");
    const otherBots = server("other", "Other bot's server");
    const shared = server("shared", "Shared", { allBots: true });
    const quiet = server("quiet", "Quiet");
    const assignments = [
      { botId: "bot-1", server: docs },
      { botId: "bot-1", server: down },
      { botId: "bot-1", server: quiet },
      { botId: "bot-2", server: otherBots },
    ].map((row) => ({
      ...row,
      serverId: row.server.id,
      spaceId: "w1",
      userId: "u1",
      allowAllTools: true,
      allowedTools: [],
    }));
    const prisma = {
      botMcpServer: {
        findMany: vi.fn(async ({ where }: { where: { botId: string } }) =>
          assignments.filter((row) => row.botId === where.botId),
        ),
      },
      mcpServer: { findMany: vi.fn(async () => [shared]) },
    };
    const servers = mcpServers({
      docs: { instructions: "Search before you fetch." },
      down: { instructions: "Never seen.", down: true },
      other: { instructions: "Not this bot's." },
      shared: { instructions: "Shared with every bot." },
      quiet: {},
    });
    vi.stubGlobal("fetch", servers.fetch);
    const mcp = new McpConnector(prisma as never, {} as never, { network: NETWORK });
    const registry = new ConnectorRegistry([mcp]);
    try {
      // Before discovery no session is open, so nothing counts as connected.
      expect(await registry.discoverInstructions(context("bot-1"))).toEqual([]);
      await registry.discoverTools(context("bot-1"));
      await registry.discoverTools(context("bot-2"));
      const initialized = servers.requests.filter((line) => line.endsWith(" initialize"));
      expect(initialized.sort()).toEqual([
        "docs initialize",
        "other initialize",
        "quiet initialize",
        "shared initialize",
      ]);
      const before = servers.fetch.mock.calls.length;
      const instructions = await registry.discoverInstructions(context("bot-1"));
      expect(servers.fetch.mock.calls.length).toBe(before);
      expect(instructions).toEqual(
        expect.arrayContaining([
          { source: "Docs", text: "Search before you fetch." },
          { source: "Shared", text: "Shared with every bot." },
        ]),
      );
      expect(instructions).toHaveLength(2);
    } finally {
      await mcp.close();
    }
  });

  it("stop once the server's settings change until discovery reconnects it", async () => {
    const docs = server("docs", "Docs");
    const assignment = {
      botId: "bot-1",
      serverId: "docs",
      spaceId: "w1",
      userId: "u1",
      allowAllTools: true,
      allowedTools: [],
      server: docs,
    };
    const prisma = {
      botMcpServer: { findMany: vi.fn(async () => [assignment]) },
      mcpServer: { findMany: vi.fn(async () => []) },
    };
    vi.stubGlobal("fetch", mcpServers({ docs: { instructions: "Use search." } }).fetch);
    const mcp = new McpConnector(prisma as never, {} as never, { network: NETWORK });
    try {
      await mcp.discoverTools(context("bot-1"));
      expect(await mcp.discoverInstructions(context("bot-1"))).toEqual([
        { source: "Docs", text: "Use search." },
      ]);
      docs.revision += 1;
      expect(await mcp.discoverInstructions(context("bot-1"))).toEqual([]);
      await mcp.discoverTools(context("bot-1"));
      expect(await mcp.discoverInstructions(context("bot-1"))).toEqual([
        { source: "Docs", text: "Use search." },
      ]);
    } finally {
      await mcp.close();
    }
  });
});
