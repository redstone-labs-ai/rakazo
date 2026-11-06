import { describe, expect, it, vi } from "vitest";
import type { DefaultMcpServer } from "./default-mcp-servers.js";
import { defaultMcpServersFromEnv, inAudience, oidcRolesClaim } from "./default-mcp-servers.js";

const docs = {
  slug: "docs",
  name: "Docs",
  description: "Team documents",
  transport: "streamable_http",
  endpoint: "https://mcp.example.test/mcp",
};

describe("defaultMcpServersFromEnv", () => {
  it("reads an array or a catalog object, defaulting to every bot", () => {
    const fromArray = defaultMcpServersFromEnv(JSON.stringify([docs]));
    expect(fromArray).toEqual([
      {
        input: expect.objectContaining({ ...docs, enabled: true }),
        shareWithAllBots: true,
      },
    ]);
    expect(defaultMcpServersFromEnv(JSON.stringify({ version: 1, servers: [docs] }))).toEqual(
      fromArray,
    );
  });

  it("keeps a client id, scope and audience but never a client secret", () => {
    const [server] = defaultMcpServersFromEnv(
      JSON.stringify([
        {
          ...docs,
          shareWithAllBots: false,
          oauthClient: { clientId: "public-client", clientSecret: "do-not-copy", scope: "read" },
          audience: { roles: ["admin"], oidcRoles: ["Investments"] },
        },
      ]),
    );
    expect(server?.shareWithAllBots).toBe(false);
    expect(server?.input).toMatchObject({
      oauthClient: { clientId: "public-client", scope: "read" },
    });
    expect(JSON.stringify(server)).not.toContain("do-not-copy");
    expect(server?.audience).toEqual({ roles: ["admin"], oidcRoles: ["Investments"] });
  });

  it("logs and leaves out what does not parse", () => {
    const warn = vi.fn();
    const servers = defaultMcpServersFromEnv(
      JSON.stringify([
        docs,
        { ...docs, name: "Again" },
        { ...docs, slug: "Bad Slug" },
        { ...docs, slug: "local", transport: "stdio", command: "run" },
        { ...docs, slug: "aud", audience: { roles: "admin" } },
        { ...docs, slug: "off", enabled: false },
        "not an object",
      ]),
      warn,
    );
    expect(servers.map((server) => server.input.slug)).toEqual(["docs"]);
    expect(warn).toHaveBeenCalledTimes(5);
    expect(warn).toHaveBeenCalledWith(expect.stringContaining("duplicate slug docs"));
    expect(defaultMcpServersFromEnv("{", warn)).toEqual([]);
    expect(defaultMcpServersFromEnv('"text"', warn)).toEqual([]);
    expect(defaultMcpServersFromEnv(undefined, warn)).toEqual([]);
  });
});

describe("inAudience", () => {
  const server = (audience?: DefaultMcpServer["audience"]): DefaultMcpServer => ({
    input: { ...docs, transport: "streamable_http", enabled: true, headers: {} },
    shareWithAllBots: true,
    ...(audience ? { audience } : {}),
  });
  const none = { roles: [], oidcRoles: [] };

  it("offers a server without an audience to everyone", () => {
    expect(inAudience(server(), none)).toBe(true);
    expect(inAudience(server({ roles: [] }), none)).toBe(true);
  });

  it("offers a server to anyone holding one of its roles from either source", () => {
    const audience = { roles: ["admin"], oidcRoles: ["Investments"] };
    expect(inAudience(server(audience), none)).toBe(false);
    expect(inAudience(server(audience), { roles: ["admin"], oidcRoles: [] })).toBe(true);
    expect(inAudience(server(audience), { roles: [], oidcRoles: ["Investments"] })).toBe(true);
    // A role from one source never stands in for the other.
    expect(inAudience(server(audience), { roles: ["Investments"], oidcRoles: ["admin"] })).toBe(
      false,
    );
  });
});

describe("oidcRolesClaim", () => {
  it("reads the roles claim as strings", () => {
    expect(oidcRolesClaim({ roles: ["A", 2, "B"] })).toEqual(["A", "B"]);
    expect(oidcRolesClaim({ roles: "A" })).toEqual([]);
    expect(oidcRolesClaim(undefined)).toEqual([]);
  });
});
