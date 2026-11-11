import { describe, expect, it, vi } from "vitest";
import type { TeamsListening, TeamsListeningLookup } from "./teams-listening.js";
import { groupMessageRunsAgent, managebotTeamsListening } from "./teams-listening.js";

const CHAT = "19:group-chat@thread.v2";
const encoded = Buffer.from(CHAT).toString("base64url");
const service = Buffer.from("https://smba.trafficmanager.net/amer/").toString("base64url");

function groupEvent(overrides: { provider?: string; mentionsBot?: boolean; agentId?: string }) {
  const provider = overrides.provider ?? "teams";
  return {
    provider,
    threadId: `${provider}:${encoded}:${service}`,
    groupMentionRequired: true,
    ...(overrides.mentionsBot ? { mentionsBot: true } : {}),
    ...(overrides.agentId ? { agentId: overrides.agentId } : {}),
  };
}

/** Listening per (chat, agent), as Managebot stores it. */
function lookupFrom(settings: Record<string, TeamsListening>) {
  return vi.fn<TeamsListeningLookup>(
    async ({ conversationId, agentId }) => settings[`${conversationId}|${agentId}`] ?? "mentions",
  );
}

describe("group message listening", () => {
  it("drops a message without a mention when listening is mentions", async () => {
    const lookup = lookupFrom({ [`${CHAT}|aide`]: "mentions" });
    expect(await groupMessageRunsAgent(groupEvent({}), "aide", lookup)).toBe(false);
    expect(lookup).toHaveBeenCalledWith({ conversationId: CHAT, agentId: "aide" });
  });

  it("accepts a message without a mention when listening is all", async () => {
    const lookup = lookupFrom({ [`${CHAT}|aide`]: "all" });
    expect(await groupMessageRunsAgent(groupEvent({}), "aide", lookup)).toBe(true);
  });

  it("accepts a mention in both modes without a lookup", async () => {
    for (const mode of ["mentions", "all"] as const) {
      const lookup = lookupFrom({ [`${CHAT}|aide`]: mode });
      expect(await groupMessageRunsAgent(groupEvent({ mentionsBot: true }), "aide", lookup)).toBe(
        true,
      );
      expect(lookup).not.toHaveBeenCalled();
    }
  });

  it("drops a message without a mention when no lookup is configured", async () => {
    expect(await groupMessageRunsAgent(groupEvent({}), "aide", undefined)).toBe(false);
  });

  it("checks listening per agent in the same chat", async () => {
    const lookup = lookupFrom({ [`${CHAT}|aide`]: "all", [`${CHAT}|ava`]: "mentions" });
    expect(await groupMessageRunsAgent(groupEvent({}), "aide", lookup)).toBe(true);
    expect(await groupMessageRunsAgent(groupEvent({}), "ava", lookup)).toBe(false);
  });

  it("runs only the agent that owns the receiving bot", async () => {
    const lookup = lookupFrom({ [`${CHAT}|ava`]: "all", [`${CHAT}|aide`]: "all" });
    const toAva = groupEvent({ provider: "teams-ava", agentId: "ava", mentionsBot: true });
    expect(await groupMessageRunsAgent(toAva, "ava", lookup)).toBe(true);
    expect(await groupMessageRunsAgent(toAva, "aide", lookup)).toBe(false);
    const quietToAva = groupEvent({ provider: "teams-ava", agentId: "ava" });
    expect(await groupMessageRunsAgent(quietToAva, "ava", lookup)).toBe(true);
    expect(await groupMessageRunsAgent(quietToAva, "aide", lookup)).toBe(false);
  });
});

describe("Managebot listening lookup", () => {
  it("asks for the chat and agent, and reads only all as all", async () => {
    const fetchMock = vi.fn<typeof fetch>(async () => Response.json({ listening: "all" }));
    const lookup = managebotTeamsListening({
      url: "https://api.example.test/internal/teams/listening",
      token: "internal-token",
      fetch: fetchMock,
    });
    expect(await lookup({ conversationId: CHAT, agentId: "aide" })).toBe("all");
    const [url, init] = fetchMock.mock.calls[0]!;
    expect(String(url)).toContain(`conversationId=${encodeURIComponent(CHAT)}`);
    expect(String(url)).toContain("agentId=aide");
    expect(init?.headers).toEqual({ authorization: "Bearer internal-token" });
  });

  it("reads a failure or an unknown answer as mentions", async () => {
    const failing = managebotTeamsListening({
      url: "https://api.example.test/l",
      token: undefined,
      fetch: vi.fn<typeof fetch>(async () => new Response("no", { status: 500 })),
    });
    expect(await failing({ conversationId: CHAT, agentId: "aide" })).toBe("mentions");
    const throwing = managebotTeamsListening({
      url: "https://api.example.test/l",
      token: undefined,
      fetch: vi.fn<typeof fetch>(async () => {
        throw new Error("down");
      }),
    });
    expect(await throwing({ conversationId: CHAT, agentId: "aide" })).toBe("mentions");
    const odd = managebotTeamsListening({
      url: "https://api.example.test/l",
      token: undefined,
      fetch: vi.fn<typeof fetch>(async () => Response.json({ listening: "everything" })),
    });
    expect(await odd({ conversationId: CHAT, agentId: "aide" })).toBe("mentions");
  });
});
