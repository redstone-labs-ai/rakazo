import type { MessagingInboundMessage } from "@rakazo/adapter-kit";
import { teamsConversationIdOfThread } from "@rakazo/adapters";
import { getLogger } from "@rakazo/logging";

/** Managebot's listening setting for one Teams chat and one agent. */
export type TeamsListening = "mentions" | "all";

/**
 * Reads the listening setting for a Teams chat and an agent. Managebot owns the setting.
 * Rakazo never reads Managebot's database.
 */
export type TeamsListeningLookup = (input: {
  conversationId: string;
  agentId: string;
}) => Promise<TeamsListening>;

/** Bounds one lookup so a slow Managebot cannot stall the inbound sink. */
const LOOKUP_TIMEOUT_MS = 3_000;

/**
 * Managebot's internal listening endpoint:
 * `GET <url>?conversationId=<teams id>&agentId=<agent id>` answers `{ "listening": "all" }`.
 * Any other answer, or a failure, reads as `mentions`.
 */
export function managebotTeamsListening(input: {
  url: string;
  token: string | undefined;
  fetch?: typeof fetch;
}): TeamsListeningLookup {
  const send = input.fetch ?? fetch;
  return async ({ conversationId, agentId }) => {
    const url = new URL(input.url);
    url.searchParams.set("conversationId", conversationId);
    url.searchParams.set("agentId", agentId);
    try {
      const response = await send(url, {
        headers: input.token ? { authorization: `Bearer ${input.token}` } : {},
        signal: AbortSignal.timeout(LOOKUP_TIMEOUT_MS),
      });
      if (!response.ok) {
        getLogger().warn("Teams listening lookup failed", { status: response.status });
        return "mentions";
      }
      const body: unknown = await response.json();
      const listening =
        typeof body === "object" && body !== null
          ? (body as Record<string, unknown>).listening
          : undefined;
      return listening === "all" ? "all" : "mentions";
    } catch {
      getLogger().warn("Teams listening lookup failed");
      return "mentions";
    }
  };
}

/**
 * Tells if a group message starts a run for one agent.
 * - An agent's own bot runs only that agent.
 * - A message that @-mentions the bot is addressed to every agent.
 * - Where the platform needs a mention, a message without one runs an agent only when
 *   Managebot sets listening `all` for that (chat, agent). With no lookup, it does not run.
 */
export async function groupMessageRunsAgent(
  event: Pick<
    MessagingInboundMessage,
    "provider" | "threadId" | "mentionsBot" | "groupMentionRequired" | "agentId"
  >,
  agentId: string,
  lookup: TeamsListeningLookup | undefined,
): Promise<boolean> {
  if (event.agentId && event.agentId !== agentId) return false;
  if (!event.groupMentionRequired || event.mentionsBot) return true;
  if (!lookup || !event.provider.startsWith("teams")) return false;
  const conversationId = teamsConversationIdOfThread(event.threadId);
  if (!conversationId) return false;
  return (await lookup({ conversationId, agentId })) === "all";
}
