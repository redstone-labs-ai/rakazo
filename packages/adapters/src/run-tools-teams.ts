import { redactSecrets } from "@rakazo/core";
import { getLogger } from "@rakazo/logging";
import { z } from "zod";
import type { RunContext } from "./run-context.js";
import type { ToolCall, ToolHandlers } from "./run-tool-gate.js";
import { TEAMS_GROUP_CHAT_MAX_MEMBERS, TeamsGroupChatError } from "./teams-group-chats.js";

export const TEAMS_CREATE_GROUP_CHAT_TOOL = "teams_create_group_chat";

/** An Entra object id (GUID) or a user principal name. */
const TeamsMember = z
  .string()
  .trim()
  .regex(
    /^(?:[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}|[^\s@'"]+@[^\s@'"]+\.[^\s@'"]+)$/i,
    "must be an Entra user id or a user principal name",
  );

const CreateTeamsGroupChatSchema = z
  .object({
    members: z.array(TeamsMember).min(2).max(TEAMS_GROUP_CHAT_MAX_MEMBERS),
    topic: z.string().trim().min(1).max(250).optional(),
    message: z.string().trim().min(1).max(4_000),
  })
  .strict();

async function createTeamsGroupChat(ctx: RunContext, call: ToolCall) {
  const { deps, run } = ctx;
  const chats = deps.teamsGroupChats;
  if (!chats) return call.finish({ error: "Teams group chats are not enabled on this server." });
  const parsed = CreateTeamsGroupChatSchema.safeParse(call.args);
  if (!parsed.success)
    return call.finish({
      error: parsed.error.issues
        .map((issue) => `${issue.path.join(".") || "input"}: ${issue.message}`)
        .join(" "),
    });
  const members = [...new Set(parsed.data.members.map((member) => member.toLowerCase()))];
  try {
    const created = await chats.create({
      agentId: run.botId,
      members,
      ...(parsed.data.topic ? { topic: parsed.data.topic } : {}),
      firstMessage: redactSecrets(parsed.data.message, ctx.secrets),
    });
    const channel = await registerChat(ctx, created, parsed.data.topic, true);
    return call.finish({
      ok: true,
      chatId: created.chatId,
      channelId: channel.id,
      provider: created.provider,
      appInstalled: created.installed,
    });
  } catch (error) {
    if (error instanceof TeamsGroupChatError) {
      getLogger().warn("teams group chat create failed", {
        step: error.step,
        status: error.status,
        code: error.code,
      });
      // The chat exists in Teams. Register it, so it is not lost, and give its id back.
      const channel = error.chat
        ? await registerChat(ctx, error.chat, parsed.data.topic, false)
        : undefined;
      return call.finish({
        error: error.message,
        ...(error.chat ? { chatId: error.chat.chatId, provider: error.chat.provider } : {}),
        ...(channel ? { channelId: channel.id } : {}),
        ...(error.code ? { code: error.code } : {}),
        retryable: false,
      });
    }
    throw error;
  }
}

/**
 * Saves the chat as a messaging channel. The provider of the agent's own bot carries its agent
 * id, so inbound messages in this chat start runs only for this agent. When the bot posted the
 * first message, no intro is due.
 */
function registerChat(
  ctx: RunContext,
  chat: { threadId: string; provider: string },
  topic: string | undefined,
  posted: boolean,
) {
  return ctx.deps.prisma.messagingChannel.upsert({
    where: { threadId: chat.threadId },
    create: {
      provider: chat.provider,
      threadId: chat.threadId,
      name: topic ?? null,
      ...(posted ? { introPostedAt: new Date() } : {}),
    },
    update: {},
    select: { id: true },
  });
}

/** Teams group chats an agent's own bot creates. */
export const teamsTools: ToolHandlers = { [TEAMS_CREATE_GROUP_CHAT_TOOL]: createTeamsGroupChat };
