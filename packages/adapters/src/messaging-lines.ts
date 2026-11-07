import type { AdapterContext, JobPublisher, MessagingSurface } from "@rakazo/adapter-kit";
import { messagingDeliverJob } from "@rakazo/adapter-kit";
import type { PrismaClient } from "@rakazo/db";
import { recordAuditEvent } from "@rakazo/db";
import { getLogger } from "@rakazo/logging";
import { MessagingConversationGoneError } from "./chat-sdk-surface.js";

/** A line with the identity its replies and runs go through. */
export interface MessagingLineRow {
  id: string;
  identityId: string;
  provider: string;
  threadId: string;
  ownerAddress: string;
  suspendedAt: Date | null;
  identity: { userId: string; spaceId: string; botId: string };
}

/** The one message a line posts when it stops answering. */
export const MESSAGING_LINE_PAUSED_NOTICE =
  "Paused. Someone besides you is in this chat, or I couldn't check who is, so I've stopped answering here. Once it's only you and me again, send \"resume\".";

export type MessagingLineMembers = "owner" | "others" | "gone" | "unavailable";

/** How long a line check waits for the conversation's member list. */
const LINE_MEMBERS_TIMEOUT_MS = 15_000;

/**
 * Who is in a line's conversation besides the bot: exactly its owner, anyone else, nobody we
 * can reach any more (the bot was removed or the conversation deleted), or unknown.
 */
export async function messagingLineMembers(
  messaging: Pick<MessagingSurface, "conversationMembers">,
  line: Pick<MessagingLineRow, "id" | "threadId" | "ownerAddress">,
): Promise<MessagingLineMembers> {
  if (!messaging.conversationMembers) return "unavailable";
  try {
    const members = await messaging.conversationMembers(line.threadId, lineContext(line));
    if (!members) return "unavailable";
    return members.length === 1 && members[0] === line.ownerAddress ? "owner" : "others";
  } catch (error) {
    if (error instanceof MessagingConversationGoneError) return "gone";
    getLogger().warn("messaging line member list unavailable", {
      lineId: line.id,
      reason: error instanceof Error ? error.message : String(error),
    });
    return "unavailable";
  }
}

function lineContext(line: Pick<MessagingLineRow, "id">): AdapterContext {
  return {
    operationId: "messaging.line.members",
    traceId: `messaging.line.members:${line.id}`,
    spaceId: "",
    userId: "",
    signal: AbortSignal.timeout(LINE_MEMBERS_TIMEOUT_MS),
  };
}

type LineDeps = { prisma: PrismaClient; jobs: Pick<JobPublisher, "enqueue"> };

/**
 * Whether the line may answer now: it is not paused and its owner is alone with the bot.
 * Otherwise it is paused, posting its one notice, or unbound when the conversation is gone,
 * and nothing else is sent.
 */
export async function messagingLineAnswers(
  deps: LineDeps & { messaging: Pick<MessagingSurface, "conversationMembers"> },
  line: MessagingLineRow,
): Promise<boolean> {
  if (line.suspendedAt) return false;
  const members = await messagingLineMembers(deps.messaging, line);
  if (members === "owner") return true;
  if (members === "gone") await unbindMessagingLine(deps, line, "unreachable");
  else await pauseMessagingLine(deps, line, members === "others" ? "members" : "unavailable");
  return false;
}

/** Pauses a line once; only the call that paused it posts the notice and records it. */
export async function pauseMessagingLine(
  deps: LineDeps,
  line: MessagingLineRow,
  reason: "members" | "unavailable" | "member_added",
): Promise<boolean> {
  const now = new Date();
  const paused = await deps.prisma.$transaction(async (tx) => {
    const { count } = await tx.messagingLine.updateMany({
      where: { id: line.id, suspendedAt: null },
      data: { suspendedAt: now },
    });
    if (count === 0) return false;
    await tx.messagingOutbound.createMany({
      data: [
        {
          idempotencyKey: `line-paused:${line.id}:${now.getTime()}`,
          kind: "group",
          threadId: line.threadId,
          body: MESSAGING_LINE_PAUSED_NOTICE,
        },
      ],
      skipDuplicates: true,
    });
    return true;
  });
  if (!paused) return false;
  line.suspendedAt = now;
  await auditMessagingLine("messaging.line.suspended", line, { reason });
  await enqueueDelivery(deps);
  return true;
}

/**
 * Ends a line: its identity goes, and the line with it. `reply` is posted to the conversation
 * first when it can still be reached.
 */
export async function unbindMessagingLine(
  deps: LineDeps,
  line: MessagingLineRow,
  reason: "owner" | "owner_web" | "bot_removed" | "unreachable",
  reply?: { key: string; body: string },
): Promise<boolean> {
  const unbound = await deps.prisma.$transaction(async (tx) => {
    const { count } = await tx.messagingIdentity.deleteMany({ where: { id: line.identityId } });
    if (count === 0) return false;
    if (reply) {
      await tx.messagingOutbound.createMany({
        data: [
          { idempotencyKey: reply.key, kind: "group", threadId: line.threadId, body: reply.body },
        ],
        skipDuplicates: true,
      });
    }
    return true;
  });
  if (!unbound) return false;
  await auditMessagingLine("messaging.line.unlinked", line, {
    reason,
    actor: reason === "owner" || reason === "owner_web" ? "owner" : "system",
  });
  if (reply) await enqueueDelivery(deps);
  return true;
}

/** Bind, pause, resume and unlink events, with ids and reasons only, never message text. */
export async function auditMessagingLine(
  action:
    | "messaging.line.bound"
    | "messaging.line.suspended"
    | "messaging.line.resumed"
    | "messaging.line.unlinked",
  line: {
    id?: string;
    provider: string;
    identity: { userId: string; spaceId: string; botId: string };
  },
  options: { reason?: string; actor?: "owner" | "system"; outcome?: "success" | "denied" } = {},
): Promise<void> {
  await recordAuditEvent({
    action,
    actor:
      options.actor === "owner"
        ? { type: "user", id: line.identity.userId }
        : { type: "system", label: "messaging" },
    spaceId: line.identity.spaceId,
    target: { type: "messaging_line", id: line.id ?? null },
    outcome: options.outcome ?? "success",
    metadata: { provider: line.provider, botId: line.identity.botId, reason: options.reason },
  });
}

async function enqueueDelivery(deps: Pick<LineDeps, "jobs">): Promise<void> {
  await deps.jobs.enqueue(messagingDeliverJob()).catch((error) => {
    getLogger().error("messaging line notice enqueue error", error);
  });
}
