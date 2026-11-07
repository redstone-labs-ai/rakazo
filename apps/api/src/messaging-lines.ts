import type { JobPublisher, MessagingInboundMessage, MessagingSurface } from "@rakazo/adapter-kit";
import { messagingDeliverJob } from "@rakazo/adapter-kit";
import type { MessagingLineRow } from "@rakazo/adapters";
import {
  auditMessagingLine,
  messagingLineAnswers,
  messagingLineMembers,
  pauseMessagingLine,
  unbindMessagingLine,
} from "@rakazo/adapters";
import { sanitizeMessagingLabel } from "@rakazo/core";
import type { Prisma, PrismaClient } from "@rakazo/db";
import {
  bindMessagingLine,
  messagingLineClientNonce,
  normalizeMessagingLinkCode,
} from "@rakazo/db";
import { getLogger } from "@rakazo/logging";

/**
 * Lines: a group conversation bound to one bot as its owner's own chat. A line answers only
 * its owner, only while the owner is alone with the bot, and like a direct message, so the
 * owner's full setup applies. It never joins group discovery or fan-out.
 */
export interface MessagingLineDeps {
  prisma: PrismaClient;
  jobs: Pick<JobPublisher, "enqueue">;
  messaging?: Pick<MessagingSurface, "conversationMembers">;
}

/** Owner commands only a line takes. */
export type MessagingLineCommand = "resume" | "unlink";

export function parseMessagingLineCommand(text: string): MessagingLineCommand | null {
  const normalized = text
    .trim()
    .toLowerCase()
    .replace(/^\//, "")
    .replace(/[.!]+$/, "");
  return normalized === "resume" || normalized === "unlink" ? normalized : null;
}

export async function findMessagingLine(
  prisma: PrismaClient,
  threadId: string,
): Promise<MessagingLineRow | null> {
  return prisma.messagingLine.findUnique({
    where: { threadId },
    include: { identity: { select: { userId: true, spaceId: true, botId: true } } },
  });
}

/**
 * Inbound on a line. Membership events pause it when someone is added and unbind it when the
 * bot is removed. Messages from anyone but the owner are dropped and pause it, since someone
 * else is evidently there. The owner's messages are checked against the member list before
 * they are delivered; `deliver` hands them to the bot like a direct message, keyed to the
 * line so the run is the owner's only while the line stays bound and unpaused.
 */
export async function handleMessagingLineEvent(
  deps: MessagingLineDeps,
  event: MessagingInboundMessage,
  line: MessagingLineRow,
  deliver: (line: MessagingLineRow, clientNonce: string) => Promise<void>,
): Promise<void> {
  if (event.departures?.botLeft) {
    await unbindMessagingLine(deps, line, "bot_removed");
    return;
  }
  if (event.arrivals?.length) {
    await pauseMessagingLine(deps, line, "member_added");
    return;
  }
  if (event.departures) return;
  if (event.from !== line.ownerAddress || event.senderIsBot) {
    await pauseMessagingLine(deps, line, "members");
    return;
  }
  const text = event.contentWithoutMention ?? event.content;
  const command = parseMessagingLineCommand(text);
  const bot = () => botName(deps.prisma, line.identity.botId);
  if (command === "unlink") {
    await unbindMessagingLine(deps, line, "owner", {
      key: `line-unlinked:${line.id}`,
      body: `Unlinked. Messages here no longer reach "${await bot()}".`,
    });
    return;
  }
  if (command === "resume" && line.suspendedAt) {
    const members = await messagingLineMembers(deps.messaging ?? {}, line);
    if (members === "gone") await unbindMessagingLine(deps, line, "unreachable");
    // Still not alone with the bot: the line stays paused and says nothing more.
    if (members !== "owner") return;
    const { count } = await deps.prisma.messagingLine.updateMany({
      where: { id: line.id, suspendedAt: { not: null } },
      data: { suspendedAt: null },
    });
    if (count === 0) return;
    await auditMessagingLine("messaging.line.resumed", line, { actor: "owner" });
    await postToLine(deps, line, `line-resumed:${line.id}:${event.handle}`, {
      body: `Resumed. Messages here reach "${await bot()}" again.`,
    });
    return;
  }
  // A fresh group code from the owner re-points the line at another of their bots.
  if (normalizeMessagingLinkCode(text) && (await tryBindMessagingLine(deps, event))) return;
  if (!(await messagingLineAnswers({ ...deps, messaging: deps.messaging ?? {} }, line))) return;
  await deliver(line, messagingLineClientNonce(line.id, event.provider, event.handle));
}

/**
 * A group link code posted in a conversation. Returns false when the text names no code, so
 * the message goes on as an ordinary group message; any real code is spent either way.
 */
export async function tryBindMessagingLine(
  deps: MessagingLineDeps,
  event: MessagingInboundMessage,
): Promise<boolean> {
  const code = normalizeMessagingLinkCode(event.contentWithoutMention ?? event.content);
  if (!code || event.senderIsBot) return false;
  const name = event.channelName ? sanitizeMessagingLabel(event.channelName) : null;
  const result = await bindMessagingLine(
    deps.prisma,
    {
      code,
      provider: event.provider,
      threadId: event.threadId,
      ownerAddress: event.from,
      name,
      directoryUser: event.directoryUser,
      ownerAlone: async () =>
        (await messagingLineMembers(deps.messaging ?? {}, {
          id: event.threadId,
          threadId: event.threadId,
          ownerAddress: event.from,
        })) === "owner",
    },
    async (tx, bound) => {
      await writeLineReply(
        tx,
        event.threadId,
        `line-bound:${bound.lineId}:${event.handle}`,
        `Linked. This chat now reaches "${await botName(tx, bound.botId)}" like a direct chat. I answer only while you're the only person here, and anyone added later may be able to see what's already been said.`,
      );
    },
  );
  if (result.status === "bound") {
    await auditMessagingLine(
      "messaging.line.bound",
      { id: result.lineId, provider: event.provider, identity: result },
      { actor: "owner" },
    );
    await enqueueDelivery(deps);
    return true;
  }
  if (result.reason === "code") return false;
  getLogger().info("messaging line not bound", { provider: event.provider, reason: result.reason });
  const issuer = result.issuer;
  if (!issuer) return true;
  await auditMessagingLine(
    "messaging.line.bound",
    { provider: event.provider, identity: issuer },
    // A sender who is not the code's issuer is someone else holding a leaked code.
    {
      actor: result.reason === "directory" ? "system" : "owner",
      outcome: "denied",
      reason: result.reason,
    },
  );
  // Only the code's own issuer, verified by their directory account, hears why.
  const reply =
    result.reason === "bot_linked"
      ? `"${await botName(deps.prisma, issuer.botId)}" already answers another chat. Unlink it there first, then send a new code.`
      : result.reason === "members"
        ? "Not linked: someone besides you is in this chat. A group chat code links only a chat with just you and me."
        : null;
  if (reply) {
    await writeLineReply(deps.prisma, event.threadId, `line-refused:${event.handle}`, reply);
    await enqueueDelivery(deps);
  }
  return true;
}

async function postToLine(
  deps: MessagingLineDeps,
  line: MessagingLineRow,
  key: string,
  reply: { body: string },
): Promise<void> {
  await writeLineReply(deps.prisma, line.threadId, key, reply.body);
  await enqueueDelivery(deps);
}

/** Line notices post straight to the conversation, outside the line's reply check. */
async function writeLineReply(
  db: Pick<Prisma.TransactionClient, "messagingOutbound">,
  threadId: string,
  key: string,
  body: string,
): Promise<void> {
  await db.messagingOutbound.createMany({
    data: [{ idempotencyKey: key, kind: "group", threadId, body }],
    skipDuplicates: true,
  });
}

async function botName(db: Pick<Prisma.TransactionClient, "bot">, botId: string) {
  const bot = await db.bot.findUnique({ where: { id: botId }, select: { name: true } });
  return bot?.name ?? "your agent";
}

async function enqueueDelivery(deps: Pick<MessagingLineDeps, "jobs">): Promise<void> {
  await deps.jobs.enqueue(messagingDeliverJob()).catch((error) => {
    getLogger().error("messaging line reply enqueue error", error);
  });
}
