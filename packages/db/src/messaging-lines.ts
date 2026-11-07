import type { Prisma, PrismaClient } from "./client.js";

/**
 * A group conversation bound to one bot as its owner's own chat. Its messaging identity has
 * the address `<owner address>#<thread id>`, so it never matches a person's 1:1 address, and
 * `botId` stays unique: a bot answers either someone's 1:1 chat or one line.
 */
export function messagingLineAddress(ownerAddress: string, threadId: string): string {
  return `${ownerAddress}#${threadId}`;
}

const LINE_NONCE_PREFIX = "messaging-line:";

/** Client nonce of a message an owner sent on a line: `messaging-line:<line>:<provider>:<id>`. */
export function messagingLineClientNonce(lineId: string, provider: string, handle: string) {
  return `${LINE_NONCE_PREFIX}${lineId}:${provider}:${handle}`;
}

/**
 * Whether a run woken by a line message still answers its owner: null when the message did
 * not come from a line; false once that line is paused or unbound.
 */
export async function messagingLineRunFromOwner(
  prisma: Pick<PrismaClient, "messagingLine">,
  sourceClientNonce: string | null | undefined,
): Promise<boolean | null> {
  if (!sourceClientNonce?.startsWith(LINE_NONCE_PREFIX)) return null;
  const lineId = sourceClientNonce.slice(LINE_NONCE_PREFIX.length).split(":", 1)[0];
  if (!lineId) return false;
  const line = await prisma.messagingLine.findUnique({
    where: { id: lineId },
    select: { suspendedAt: true },
  });
  return Boolean(line && !line.suspendedAt);
}

export type MessagingLineBinding =
  | {
      status: "bound";
      lineId: string;
      identityId: string;
      userId: string;
      spaceId: string;
      botId: string;
    }
  | {
      status: "refused";
      /**
       * code: no unexpired group code; directory: the sender is not the code's issuer;
       * members: someone besides the owner is in the conversation; bot_linked: the bot already
       * answers a chat app; taken: another person's line.
       */
      reason: "code" | "directory" | "members" | "bot_linked" | "taken";
      /** The issuer, once the code was valid. */
      issuer?: { userId: string; spaceId: string; botId: string };
    };

/**
 * Binds a group conversation to the bot a group link code names. The code is spent by any
 * attempt, since everyone in the conversation saw it. It binds only when the sender's verified
 * directory account is the issuer's own SSO account and the owner is alone with the bot; a
 * leaked code alone does nothing. Binding the issuer's own line again re-points it and lifts a
 * pause.
 */
export async function bindMessagingLine(
  prisma: PrismaClient,
  request: {
    code: string;
    provider: string;
    threadId: string;
    ownerAddress: string;
    name: string | null;
    directoryUser: { tenantId: string; objectId: string } | undefined;
    /** Whether the conversation's members, without the bot, are exactly the sender. */
    ownerAlone: () => Promise<boolean>;
  },
  confirm?: (
    tx: Prisma.TransactionClient,
    bound: Extract<MessagingLineBinding, { status: "bound" }>,
  ) => Promise<void>,
): Promise<MessagingLineBinding> {
  const row = await prisma.messagingLinkCode.findUnique({ where: { code: request.code } });
  if (!row) return { status: "refused", reason: "code" };
  const { count } = await prisma.messagingLinkCode.deleteMany({ where: { id: row.id } });
  if (count === 0 || !row.forGroup || row.expiresAt.getTime() < Date.now())
    return { status: "refused", reason: "code" };
  const issuer = { userId: row.userId, spaceId: row.spaceId, botId: row.botId };
  const directory = request.directoryUser
    ? await prisma.directoryIdentity.findUnique({
        where: {
          tenantId_objectId: {
            tenantId: request.directoryUser.tenantId.toLowerCase(),
            objectId: request.directoryUser.objectId.toLowerCase(),
          },
        },
        select: { userId: true },
      })
    : null;
  if (directory?.userId !== row.userId) return { status: "refused", reason: "directory", issuer };
  if (!(await request.ownerAlone())) return { status: "refused", reason: "members", issuer };
  try {
    return await prisma.$transaction(async (tx) => {
      const existing = await tx.messagingLine.findUnique({
        where: { threadId: request.threadId },
        include: { identity: { select: { userId: true } } },
      });
      let bound: Extract<MessagingLineBinding, { status: "bound" }>;
      if (existing) {
        if (existing.identity.userId !== row.userId)
          return { status: "refused", reason: "taken", issuer } as const;
        await tx.messagingIdentity.update({
          where: { id: existing.identityId },
          data: { botId: row.botId, spaceId: row.spaceId },
        });
        await tx.messagingLine.update({
          where: { id: existing.id },
          data: { suspendedAt: null, ...(request.name ? { name: request.name } : {}) },
        });
        bound = {
          status: "bound",
          lineId: existing.id,
          identityId: existing.identityId,
          ...issuer,
        };
      } else {
        const identity = await tx.messagingIdentity.create({
          data: {
            provider: request.provider,
            address: messagingLineAddress(request.ownerAddress, request.threadId),
            dmThreadId: request.threadId,
            userId: row.userId,
            spaceId: row.spaceId,
            botId: row.botId,
            verifiedAt: new Date(),
          },
        });
        const line = await tx.messagingLine.create({
          data: {
            identityId: identity.id,
            provider: request.provider,
            threadId: request.threadId,
            ownerAddress: request.ownerAddress,
            name: request.name,
          },
        });
        bound = { status: "bound", lineId: line.id, identityId: identity.id, ...issuer };
      }
      await confirm?.(tx, bound);
      return bound;
    });
  } catch (error) {
    // botId is unique: the bot already answers someone's 1:1 chat or another line.
    if (error && typeof error === "object" && "code" in error && error.code === "P2002")
      return { status: "refused", reason: "bot_linked", issuer };
    throw error;
  }
}
