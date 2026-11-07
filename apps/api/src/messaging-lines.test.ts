import type { MessagingInboundMessage } from "@rakazo/adapter-kit";
import { MessagingConversationGoneError } from "@rakazo/adapters";
import { installAuditEventStore } from "@rakazo/db";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { MessagingInboundDeps } from "./messaging-inbound.js";
import { createMessagingInboundHandler } from "./messaging-inbound.js";
import { parseMessagingLineCommand } from "./messaging-lines.js";

const OWNER = "29:owner";
const THREAD = "teams:chat:service";
const TENANT = "tenant-1";

type Line = {
  id: string;
  identityId: string;
  provider: string;
  threadId: string;
  ownerAddress: string;
  name: string | null;
  suspendedAt: Date | null;
};
type Identity = {
  id: string;
  provider: string;
  address: string;
  dmThreadId: string | null;
  userId: string;
  spaceId: string;
  botId: string;
};

/** An in-memory store with just the tables lines touch, and a member list per conversation. */
function lineWorld(
  options: {
    line?: Partial<Line> | null;
    members?: string[] | Error;
    code?: { forGroup?: boolean; botId?: string; expired?: boolean } | null;
    directoryOwner?: string;
    linkedBotIds?: string[];
    groupMembers?: Array<Record<string, unknown>>;
  } = {},
) {
  const identities: Identity[] = [];
  const lines: Line[] = [];
  if (options.line !== null) {
    identities.push({
      id: "mi-line",
      provider: "teams",
      address: `${OWNER}#${THREAD}`,
      dmThreadId: THREAD,
      userId: "user-1",
      spaceId: "ws-1",
      botId: "bot-1",
    });
    lines.push({
      id: "line-1",
      identityId: "mi-line",
      provider: "teams",
      threadId: THREAD,
      ownerAddress: OWNER,
      name: "Chief only",
      suspendedAt: null,
      ...options.line,
    });
  }
  const codes =
    options.code === null || options.code === undefined
      ? []
      : [
          {
            id: "code-1",
            code: "ABCD2345",
            userId: "user-1",
            spaceId: "ws-1",
            botId: options.code.botId ?? "bot-2",
            forGroup: options.code.forGroup ?? true,
            expiresAt: new Date(Date.now() + (options.code.expired ? -1 : 1) * 60_000),
          },
        ];
  const outbound: Array<Record<string, unknown>> = [];
  const groupMembers = options.groupMembers ?? [];
  const audit: Array<Record<string, unknown>> = [];
  const withIdentity = (line: Line | undefined) => {
    if (!line) return null;
    const identity = identities.find((row) => row.id === line.identityId);
    return identity ? { ...line, identity } : null;
  };
  const prisma = {
    messagingLine: {
      findUnique: vi.fn(async ({ where }: { where: { threadId?: string; identityId?: string } }) =>
        withIdentity(
          lines.find(
            (line) =>
              (where.threadId === undefined || line.threadId === where.threadId) &&
              (where.identityId === undefined || line.identityId === where.identityId),
          ),
        ),
      ),
      updateMany: vi.fn(
        async ({
          where,
          data,
        }: {
          where: { id: string; suspendedAt: null | { not: null } };
          data: Partial<Line>;
        }) => {
          const line = lines.find((row) => row.id === where.id);
          const paused = Boolean(line?.suspendedAt);
          if (!line || (where.suspendedAt === null ? paused : !paused)) return { count: 0 };
          Object.assign(line, data);
          return { count: 1 };
        },
      ),
      update: vi.fn(async ({ where, data }: { where: { id: string }; data: Partial<Line> }) => {
        const line = lines.find((row) => row.id === where.id)!;
        Object.assign(line, data);
        return line;
      }),
      create: vi.fn(async ({ data }: { data: Omit<Line, "id" | "suspendedAt"> }) => {
        const line = { id: `line-${lines.length + 1}`, suspendedAt: null, ...data };
        lines.push(line);
        return line;
      }),
    },
    messagingIdentity: {
      // Ordinary group members' identities resolve by id for the fan-out; no 1:1 sender is linked.
      findUnique: vi.fn(async ({ where }: { where: { id?: string } }) =>
        where.id
          ? { id: where.id, userId: "user-9", spaceId: "ws-9", botId: `bot-${where.id}` }
          : null,
      ),
      update: vi.fn(async ({ where, data }: { where: { id: string }; data: object }) => {
        const identity = identities.find((row) => row.id === where.id)!;
        Object.assign(identity, data);
        return identity;
      }),
      create: vi.fn(async ({ data }: { data: Omit<Identity, "id"> }) => {
        if (
          (options.linkedBotIds ?? []).includes(data.botId) ||
          identities.some((row) => row.botId === data.botId)
        )
          throw Object.assign(new Error("Unique constraint failed on botId"), { code: "P2002" });
        const identity = { id: `mi-${identities.length + 1}`, ...data };
        identities.push(identity);
        return identity;
      }),
      deleteMany: vi.fn(async ({ where }: { where: { id: string } }) => {
        const index = identities.findIndex((row) => row.id === where.id);
        if (index < 0) return { count: 0 };
        identities.splice(index, 1);
        // The line goes with its identity.
        const line = lines.findIndex((row) => row.identityId === where.id);
        if (line >= 0) lines.splice(line, 1);
        return { count: 1 };
      }),
    },
    messagingLinkCode: {
      findUnique: vi.fn(
        async ({ where }: { where: { code: string } }) =>
          codes.find((row) => row.code === where.code) ?? null,
      ),
      deleteMany: vi.fn(async ({ where }: { where: { id: string } }) => {
        const index = codes.findIndex((row) => row.id === where.id);
        if (index < 0) return { count: 0 };
        codes.splice(index, 1);
        return { count: 1 };
      }),
    },
    directoryIdentity: {
      findUnique: vi.fn(
        async ({
          where,
        }: {
          where: { tenantId_objectId: { tenantId: string; objectId: string } };
        }) =>
          where.tenantId_objectId.tenantId === TENANT &&
          where.tenantId_objectId.objectId === "oid-owner"
            ? { userId: options.directoryOwner ?? "user-1" }
            : null,
      ),
    },
    messagingOutbound: {
      createMany: vi.fn(async ({ data }: { data: Array<Record<string, unknown>> }) => {
        for (const row of data)
          if (!outbound.some((existing) => existing.idempotencyKey === row.idempotencyKey))
            outbound.push(row);
        return { count: data.length };
      }),
    },
    bot: { findUnique: vi.fn(async () => ({ name: "Chief" })) },
    thread: { findFirst: vi.fn(async () => ({ id: "thread-1" })) },
    routine: { findMany: vi.fn(async () => []) },
    messagingChannel: {
      upsert: vi.fn(async () => ({
        id: "ch-1",
        provider: "teams",
        threadId: THREAD,
        name: "Ops",
        introPostedAt: new Date(),
      })),
    },
    messagingChannelMember: {
      findUnique: vi.fn(
        async ({ where }: { where: { channelId_address: { address: string } } }) =>
          groupMembers.find((row) => row.address === where.channelId_address.address) ?? null,
      ),
      findMany: vi.fn(async () => groupMembers.filter((row) => row.status === "approved")),
      upsert: vi.fn(async () => ({})),
      updateMany: vi.fn(async () => ({ count: 0 })),
    },
    user: { findUnique: vi.fn(async () => ({ name: "Ada Owner" })) },
    $transaction: vi.fn(async (work: (tx: unknown) => Promise<unknown>) => work(prisma)),
  };
  const conversationMembers = vi.fn(async () => {
    if (options.members instanceof Error) throw options.members;
    return options.members ?? [OWNER];
  });
  const sendUserMessage = vi.fn(async () => ({ messageId: "msg-1", runId: "run-1", seq: 1 }));
  const enqueue = vi.fn(async () => undefined);
  installAuditEventStore({
    auditEvent: {
      createMany: vi.fn(async ({ data }: { data: Array<Record<string, unknown>> }) => {
        audit.push(...data);
        return { count: data.length };
      }),
    },
    space: { findUnique: vi.fn(async () => ({ organizationId: "org-1", name: "Home" })) },
    user: { findUnique: vi.fn(async () => ({ email: "owner@example.test" })) },
    member: { findMany: vi.fn(async () => []) },
  } as unknown as Parameters<typeof installAuditEventStore>[0]);
  const deps = {
    prisma,
    events: { sendUserMessage, notify: vi.fn(async () => undefined) },
    jobs: { enqueue },
    provision: vi.fn(),
    openSignup: false,
    signupPolicy: { signupsEnabled: undefined, signupAllowlist: undefined },
    messaging: { conversationMembers },
  } as unknown as MessagingInboundDeps;
  return {
    deps,
    handle: createMessagingInboundHandler(deps),
    lines,
    identities,
    codes,
    outbound,
    audit,
    sendUserMessage,
    conversationMembers,
  };
}

function lineMessage(overrides: Partial<MessagingInboundMessage> = {}): MessagingInboundMessage {
  return {
    type: "message",
    provider: "teams",
    handle: "h-1",
    threadId: THREAD,
    isDirect: false,
    from: OWNER,
    fromLabel: "Ada Owner",
    channelName: "Chief only",
    participants: [OWNER],
    content: "What is due today?",
    mediaUrl: null,
    groupMentionRequired: true,
    directoryUser: { tenantId: TENANT, objectId: "oid-owner" },
    ...overrides,
  };
}

const PAUSED =
  "Paused. Someone besides you is in this chat, or I couldn't check who is, so I've stopped answering here. Once it's only you and me again, send \"resume\".";

afterEach(() => installAuditEventStore(undefined));

describe("a bound line", () => {
  it("answers every owner message like a direct chat, mention or not", async () => {
    const world = lineWorld();
    await world.handle(lineMessage());

    expect(world.sendUserMessage).toHaveBeenCalledWith(
      expect.objectContaining({
        botId: "bot-1",
        userId: "user-1",
        trigger: "messaging",
        blocks: [{ kind: "text", text: "What is due today?" }],
        clientNonce: "messaging-line:line-1:teams:h-1",
      }),
    );
    expect(world.conversationMembers).toHaveBeenCalledWith(THREAD, expect.anything());
    expect(world.deps.prisma.messagingChannel.upsert).not.toHaveBeenCalled();
  });

  it("pauses once, with one notice and no run, when a second person is in the chat", async () => {
    const world = lineWorld({ members: [OWNER, "29:grace"] });
    await world.handle(lineMessage());
    await world.handle(lineMessage({ handle: "h-2" }));

    expect(world.sendUserMessage).not.toHaveBeenCalled();
    expect(world.lines[0]?.suspendedAt).toBeInstanceOf(Date);
    expect(world.outbound).toEqual([
      expect.objectContaining({ kind: "group", threadId: THREAD, body: PAUSED }),
    ]);
    expect(world.audit).toEqual([
      expect.objectContaining({
        action: "messaging.line.suspended",
        targetId: "line-1",
        metadata: expect.objectContaining({ reason: "members" }),
      }),
    ]);
    expect(JSON.stringify(world.audit)).not.toContain("What is due today?");
  });

  it("pauses when the member list cannot be read", async () => {
    const world = lineWorld({ members: new Error("status 500") });
    await world.handle(lineMessage());
    expect(world.sendUserMessage).not.toHaveBeenCalled();
    expect(world.lines[0]?.suspendedAt).toBeInstanceOf(Date);
  });

  it("pauses as soon as Teams reports someone added", async () => {
    const world = lineWorld();
    await world.handle(lineMessage({ from: "", content: "", arrivals: ["29:grace"] }));
    expect(world.lines[0]?.suspendedAt).toBeInstanceOf(Date);
    expect(world.outbound).toHaveLength(1);
    expect(world.conversationMembers).not.toHaveBeenCalled();
  });

  it("drops a message from anyone but the owner, and pauses", async () => {
    const world = lineWorld();
    await world.handle(lineMessage({ from: "29:grace", content: "What did Ada ask you?" }));
    expect(world.sendUserMessage).not.toHaveBeenCalled();
    expect(world.lines[0]?.suspendedAt).toBeInstanceOf(Date);
  });

  it("delivers nothing while paused, and resumes only once the owner is alone again", async () => {
    const world = lineWorld({ line: { suspendedAt: new Date() }, members: [OWNER, "29:grace"] });
    await world.handle(lineMessage({ content: "What is due today?" }));
    await world.handle(lineMessage({ handle: "h-2", content: "resume" }));
    expect(world.lines[0]?.suspendedAt).toBeInstanceOf(Date);
    expect(world.sendUserMessage).not.toHaveBeenCalled();
    expect(world.outbound).toEqual([]);

    world.conversationMembers.mockResolvedValue([OWNER]);
    await world.handle(
      lineMessage({ handle: "h-3", content: "@Rakazo resume", contentWithoutMention: "resume" }),
    );
    expect(world.lines[0]?.suspendedAt).toBeNull();
    expect(world.outbound).toEqual([
      expect.objectContaining({ body: 'Resumed. Messages here reach "Chief" again.' }),
    ]);
    expect(world.audit.map((row) => row.action)).toEqual(["messaging.line.resumed"]);
    await world.handle(lineMessage({ handle: "h-4" }));
    expect(world.sendUserMessage).toHaveBeenCalledTimes(1);
  });

  it("unlinks on the owner's unlink, posting one reply to the chat", async () => {
    const world = lineWorld();
    await world.handle(lineMessage({ content: "unlink" }));
    expect(world.lines).toEqual([]);
    expect(world.identities).toEqual([]);
    expect(world.outbound).toEqual([
      expect.objectContaining({
        kind: "group",
        threadId: THREAD,
        body: 'Unlinked. Messages here no longer reach "Chief".',
      }),
    ]);
    expect(world.audit).toEqual([
      expect.objectContaining({ action: "messaging.line.unlinked", actorType: "user" }),
    ]);
  });

  it("unbinds when the bot is removed or Teams no longer has the chat", async () => {
    const removed = lineWorld();
    await removed.handle(
      lineMessage({ from: "", content: "", departures: { addresses: [], botLeft: true } }),
    );
    expect(removed.lines).toEqual([]);

    const gone = lineWorld({ members: new MessagingConversationGoneError("status 403") });
    await gone.handle(lineMessage());
    expect(gone.lines).toEqual([]);
    expect(gone.sendUserMessage).not.toHaveBeenCalled();
    expect(gone.audit).toEqual([
      expect.objectContaining({
        action: "messaging.line.unlinked",
        metadata: expect.objectContaining({ reason: "unreachable" }),
      }),
    ]);
  });
});

describe("binding a line", () => {
  const codeMessage = (overrides: Partial<MessagingInboundMessage> = {}) =>
    lineMessage({
      content: "@Rakazo ABCD-2345",
      contentWithoutMention: "ABCD-2345",
      mentionsBot: true,
      ...overrides,
    });

  it("binds a group chat when the issuer posts their code alone with the bot", async () => {
    const world = lineWorld({ line: null, code: {} });
    await world.handle(codeMessage());

    expect(world.lines).toEqual([
      expect.objectContaining({ threadId: THREAD, ownerAddress: OWNER, suspendedAt: null }),
    ]);
    expect(world.identities).toEqual([
      expect.objectContaining({
        address: `${OWNER}#${THREAD}`,
        dmThreadId: THREAD,
        userId: "user-1",
        botId: "bot-2",
      }),
    ]);
    expect(world.outbound).toEqual([
      expect.objectContaining({
        kind: "group",
        threadId: THREAD,
        body: expect.stringContaining('Linked. This chat now reaches "Chief"'),
      }),
    ]);
    expect(world.codes).toEqual([]);
    expect(world.audit).toEqual([
      expect.objectContaining({ action: "messaging.line.bound", outcome: "success" }),
    ]);
    expect(world.sendUserMessage).not.toHaveBeenCalled();
  });

  it("does nothing with a leaked code posted by someone else, and spends it", async () => {
    const stranger = lineWorld({ line: null, code: {}, members: ["29:mallory"] });
    await stranger.handle(
      codeMessage({
        from: "29:mallory",
        directoryUser: { tenantId: TENANT, objectId: "oid-mallory" },
      }),
    );
    const unverified = lineWorld({ line: null, code: {}, members: ["29:mallory"] });
    await unverified.handle(codeMessage({ from: "29:mallory", directoryUser: undefined }));

    for (const world of [stranger, unverified]) {
      expect(world.lines).toEqual([]);
      expect(world.outbound).toEqual([]);
      expect(world.codes).toEqual([]);
      // The member list is never read for someone who is not the code's issuer.
      expect(world.conversationMembers).not.toHaveBeenCalled();
      expect(world.audit).toEqual([
        expect.objectContaining({
          action: "messaging.line.bound",
          outcome: "denied",
          actorType: "system",
          metadata: expect.objectContaining({ reason: "directory" }),
        }),
      ]);
    }
  });

  it("refuses a directory account that belongs to someone other than the issuer", async () => {
    const world = lineWorld({ line: null, code: {}, directoryOwner: "user-2" });
    await world.handle(codeMessage());
    expect(world.lines).toEqual([]);
    expect(world.outbound).toEqual([]);
  });

  it("refuses, saying why, when someone else is in the chat", async () => {
    const world = lineWorld({ line: null, code: {}, members: [OWNER, "29:grace"] });
    await world.handle(codeMessage());
    expect(world.lines).toEqual([]);
    expect(world.outbound).toEqual([
      expect.objectContaining({
        body: "Not linked: someone besides you is in this chat. A group chat code links only a chat with just you and me.",
      }),
    ]);
  });

  it("explains that a bot answering another chat can't also be a line", async () => {
    const world = lineWorld({ line: null, code: {}, linkedBotIds: ["bot-2"] });
    await world.handle(codeMessage());
    expect(world.lines).toEqual([]);
    expect(world.outbound).toEqual([
      expect.objectContaining({
        body: '"Chief" already answers another chat. Unlink it there first, then send a new code.',
      }),
    ]);
  });

  it("binds nothing with a direct chat's code, which a group spends, or with text that is no code", async () => {
    const direct = lineWorld({ line: null, code: { forGroup: false } });
    await direct.handle(codeMessage());
    expect(direct.lines).toEqual([]);
    expect(direct.codes).toEqual([]);
    expect(direct.deps.prisma.messagingChannel.upsert).toHaveBeenCalled();

    const none = lineWorld({ line: null });
    await none.handle(codeMessage());
    expect(none.lines).toEqual([]);
    expect(none.deps.prisma.messagingChannel.upsert).toHaveBeenCalled();
  });
});

describe("ordinary groups where every message arrives", () => {
  const approved = [
    { id: "cm-1", address: OWNER, identityId: "mi-a", status: "approved" },
    { id: "cm-2", address: "29:grace", identityId: "mi-b", status: "approved" },
  ];

  beforeEach(() => vi.clearAllMocks());

  it("starts no run for a message that does not @mention the bot", async () => {
    const world = lineWorld({ line: null, groupMembers: approved.map((row) => ({ ...row })) });
    await world.handle(lineMessage({ participants: [OWNER, "29:grace"] }));
    expect(world.sendUserMessage).not.toHaveBeenCalled();
  });

  it("fans a message that @mentions the bot out to every approved member's bot", async () => {
    const world = lineWorld({ line: null, groupMembers: approved.map((row) => ({ ...row })) });
    await world.handle(
      lineMessage({ participants: [OWNER, "29:grace"], mentionsBot: true, content: "@Rakazo hi" }),
    );
    expect(world.sendUserMessage).toHaveBeenCalledTimes(2);
  });
});

describe("parseMessagingLineCommand", () => {
  it("takes only a bare resume or unlink", () => {
    expect(parseMessagingLineCommand(" Resume. ")).toBe("resume");
    expect(parseMessagingLineCommand("/unlink")).toBe("unlink");
    expect(parseMessagingLineCommand("please resume")).toBeNull();
  });
});
