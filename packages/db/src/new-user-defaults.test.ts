import type { Actor } from "@rakazo/contracts";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { Prisma, PrismaClient } from "./client.js";
import { provisionMessagingIdentity } from "./messaging.js";
import type { NewUserDefaults } from "./new-user-defaults.js";
import { firstBotWithDefaults } from "./new-user-defaults.js";

const mocks = vi.hoisted(() => ({
  createBot: vi.fn(),
  bootstrapUserSpace: vi.fn(async () => ({ spaceId: "space-1" })),
}));
vi.mock("./repos.js", () => ({ createRepos: () => ({ createBot: mocks.createBot }) }));
vi.mock("./bootstrap-user.js", () => ({ bootstrapUserSpace: mocks.bootstrapUserSpace }));

const actor: Actor = {
  userId: "user-1",
  spaceId: "space-1",
  email: "user@rakazo.test",
  isDeploymentOwner: false,
};
const tx = {} as Prisma.TransactionClient;

describe("firstBotWithDefaults", () => {
  const seedFirstBot = vi.fn(async () => undefined);
  const defaults: NewUserDefaults = { botInstructions: "Be brief.", seedFirstBot };

  it("fills empty instructions and seeds the bot as the given actor", async () => {
    const bot = firstBotWithDefaults(actor, { name: "Chief", instructions: "  " }, defaults);
    expect(bot).toMatchObject({ name: "Chief", instructions: "Be brief." });
    await bot.seed?.(tx, "bot-1");
    expect(seedFirstBot).toHaveBeenCalledWith(tx, actor, "bot-1");
  });

  it("keeps instructions the bot was given", () => {
    expect(firstBotWithDefaults(actor, { instructions: "Mine." }, defaults).instructions).toBe(
      "Mine.",
    );
  });

  it("changes nothing without defaults", () => {
    expect(firstBotWithDefaults(actor, { instructions: "" }, undefined)).toEqual({
      instructions: "",
    });
  });
});

describe("provisionMessagingIdentity with new user defaults", () => {
  function prisma() {
    return {
      messagingIdentity: {
        findUnique: vi.fn(async () => null),
        create: vi.fn(async () => ({})),
      },
      user: {
        findUnique: vi.fn(async () => ({ id: "user-1", email: "msg-1@messaging.invalid" })),
      },
      spaceMember: { findFirst: vi.fn(async () => null) },
      bot: { findFirst: vi.fn(async () => null) },
      thread: { findFirst: vi.fn(async () => ({ id: "thread-1" })) },
    } as unknown as PrismaClient;
  }
  const signup = { signupsEnabled: "true", signupAllowlist: "" };
  const request = { provider: "teams", address: "user-aad-1" };

  beforeEach(() => {
    mocks.createBot.mockReset();
    mocks.createBot.mockResolvedValue({ id: "bot-1" });
    mocks.bootstrapUserSpace.mockClear();
  });

  it("gives the Assistant the default instructions, routines, timezone and language", async () => {
    const seedFirstBot = vi.fn(async () => undefined);
    const defaults: NewUserDefaults = {
      botInstructions: "Answer in short paragraphs.",
      timezone: "America/Denver",
      locale: "es",
      seedFirstBot,
    };
    await provisionMessagingIdentity(prisma(), request, signup, defaults);

    expect(mocks.bootstrapUserSpace).toHaveBeenCalledWith(
      expect.anything(),
      expect.objectContaining({ id: "user-1" }),
      signup,
      defaults,
    );
    const [botActor, input] = mocks.createBot.mock.calls[0]!;
    expect(botActor).toMatchObject({ userId: "user-1", spaceId: "space-1" });
    expect(input).toMatchObject({ name: "Assistant", instructions: "Answer in short paragraphs." });
    await input.seed(tx, "bot-1");
    expect(seedFirstBot).toHaveBeenCalledWith(tx, botActor, "bot-1");
  });

  it("keeps the built-in Assistant instructions without defaults", async () => {
    await provisionMessagingIdentity(prisma(), request, signup);
    const [, input] = mocks.createBot.mock.calls[0]!;
    expect(input.instructions).toContain("You are the owner's personal agent.");
    expect(input.seed).toBeUndefined();
  });
});
