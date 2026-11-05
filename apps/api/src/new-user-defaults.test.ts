import { createRouterClient } from "@orpc/server";
import type { Actor } from "@rakazo/contracts";
import { FIRST_BOT_SPAWN_KEY } from "@rakazo/contracts";
import type { NewUserDefaults, Prisma } from "@rakazo/db";
import { describe, expect, it, vi } from "vitest";
import { loadEnv } from "./env.js";
import { newUserDefaultsFromEnv } from "./new-user-defaults.js";
import type { RouterDeps } from "./router.js";
import { createRouter } from "./router.js";

const createBot = vi.hoisted(() => vi.fn());
vi.mock("@rakazo/db", async (original) => ({
  ...(await original<Record<string, unknown>>()),
  createRepos: () =>
    new Proxy({ createBot }, { get: (target, key) => target[key as "createBot"] ?? vi.fn() }),
}));

const actor: Actor = {
  spaceId: "space-1",
  userId: "user-1",
  email: "user@rakazo.test",
  isDeploymentOwner: false,
};

const ROUTINES = JSON.stringify([
  { name: "Morning brief", prompt: "Summarize my day.", crons: ["0 8 * * 1-5"] },
  {
    name: "Weekly review",
    prompt: "Review the week.",
    crons: ["0 16 * * 5"],
    timezone: "Europe/London",
  },
]);

/** Runs the seed against a transaction that only records the routines it creates. */
async function seededRoutines(defaults: NewUserDefaults) {
  const create = vi.fn(async (_args: { data: Record<string, unknown> }) => ({}));
  const tx = { routine: { create } } as unknown as Prisma.TransactionClient;
  await defaults.seedFirstBot?.(tx, actor, "bot-1");
  return create.mock.calls.map(([args]) => args.data);
}

describe("newUserDefaultsFromEnv", () => {
  it("reads every default and seeds paused routines in the default timezone", async () => {
    const warn = vi.fn();
    const defaults = newUserDefaultsFromEnv(
      {
        RAKAZO_DEFAULT_BOT_INSTRUCTIONS: " Answer in short paragraphs. ",
        RAKAZO_DEFAULT_TIMEZONE: "America/Denver",
        RAKAZO_DEFAULT_LOCALE: "es",
        RAKAZO_DEFAULT_ROUTINES: ROUTINES,
      },
      warn,
    );
    expect(defaults).toMatchObject({
      botInstructions: "Answer in short paragraphs.",
      timezone: "America/Denver",
      locale: "es",
    });
    expect(await seededRoutines(defaults)).toEqual([
      expect.objectContaining({
        spaceId: "space-1",
        userId: "user-1",
        botId: "bot-1",
        name: "Morning brief",
        crons: ["0 8 * * 1-5"],
        timezone: "America/Denver",
        active: false,
        nextRunAt: null,
      }),
      expect.objectContaining({ name: "Weekly review", timezone: "Europe/London", active: false }),
    ]);
    expect(warn).not.toHaveBeenCalled();
  });

  it("leaves everything unset by default, and routines in UTC without a default timezone", async () => {
    expect(newUserDefaultsFromEnv({})).toEqual({
      timezone: undefined,
      locale: undefined,
      botInstructions: undefined,
      seedFirstBot: undefined,
    });
    const defaults = newUserDefaultsFromEnv({ RAKAZO_DEFAULT_ROUTINES: ROUTINES });
    expect((await seededRoutines(defaults))[0]).toMatchObject({ timezone: "UTC" });
  });

  it.each([
    ["RAKAZO_DEFAULT_TIMEZONE", "Mars/Olympus"],
    ["RAKAZO_DEFAULT_TIMEZONE", "+02:00"],
    ["RAKAZO_DEFAULT_LOCALE", "klingon"],
    ["RAKAZO_DEFAULT_ROUTINES", "{not json"],
    ["RAKAZO_DEFAULT_ROUTINES", '{"name":"Not a list"}'],
    ["RAKAZO_DEFAULT_ROUTINES", '[{"name":"No schedule","prompt":"Hi","crons":[]}]'],
    ["RAKAZO_DEFAULT_ROUTINES", '[{"name":"Bad cron","prompt":"Hi","crons":["every day"]}]'],
    ["RAKAZO_DEFAULT_ROUTINES", '[{"prompt":"No name","crons":["0 8 * * *"]}]'],
    [
      "RAKAZO_DEFAULT_ROUTINES",
      '[{"name":"Bad zone","prompt":"Hi","crons":["0 8 * * *"],"timezone":"Nowhere"}]',
    ],
  ])("ignores an invalid %s with a warning", (key, value) => {
    const warn = vi.fn();
    const defaults = newUserDefaultsFromEnv({ [key]: value }, warn);
    expect(defaults).toEqual({
      timezone: undefined,
      locale: undefined,
      botInstructions: undefined,
      seedFirstBot: undefined,
    });
    expect(warn).toHaveBeenCalledWith(expect.stringContaining(`Ignoring ${key}:`));
  });

  it("never stops the server over a bad default", () => {
    const env = loadEnv({
      DATABASE_URL: "postgres://rakazo:rakazo@127.0.0.1:5433/rakazo",
      NODE_ENV: "test",
      RAKAZO_DEFAULT_TIMEZONE: "Mars/Olympus",
      RAKAZO_DEFAULT_ROUTINES: "[",
      RAKAZO_DEFAULT_LOCALE: "fr",
    });
    expect(env.newUserDefaults).toMatchObject({ timezone: undefined, locale: "fr" });
    expect(env.newUserDefaults.seedFirstBot).toBeUndefined();
  });
});

describe("bots.create with new user defaults", () => {
  const defaults = newUserDefaultsFromEnv({
    RAKAZO_DEFAULT_BOT_INSTRUCTIONS: "Answer in short paragraphs.",
    RAKAZO_DEFAULT_ROUTINES: ROUTINES,
  });
  const created = new Error("stop after createBot");

  async function create(input: { instructions: string; spawnKey?: string }) {
    createBot.mockReset();
    createBot.mockRejectedValue(created);
    const client = createRouterClient(
      createRouter({
        prisma: {},
        env: { agentRuntime: "scripted", newUserDefaults: defaults },
      } as unknown as RouterDeps),
      { context: { actor } },
    );
    await expect(
      client.bots.create({ name: "Chief", description: "", notifyOnFinish: true, ...input }),
    ).rejects.toThrow();
    expect(createBot).toHaveBeenCalledOnce();
    const [, request] = createBot.mock.calls[0]!;
    return request as {
      instructions: string;
      seed?: (tx: Prisma.TransactionClient, botId: string) => Promise<void>;
    };
  }

  it("gives the onboarding bot the default instructions and paused routines", async () => {
    const request = await create({ instructions: "", spawnKey: FIRST_BOT_SPAWN_KEY });
    expect(request.instructions).toBe("Answer in short paragraphs.");
    const routineCreate = vi.fn(async () => ({}));
    await request.seed?.(
      { routine: { create: routineCreate } } as unknown as Prisma.TransactionClient,
      "bot-1",
    );
    expect(routineCreate).toHaveBeenCalledTimes(2);
    expect(routineCreate).toHaveBeenCalledWith({
      data: expect.objectContaining({ botId: "bot-1", userId: "user-1", active: false }),
    });
  });

  it("keeps instructions the onboarding bot was given", async () => {
    const request = await create({ instructions: "Be brief.", spawnKey: FIRST_BOT_SPAWN_KEY });
    expect(request.instructions).toBe("Be brief.");
  });

  it("leaves other bots alone", async () => {
    const request = await create({ instructions: "" });
    expect(request.instructions).toBe("");
    expect(request.seed).toBeUndefined();
  });
});
