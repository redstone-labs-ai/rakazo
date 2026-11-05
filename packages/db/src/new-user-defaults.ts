import type { Actor } from "@rakazo/contracts";
import type { Prisma } from "./client.js";

/** A deployment's starting setup for new accounts. Each default fills only what is still empty. */
export interface NewUserDefaults {
  /** IANA timezone for an account that has none yet. */
  timezone?: string;
  /** App language for an account that has none yet. */
  locale?: string;
  /** Instructions for a user's first bot when it is created without any. */
  botInstructions?: string;
  /** Adds starting content, such as paused routines, to a user's first bot as it is created. */
  seedFirstBot?: (tx: Prisma.TransactionClient, actor: Actor, botId: string) => Promise<void>;
}

/** A user's first bot as the defaults complete it. Instructions it was given are kept. */
export function firstBotWithDefaults<T extends { instructions: string }>(
  actor: Actor,
  input: T,
  defaults: NewUserDefaults | undefined,
): T & { seed?: (tx: Prisma.TransactionClient, botId: string) => Promise<void> } {
  const seedFirstBot = defaults?.seedFirstBot;
  return {
    ...input,
    instructions: input.instructions.trim()
      ? input.instructions
      : (defaults?.botInstructions ?? input.instructions),
    ...(seedFirstBot
      ? {
          seed: (tx: Prisma.TransactionClient, botId: string) => seedFirstBot(tx, actor, botId),
        }
      : {}),
  };
}
