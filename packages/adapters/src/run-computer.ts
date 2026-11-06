import type { ComputerRef } from "@rakazo/adapter-kit";

const startFailures = new WeakSet<object>();

/**
 * Whether the error is a run computer's failed start, as the run's setup catches it after the
 * turn rethrows it. Its message is the computer provider's, not the model's.
 */
export function isComputerStartFailure(error: unknown): boolean {
  return typeof error === "object" && error !== null && startFailures.has(error);
}

/**
 * A run's computer, started the first time a tool needs it.
 *
 * Remote computers bill while they run, so a run that never touches its computer must not
 * create or resume one. A failed start is kept: every later use rethrows it as a setup failure.
 */
export function lazyRunComputer(start: () => Promise<ComputerRef>, onFailure?: () => void) {
  let started: Promise<ComputerRef> | undefined;
  let ready: ComputerRef | undefined;
  let failed = false;
  let failure: unknown;
  return {
    get(): Promise<ComputerRef> {
      started ??= start().then(
        (computer) => {
          ready = computer;
          return computer;
        },
        (error: unknown) => {
          failed = true;
          failure = error;
          if (typeof error === "object" && error !== null) startFailures.add(error);
          onFailure?.();
          throw error;
        },
      );
      return started;
    },
    /** The computer if a tool already started it. */
    get ready(): ComputerRef | undefined {
      return ready;
    },
    throwIfFailed(): void {
      if (failed) throw failure;
    },
  };
}
