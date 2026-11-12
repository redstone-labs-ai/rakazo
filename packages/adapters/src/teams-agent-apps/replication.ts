import { MicrosoftApiError } from "./http.js";

/**
 * Entra directory replication lag: an application or service principal created a moment ago may
 * not be visible to the next Graph request yet. Graph then answers 404 `Request_ResourceNotFound`
 * (or 400 carrying that code, on app role grant writes for just-created principals).
 *
 * `withReplicationRetry` retries ONLY that error, and only when the caller says the target was
 * just created; every other error, and every guard check, is untouched.
 */

/** Backoff between attempts: ~60s in all. */
export const REPLICATION_RETRY_DELAYS_MS = [2_000, 4_000, 8_000, 16_000, 30_000] as const;
/** An object created within this window may still be replicating. */
export const REPLICATION_WINDOW_MS = 10 * 60_000;
const NOT_FOUND_CODE = "Request_ResourceNotFound";

export type Sleep = (ms: number) => Promise<void>;
export const defaultSleep: Sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/** The error is Graph not seeing a just-created object yet. */
export function isReplicationLag(error: unknown): boolean {
  if (!(error instanceof MicrosoftApiError)) return false;
  if (error.status !== 404 && error.status !== 400) return false;
  return error.code === NOT_FOUND_CODE || error.innerCode === NOT_FOUND_CODE;
}

/** Created (per a stored timestamp) recently enough that replication lag is plausible. */
export function createdRecently(createdAt: Date | null | undefined, now: Date): boolean {
  if (!createdAt) return false;
  const age = now.getTime() - createdAt.getTime();
  return age >= 0 && age <= REPLICATION_WINDOW_MS;
}

export interface ReplicationRetryOptions {
  /** True only when the target was just created; otherwise `fn` runs once. */
  recent: boolean;
  sleep?: Sleep;
  delays?: readonly number[];
}

export async function withReplicationRetry<T>(
  fn: () => Promise<T>,
  options: ReplicationRetryOptions,
): Promise<T> {
  if (!options.recent) return fn();
  const delays = options.delays ?? REPLICATION_RETRY_DELAYS_MS;
  const sleep = options.sleep ?? defaultSleep;
  for (let attempt = 0; ; attempt++) {
    try {
      return await fn();
    } catch (error) {
      const delay = delays[attempt];
      if (delay === undefined || !isReplicationLag(error)) throw error;
      await sleep(delay);
    }
  }
}
