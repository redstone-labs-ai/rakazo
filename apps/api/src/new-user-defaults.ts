import type { CreateRoutineInput as CreateRoutineInputType } from "@rakazo/contracts";
import { AppLocaleSchema, CreateRoutineInput, TimeZoneSchema } from "@rakazo/contracts";
import type { NewUserDefaults } from "@rakazo/db";
import { getLogger } from "@rakazo/logging";
import { assertRoutineSchedule, routineNextRunAt, routineRowData } from "./routines.js";

/** Stands in for the first bot's id while a default routine is checked at startup. */
const PENDING_BOT_ID = "first-bot";

type Parsed<T> = { success: true; data: T } | { success: false; error: string };
type SchemaResult<T> =
  | { success: true; data: T }
  | { success: false; error: { issues: Array<{ path: PropertyKey[]; message: string }> } };

/**
 * The starting setup a deployment gives new accounts, from RAKAZO_DEFAULT_* settings. A value
 * that does not parse is logged and left out, so a typo never stops the server.
 */
export function newUserDefaultsFromEnv(
  source: NodeJS.ProcessEnv,
  warn: (message: string) => void = (message) => getLogger().warn(message),
): NewUserDefaults {
  const timezone = parsed(
    "RAKAZO_DEFAULT_TIMEZONE",
    source.RAKAZO_DEFAULT_TIMEZONE,
    warn,
    (value) => checked(TimeZoneSchema.safeParse(value)),
  );
  const locale = parsed("RAKAZO_DEFAULT_LOCALE", source.RAKAZO_DEFAULT_LOCALE, warn, (value) =>
    checked(AppLocaleSchema.safeParse(value)),
  );
  const routines = parsed(
    "RAKAZO_DEFAULT_ROUTINES",
    source.RAKAZO_DEFAULT_ROUTINES,
    warn,
    (value) => defaultRoutines(value, timezone),
  );
  return {
    timezone,
    locale,
    botInstructions: source.RAKAZO_DEFAULT_BOT_INSTRUCTIONS?.trim() || undefined,
    seedFirstBot: routines?.length
      ? async (tx, actor, botId) => {
          for (const routine of routines) {
            const input = { ...routine, botId };
            await tx.routine.create({
              data: routineRowData(actor, input, routineNextRunAt(input)),
            });
          }
        }
      : undefined,
  };
}

/** A JSON array of `{ name, prompt, crons, timezone? }`. Each routine starts paused. */
function defaultRoutines(
  value: string,
  timezone: string | undefined,
): Parsed<CreateRoutineInputType[]> {
  let json: unknown;
  try {
    json = JSON.parse(value);
  } catch {
    return { success: false, error: "not valid JSON" };
  }
  if (!Array.isArray(json)) return { success: false, error: "expected a JSON array" };
  const routines: CreateRoutineInputType[] = [];
  for (const [index, entry] of json.entries()) {
    const fields: Record<string, unknown> = typeof entry === "object" && entry ? entry : {};
    const result = checked(
      CreateRoutineInput.safeParse({
        botId: PENDING_BOT_ID,
        name: fields.name,
        prompt: fields.prompt,
        crons: fields.crons,
        timezone: fields.timezone ?? timezone,
        active: false,
      }),
    );
    if (!result.success) return { success: false, error: `routine ${index}: ${result.error}` };
    // The same schedule checks a routine created in the app gets.
    try {
      assertRoutineSchedule(result.data);
      routineNextRunAt(result.data);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return { success: false, error: `routine ${index}: ${message}` };
    }
    routines.push(result.data);
  }
  return { success: true, data: routines };
}

function checked<T>(result: SchemaResult<T>): Parsed<T> {
  if (result.success) return result;
  return {
    success: false,
    error: result.error.issues
      .map((issue) => `${issue.path.map(String).join(".") || "value"}: ${issue.message}`)
      .join("; "),
  };
}

function parsed<T>(
  key: string,
  raw: string | undefined,
  warn: (message: string) => void,
  parse: (value: string) => Parsed<T>,
): T | undefined {
  const value = raw?.trim();
  if (!value) return undefined;
  const result = parse(value);
  if (result.success) return result.data;
  warn(`Ignoring ${key}: ${result.error}`);
  return undefined;
}
