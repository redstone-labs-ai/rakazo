import { AppLocaleSchema } from "@rakazo/contracts";
import { getLogger } from "@rakazo/logging";

export type UiLocale = (typeof AppLocaleSchema.options)[number];

/**
 * Web UI languages a deployment offers (RAKAZO_UI_LOCALES), for example `en` or `en,de`.
 * Unset or empty offers every language. Unknown entries are logged and left out, so a typo
 * never stops the server; a list with no valid entry offers every language.
 */
export function uiLocalesFromEnv(
  value: string | undefined,
  warn: (message: string) => void = (message) => getLogger().warn(message),
): UiLocale[] | undefined {
  const locales: UiLocale[] = [];
  for (const item of (value ?? "").split(/[\s,]+/).filter(Boolean)) {
    const parsed = AppLocaleSchema.safeParse(item);
    if (!parsed.success) {
      warn(`RAKAZO_UI_LOCALES: ignoring unknown locale "${item}"`);
      continue;
    }
    if (!locales.includes(parsed.data)) locales.push(parsed.data);
  }
  return locales.length ? locales : undefined;
}
