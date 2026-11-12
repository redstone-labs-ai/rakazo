/**
 * Redstone fork only (see REDSTONE.md). UI strings the fork adds ship in English only: Redstone
 * deploys set RAKAZO_UI_LOCALES=en, so no other catalog is ever shown. List each fork-added
 * source message here (its English msgid) so the every-locale catalog checks skip it. Never
 * hand-write or machine-generate translations for these; `redstone-english-only.test.ts` keeps
 * their non-English entries empty.
 */
export const REDSTONE_ENGLISH_ONLY_MESSAGES: ReadonlySet<string> = new Set<string>([]);
