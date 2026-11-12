import { AppLocaleSchema } from "@rakazo/contracts";
import { matchUiLocale } from "@rakazo/core";

export const UI_LOCALES = AppLocaleSchema.options;

export type UiLocale = (typeof UI_LOCALES)[number];

export const UI_LOCALE_STORAGE_KEY = "rakazo.uiLocale";
/** Last languages the server offered, so a slow or failed capabilities fetch keeps the limit. */
export const ALLOWED_UI_LOCALES_STORAGE_KEY = "rakazo.allowedUiLocales";

let allowedUiLocales: readonly UiLocale[] = UI_LOCALES;

/** Languages this deployment offers (`uiLocales` from the server); every language by default. */
export function getAllowedUiLocales(): readonly UiLocale[] {
  return allowedUiLocales;
}

/**
 * Limit the languages the UI can activate. `undefined` or an empty list offers every
 * language. The list is remembered so the next load applies it before the server answers.
 */
export function setAllowedUiLocales(
  locales: readonly string[] | null | undefined,
  storage: Pick<Storage, "setItem" | "removeItem"> | null = typeof localStorage !== "undefined"
    ? localStorage
    : null,
): void {
  const known = (locales ?? []).filter((locale): locale is UiLocale =>
    UI_LOCALES.some((supported) => supported === locale),
  );
  allowedUiLocales = known.length ? known : UI_LOCALES;
  if (!storage) return;
  try {
    if (known.length) storage.setItem(ALLOWED_UI_LOCALES_STORAGE_KEY, known.join(","));
    else storage.removeItem(ALLOWED_UI_LOCALES_STORAGE_KEY);
  } catch {
    // Ignore quota / private-mode failures; the in-memory limit still applies.
  }
}

/** Apply the limit remembered from the last server answer. */
export function restoreAllowedUiLocales(
  storage: Pick<Storage, "getItem"> | null = typeof localStorage !== "undefined"
    ? localStorage
    : null,
): void {
  let stored: string | null = null;
  try {
    stored = storage?.getItem(ALLOWED_UI_LOCALES_STORAGE_KEY) ?? null;
  } catch {
    stored = null;
  }
  const known = (stored ?? "")
    .split(",")
    .filter((locale): locale is UiLocale => UI_LOCALES.some((supported) => supported === locale));
  allowedUiLocales = known.length ? known : UI_LOCALES;
}

/** The locale itself when the deployment offers it, else English, else the first offered. */
export function allowedUiLocale(locale: UiLocale): UiLocale {
  if (allowedUiLocales.includes(locale)) return locale;
  return allowedUiLocales.includes("en") ? "en" : (allowedUiLocales[0] ?? "en");
}

export const UI_LOCALE_LABELS: Record<UiLocale, string> = {
  en: "English",
  de: "Deutsch",
  ko: "한국어",
  tr: "Türkçe",
  hi: "हिन्दी",
  "pt-BR": "Português (Brasil)",
  "zh-CN": "简体中文",
  "zh-TW": "繁體中文",
  es: "Español",
  ru: "Русский",
  fr: "Français",
  he: "עברית",
  nl: "Nederlands",
  ja: "日本語",
  pl: "Polski",
};

/** Normalize BCP-47 tags (`de-DE`, `ko-KR`, `pt-BR`, `zh-CN`, `es-ES`) to a supported UI locale, else `en`. */
export function normalizeUiLocale(raw: string | null | undefined): UiLocale {
  return matchUiLocale(raw, UI_LOCALES);
}

function readStoredLocale(storage: Pick<Storage, "getItem"> | null | undefined): string | null {
  if (!storage) return null;
  try {
    return storage.getItem(UI_LOCALE_STORAGE_KEY);
  } catch {
    return null;
  }
}

function readEnvDefault(env: { VITE_DEFAULT_UI_LOCALE?: string } | undefined): string | null {
  const value = env?.VITE_DEFAULT_UI_LOCALE;
  return typeof value === "string" && value.trim() ? value.trim() : null;
}

function readNavigatorLanguage(
  nav: Pick<Navigator, "language" | "languages"> | null | undefined,
): string | null {
  if (!nav) return null;
  if (typeof nav.language === "string" && nav.language) return nav.language;
  const first = nav.languages?.[0];
  return typeof first === "string" && first ? first : null;
}

export type ResolveUiLocaleOptions = {
  stored?: string | null;
  envDefault?: string | null;
  navigatorLanguage?: string | null;
  storage?: Pick<Storage, "getItem"> | null;
  env?: { VITE_DEFAULT_UI_LOCALE?: string };
  navigator?: Pick<Navigator, "language" | "languages"> | null;
};

/**
 * Order: saved choice (`localStorage`) → `VITE_DEFAULT_UI_LOCALE` →
 * `navigator.language` → English, limited to the languages the deployment offers.
 */
export function resolveUiLocale(options: ResolveUiLocaleOptions = {}): UiLocale {
  return allowedUiLocale(preferredUiLocale(options));
}

function preferredUiLocale(options: ResolveUiLocaleOptions): UiLocale {
  const stored =
    options.stored !== undefined
      ? options.stored
      : readStoredLocale(
          options.storage ?? (typeof localStorage !== "undefined" ? localStorage : null),
        );
  if (stored) return normalizeUiLocale(stored);

  const envDefault =
    options.envDefault !== undefined
      ? options.envDefault
      : readEnvDefault(
          options.env ??
            (typeof import.meta !== "undefined"
              ? {
                  VITE_DEFAULT_UI_LOCALE: (import.meta as ImportMeta).env?.VITE_DEFAULT_UI_LOCALE,
                }
              : undefined),
        );
  if (envDefault) return normalizeUiLocale(envDefault);

  const navigatorLanguage =
    options.navigatorLanguage !== undefined
      ? options.navigatorLanguage
      : readNavigatorLanguage(
          options.navigator ?? (typeof navigator !== "undefined" ? navigator : null),
        );
  if (navigatorLanguage) return normalizeUiLocale(navigatorLanguage);

  return "en";
}

export function persistUiLocale(
  locale: UiLocale,
  storage: Pick<Storage, "setItem"> | null = typeof localStorage !== "undefined"
    ? localStorage
    : null,
): void {
  if (!storage) return;
  try {
    storage.setItem(UI_LOCALE_STORAGE_KEY, locale);
  } catch {
    // Ignore quota / private-mode failures; in-memory locale still applies.
  }
}
