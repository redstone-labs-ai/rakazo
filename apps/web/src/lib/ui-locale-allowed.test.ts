import { afterEach, describe, expect, it, vi } from "vitest";
import { activateUiLocale, i18n, setCatalogLoadersForTests } from "./i18n";
import type { UiLocale } from "./ui-locale";
import {
  ALLOWED_UI_LOCALES_STORAGE_KEY,
  getAllowedUiLocales,
  resolveUiLocale,
  restoreAllowedUiLocales,
  setAllowedUiLocales,
  UI_LOCALES,
} from "./ui-locale";

vi.mock("./apply-ui-direction", () => ({ applyUiDirection: vi.fn() }));

function memoryStorage() {
  const store = new Map<string, string>();
  return {
    store,
    getItem: (key: string) => store.get(key) ?? null,
    setItem: (key: string, value: string) => {
      store.set(key, value);
    },
    removeItem: (key: string) => {
      store.delete(key);
    },
  };
}

const browser = (language: string) => ({
  stored: null,
  envDefault: null,
  navigatorLanguage: language,
});

afterEach(() => {
  setAllowedUiLocales(undefined, null);
  setCatalogLoadersForTests(null);
});

describe("deployment UI languages (RAKAZO_UI_LOCALES)", () => {
  it("keeps every language and browser detection when unset", () => {
    setAllowedUiLocales(undefined, null);
    expect(getAllowedUiLocales()).toEqual(UI_LOCALES);
    expect(resolveUiLocale(browser("pl-PL"))).toBe("pl");
    expect(resolveUiLocale(browser("de-DE"))).toBe("de");
  });

  it.each(["pl", "pl-PL", "de", "de-DE"])(
    "gives a %s browser English when only en is offered",
    (tag) => {
      setAllowedUiLocales(["en"], null);
      expect(resolveUiLocale(browser(tag))).toBe("en");
    },
  );

  it("ignores a saved choice the deployment no longer offers", () => {
    setAllowedUiLocales(["en"], null);
    expect(resolveUiLocale({ stored: "pl", envDefault: null, navigatorLanguage: null })).toBe("en");
    expect(resolveUiLocale({ stored: null, envDefault: "de", navigatorLanguage: null })).toBe("en");
  });

  it("keeps an offered language", () => {
    setAllowedUiLocales(["en", "de"], null);
    expect(resolveUiLocale(browser("de-DE"))).toBe("de");
    expect(resolveUiLocale(browser("pl-PL"))).toBe("en");
  });

  it("never loads the catalog of a language that is not offered", async () => {
    const loaded: string[] = [];
    const loaders = Object.fromEntries(
      UI_LOCALES.map((locale) => [
        locale,
        async () => {
          loaded.push(locale);
          return { messages: {} };
        },
      ]),
    ) as Record<UiLocale, () => Promise<{ messages: Record<string, string> }>>;
    setCatalogLoadersForTests(loaders);
    setAllowedUiLocales(["en"], null);
    await expect(activateUiLocale("pl")).resolves.toBe("en");
    expect(i18n.locale).toBe("en");
    expect(loaded).toEqual(["en"]);
  });

  it("remembers the limit for the next load and clears it when unset", () => {
    const storage = memoryStorage();
    setAllowedUiLocales(["en", "xx"], storage);
    expect(storage.store.get(ALLOWED_UI_LOCALES_STORAGE_KEY)).toBe("en");
    setAllowedUiLocales(undefined, null);
    restoreAllowedUiLocales(storage);
    expect(getAllowedUiLocales()).toEqual(["en"]);
    setAllowedUiLocales(undefined, storage);
    expect(storage.store.has(ALLOWED_UI_LOCALES_STORAGE_KEY)).toBe(false);
    restoreAllowedUiLocales(storage);
    expect(getAllowedUiLocales()).toEqual(UI_LOCALES);
  });
});
