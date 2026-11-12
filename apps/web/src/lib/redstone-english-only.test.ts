import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { formatter } from "@lingui/format-po";
import { describe, expect, it } from "vitest";
import { REDSTONE_ENGLISH_ONLY_MESSAGES } from "./redstone-english-only";
import { UI_LOCALES } from "./ui-locale";

async function readCatalog(locale: string) {
  const filename = fileURLToPath(new URL(`../locales/${locale}/messages.po`, import.meta.url));
  return formatter().parse(readFileSync(filename, "utf8"), {
    locale,
    sourceLocale: "en",
    filename,
  });
}

describe("Redstone English-only UI strings", () => {
  it("lists only messages the English catalog has", async () => {
    const source = Object.values(await readCatalog("en")).map((entry) => entry.message);
    for (const message of REDSTONE_ENGLISH_ONLY_MESSAGES) {
      expect(source, message).toContain(message);
    }
  });

  it.each(UI_LOCALES.filter((locale) => locale !== "en"))(
    "ships no %s translation for them",
    async (locale) => {
      const catalog = await readCatalog(locale);
      for (const entry of Object.values(catalog)) {
        if (!entry.message || !REDSTONE_ENGLISH_ONLY_MESSAGES.has(entry.message)) continue;
        expect(entry.translation?.trim() ?? "", entry.message).toBe("");
      }
    },
  );
});
