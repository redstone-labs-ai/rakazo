import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { UI_LOCALES } from "../lib/ui-locale";

const read = (path: string) => readFileSync(fileURLToPath(new URL(path, import.meta.url)), "utf8");

/** Lingui's message id text for one macro: simple `${name}` → `{name}`, others → `{0}`, `{1}`…. */
function messageOf(raw: string, template: boolean): string {
  let index = 0;
  const placeholders = template ? /\$\{([^}]+)\}/g : /\{([^}]+)\}/g;
  return raw
    .replace(placeholders, (_match, expression: string) => {
      const name = expression.trim();
      return /^[A-Za-z_$][\w$]*$/.test(name) ? `{${name}}` : `{${index++}}`;
    })
    .replace(/\s+/g, " ")
    .trim();
}

/** Every `<Trans>…</Trans>` and t`…` message in the card's source. */
function cardMessages(): string[] {
  const source = read("./TeamsAppCard.tsx");
  const trans = [...source.matchAll(/<Trans>([\s\S]*?)<\/Trans>/g)].map((match) =>
    messageOf(match[1] ?? "", false),
  );
  const tagged = [...source.matchAll(/\bt`([^`]*)`/g)].map((match) =>
    messageOf(match[1] ?? "", true),
  );
  return [...new Set([...trans, ...tagged])];
}

function msgids(catalog: string): Set<string> {
  return new Set(
    [...catalog.matchAll(/^msgid (".*")$/gm)].map((match) => JSON.parse(match[1] ?? '""')),
  );
}

describe("TeamsAppCard catalog", () => {
  it("finds the card's strings", () => {
    expect(cardMessages()).toContain("Give {botName} its own Teams app");
    expect(cardMessages()).toContain("In {0} chats");
  });

  // A string missing from the catalogs renders as its hashed id (e.g. "E9bKSA") in the app.
  it("has every card string in every locale catalog (run `pnpm intl:extract`)", () => {
    const messages = cardMessages();
    for (const locale of UI_LOCALES) {
      const ids = msgids(read(`../locales/${locale}/messages.po`));
      const missing = messages.filter((message) => !ids.has(message));
      expect(missing, locale).toEqual([]);
    }
  });
});
