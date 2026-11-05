import { describe, expect, it, vi } from "vitest";
import { hiddenPluginsFromEnv, isHiddenPlugin } from "./hidden-plugins.js";

describe("hidden marketplace plugins", () => {
  it("hides nothing when unset or empty", () => {
    expect(hiddenPluginsFromEnv(undefined)).toEqual([]);
    expect(hiddenPluginsFromEnv(" ")).toEqual([]);
    expect(isHiddenPlugin([], "https://github.com/acme/plugins#mail")).toBe(false);
    expect(isHiddenPlugin(undefined, "https://github.com/acme/plugins#mail")).toBe(false);
  });

  it("reads GitHub shorthand and full sources, separated by commas or whitespace", () => {
    expect(
      hiddenPluginsFromEnv(
        "Acme/Plugins#tools/mail,\nhttps://github.com/acme/plugins/tree/main/tools/chat  acme/plugins#tools/cal*",
      ),
    ).toEqual([
      "https://github.com/acme/plugins#tools/mail",
      "https://github.com/acme/plugins#tools/chat",
      "https://github.com/acme/plugins#tools/cal*",
    ]);
  });

  it("matches a source however it is pinned or written", () => {
    const hidden = hiddenPluginsFromEnv("acme/plugins#tools/mail acme/plugins#tools/cal*");
    for (const source of [
      "https://github.com/acme/plugins#tools/mail",
      `https://github.com/acme/plugins@${"a".repeat(40)}#tools/mail`,
      "https://github.com/acme/plugins/tree/main/tools/mail",
      "https://github.com/acme/plugins#tools/calendar",
    ]) {
      expect(isHiddenPlugin(hidden, source)).toBe(true);
    }
    for (const source of [
      "https://github.com/acme/plugins#tools/mailer",
      "https://github.com/acme/plugins#tools/chat",
      "https://github.com/other/plugins#tools/mail",
      "not a source",
    ]) {
      expect(isHiddenPlugin(hidden, source)).toBe(false);
    }
  });

  it("warns about and ignores entries that are not plugin sources", () => {
    const warn = vi.fn();
    expect(
      hiddenPluginsFromEnv("acme acme/*/mail acme/plugins#../mail acme/plugins#mail", warn),
    ).toEqual(["https://github.com/acme/plugins#mail"]);
    expect(warn).toHaveBeenCalledTimes(3);
    expect(warn.mock.calls[0]![0]).toContain("RAKAZO_HIDDEN_MARKETPLACE_PLUGINS");
  });
});
