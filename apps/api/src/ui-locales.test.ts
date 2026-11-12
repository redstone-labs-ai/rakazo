import { describe, expect, it, vi } from "vitest";
import { uiLocalesFromEnv } from "./ui-locales.js";

describe("uiLocalesFromEnv", () => {
  it("offers every language when unset or empty", () => {
    expect(uiLocalesFromEnv(undefined)).toBeUndefined();
    expect(uiLocalesFromEnv(" ")).toBeUndefined();
  });

  it("keeps listed locales in order without duplicates", () => {
    expect(uiLocalesFromEnv("en")).toEqual(["en"]);
    expect(uiLocalesFromEnv("en, de en")).toEqual(["en", "de"]);
  });

  it("logs and drops unknown entries", () => {
    const warn = vi.fn();
    expect(uiLocalesFromEnv("en,xx", warn)).toEqual(["en"]);
    expect(warn).toHaveBeenCalledWith('RAKAZO_UI_LOCALES: ignoring unknown locale "xx"');
    expect(uiLocalesFromEnv("xx", warn)).toBeUndefined();
  });
});
