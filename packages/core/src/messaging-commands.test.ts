import { describe, expect, it } from "vitest";
import { parseBotChoiceCommand, parseMessagingCommand } from "./messaging-commands.js";

describe("parseBotChoiceCommand", () => {
  it("reads /use with a name, ignoring case, spacing and surrounding quotes", () => {
    expect(parseBotChoiceCommand("/use Research")).toEqual({ kind: "use", name: "Research" });
    expect(parseBotChoiceCommand("  /USE   Night Owl  ")).toEqual({
      kind: "use",
      name: "Night Owl",
    });
    expect(parseBotChoiceCommand('/use "Night Owl"')).toEqual({ kind: "use", name: "Night Owl" });
  });

  it("lists on a bare /use or /bots", () => {
    for (const text of ["/use", " /use  ", "/bots", "/BOTS"])
      expect(parseBotChoiceCommand(text)).toEqual({ kind: "list" });
  });

  it("treats everything else as a normal message", () => {
    for (const text of ["use Research", "/user x", "/bots please", "/usex", "please /use x", ""])
      expect(parseBotChoiceCommand(text)).toBeNull();
  });
});

describe("parseMessagingCommand", () => {
  it("parses the minimal command set case-insensitively", () => {
    expect(parseMessagingCommand("YES")).toBe("approve");
    expect(parseMessagingCommand("yes")).toBe("approve");
    expect(parseMessagingCommand(" yes ")).toBe("approve");
    expect(parseMessagingCommand("NO")).toBe("decline");
    expect(parseMessagingCommand("no")).toBe("decline");
    expect(parseMessagingCommand("LEAVE")).toBe("leave");
    expect(parseMessagingCommand("leave")).toBe("leave");
  });

  it("reads a bare stop or stop now, ignoring case, spacing and closing punctuation", () => {
    for (const text of ["stop", "Stop now.", "STOP!", "  stop   now ", "stop…"])
      expect(parseMessagingCommand(text)).toBe("stop");
    for (const text of ["stop it", "stopper", "don't stop", "/stop", "stop, now"])
      expect(parseMessagingCommand(text)).toBeNull();
  });

  it("treats everything else as a normal message", () => {
    expect(parseMessagingCommand("yes please")).toBeNull();
    expect(parseMessagingCommand("YES!")).toBeNull();
    expect(parseMessagingCommand("")).toBeNull();
    expect(parseMessagingCommand("hello")).toBeNull();
    expect(parseMessagingCommand("leave the group")).toBeNull();
  });
});
