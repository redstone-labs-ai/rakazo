import { describe, expect, it } from "vitest";
import {
  AUTOMATED_MESSAGE_LABEL,
  automated,
  PARTIAL_REPLY_NOTE,
  relabelToolImageMessages,
  TOOL_IMAGES_TEXT,
  withProductLabel,
} from "./automated-messages.js";
import { CUT_REPLY_CONTINUATION_PROMPT } from "./pi-runtime.js";
import { buildApprovalContinuation } from "./run-effects.js";
import { userTurnInstructions } from "./run-instructions.js";
import { formatFinishedShellCommand } from "./shell-command-stream.js";

const effects = [{ kind: "destination.write", request: { sequence: 1 } }];

function instructions(productName?: string): string {
  return userTurnInstructions({
    botInstructions: "Chief",
    groupContext: undefined,
    messagingContext: undefined,
    redactedMemoryContext: undefined,
    redactedScratchpadContext: undefined,
    hasHistoricalContext: false,
    computerInstruction: "Computer.",
    pageBrowserAllowed: false,
    workspaceInstruction: "Workspace.",
    agentEnvironmentInstruction: undefined,
    botDirectory: undefined,
    pluginLine: undefined,
    agentSkillsLine: "Skills: Rakazo.",
    taughtSkillsLine: undefined,
    replyGuidance: "Reply.",
    ...(productName ? { productName } : {}),
  })
    .filter(Boolean)
    .join("\n\n");
}

describe("product name in model-facing text", () => {
  it("keeps the Rakazo text by default", () => {
    expect(AUTOMATED_MESSAGE_LABEL).toBe("[Automated message from Rakazo, not from the user]");
    expect(withProductLabel(PARTIAL_REPLY_NOTE, undefined)).toBe(PARTIAL_REPLY_NOTE);
    expect(formatFinishedShellCommand({ code: 0, stdout: "", stderr: "" })).toBe(
      "[Automated message from Rakazo, not from the user] Shell command finished.\nexit code: 0",
    );
    expect(buildApprovalContinuation(effects, JSON.stringify)).toMatch(
      /^Rakazo is resuming after the user approved the exact tool request\(s\) below\./,
    );
    expect(instructions()).toContain(
      "When asked how Rakazo or you work, or how to set something up (apps, the user's own computer, billing), read the Rakazo skill before answering or acting.",
    );
  });

  it("uses a configured product name", () => {
    expect(automated("Hi.", "Bot")).toBe("[Automated message from Bot, not from the user] Hi.");
    expect(withProductLabel(PARTIAL_REPLY_NOTE, "Bot")).toBe(
      "[Automated message from Bot, not from the user] This reply was cut off by an error before it was finished.",
    );
    expect(withProductLabel(CUT_REPLY_CONTINUATION_PROMPT, "Bot")).toMatch(
      /^\[Automated message from Bot, not from the user\] Your previous reply was cut off/,
    );
    expect(withProductLabel("Plain text.", "Bot")).toBe("Plain text.");
    expect(formatFinishedShellCommand({ code: 0, stdout: "", stderr: "" }, "Bot")).toBe(
      "[Automated message from Bot, not from the user] Shell command finished.\nexit code: 0",
    );
    const relabeled = relabelToolImageMessages(
      { messages: [{ role: "user", content: [{ type: "text", text: TOOL_IMAGES_TEXT }] }] },
      "Bot",
    ) as { messages: Array<{ content: Array<{ text: string }> }> };
    expect(relabeled.messages[0]?.content[0]?.text).toBe(
      "[Automated message from Bot, not from the user] Screenshot(s) returned by your last tool call. The user has not replied.",
    );
    expect(buildApprovalContinuation(effects, JSON.stringify, { productName: "Bot" })).toMatch(
      /^Bot is resuming after the user approved the exact tool request\(s\) below\./,
    );
    // The help skill keeps its name: it indexes the upstream Rakazo docs.
    expect(instructions("Bot")).toContain(
      "When asked how Bot or you work, or how to set something up (apps, the user's own computer, billing), read the Rakazo skill before answering or acting.",
    );
  });
});
