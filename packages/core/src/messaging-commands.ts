/**
 * Owner text commands, parsed only in the owner's verified 1:1 conversation.
 * Everything that is not an exact command is a normal message to the bot.
 */
export type MessagingCommand = "approve" | "decline" | "leave" | "stop";

export function parseMessagingCommand(text: string): MessagingCommand | null {
  const normalized = text.trim().toUpperCase();
  if (normalized === "YES") return "approve";
  if (normalized === "NO") return "decline";
  if (normalized === "LEAVE") return "leave";
  if (/^STOP(?: NOW)?[.!?…]*$/.test(normalized.replace(/\s+/g, " "))) return "stop";
  return null;
}

/** "/use <name>" picks which of the owner's bots the conversation reaches; "/use" or "/bots" lists them. */
export type BotChoiceCommand = { kind: "list" } | { kind: "use"; name: string };

export function parseBotChoiceCommand(text: string): BotChoiceCommand | null {
  const match = /^\/(use|bots)(?:\s+([\s\S]*))?$/i.exec(text.trim());
  if (!match) return null;
  const name = (match[2] ?? "").replace(/^["“](.*)["”]$/s, "$1").trim();
  if (match[1]!.toLowerCase() === "bots") return name ? null : { kind: "list" };
  return name ? { kind: "use", name } : { kind: "list" };
}
