import type { AgentTeamsAppRecord, AgentTeamsAppStatus } from "./store.js";

/** Teams takes a few minutes to show a newly published org app; assume visible after this. */
export const AGENT_TEAMS_APP_VISIBLE_AFTER_MS = 5 * 60_000;

/** UI steps. `in_chats` means visible and installed in at least one chat. */
export type AgentTeamsVisibilityStep =
  | "creating"
  | "registered"
  | "published"
  | "visible"
  | "in_chats"
  | "removing"
  | "removed"
  | "failed";

export interface AgentTeamsVisibility {
  step: AgentTeamsVisibilityStep;
  label: string;
  /** True when `visible` is inferred from time since publish, not from a real chat install. */
  estimated: boolean;
  /** Chats the agent's app is installed in now. */
  chatCount: number;
}

const LABELS: Record<AgentTeamsVisibilityStep, string> = {
  creating: "Creating Teams identity",
  registered: "Registered with Microsoft",
  published: "Published to your org's Teams apps",
  visible: "Available in Teams",
  removing: "Removing from Teams",
  removed: "Removed from Teams",
  failed: "Teams setup failed",
  in_chats: "In Teams chats",
};

function isStatus(value: string): value is AgentTeamsAppStatus {
  return value in LABELS && value !== "in_chats";
}

/**
 * Where an agent's own Teams app stands, for the UI. Pure. Happy path:
 * creating → registered → published → visible → in_chats.
 */
export function visibilityStatus(
  app: Pick<AgentTeamsAppRecord, "status" | "publishedAt" | "firstInstalledAt">,
  now: Date,
  chatCount = 0,
): AgentTeamsVisibility {
  const status: AgentTeamsAppStatus = isStatus(app.status) ? app.status : "failed";
  const result = (step: AgentTeamsVisibilityStep, estimated: boolean): AgentTeamsVisibility => ({
    step,
    label: LABELS[step],
    estimated,
    chatCount,
  });
  if (status === "removing" || status === "removed" || status === "failed") {
    return result(status, false);
  }
  if (app.firstInstalledAt) return result(chatCount > 0 ? "in_chats" : "visible", false);
  if (
    (status === "published" || status === "visible") &&
    app.publishedAt &&
    now.getTime() - app.publishedAt.getTime() >= AGENT_TEAMS_APP_VISIBLE_AFTER_MS
  ) {
    return result("visible", true);
  }
  return result(status === "visible" ? "published" : status, false);
}
