-- Additive: nullable columns the previous release never reads.
-- Per-bot operation lease, so provision and deprovision never run at once.
ALTER TABLE "agent_teams_apps" ADD COLUMN "lockedUntil" TIMESTAMP(3);
-- When the shared Bot's handoff was posted in a chat, so a retried removal does not repost it.
ALTER TABLE "agent_teams_chats" ADD COLUMN "handoffPostedAt" TIMESTAMP(3);
