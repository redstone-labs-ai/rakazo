-- Additive: the previous release never reads these tables.
-- System-only provisioning state with no tenant column, like messaging_channels: no row policies.
CREATE TABLE "agent_teams_apps" (
    "id" TEXT NOT NULL,
    "botId" TEXT NOT NULL,
    "status" TEXT NOT NULL,
    "entraAppId" TEXT,
    "entraObjectId" TEXT,
    "secretRef" TEXT,
    "azureBotName" TEXT,
    "catalogAppId" TEXT,
    "manifestVersion" TEXT,
    "publishedAt" TIMESTAMP(3),
    "firstInstalledAt" TIMESTAMP(3),
    "lastError" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "agent_teams_apps_pkey" PRIMARY KEY ("id")
);

CREATE TABLE "agent_teams_chats" (
    "id" TEXT NOT NULL,
    "agentAppId" TEXT NOT NULL,
    "chatId" TEXT NOT NULL,
    "installId" TEXT,
    "installedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "removedAt" TIMESTAMP(3),

    CONSTRAINT "agent_teams_chats_pkey" PRIMARY KEY ("id")
);

CREATE UNIQUE INDEX "agent_teams_apps_botId_key" ON "agent_teams_apps"("botId");
CREATE UNIQUE INDEX "agent_teams_apps_entraAppId_key" ON "agent_teams_apps"("entraAppId");
CREATE UNIQUE INDEX "agent_teams_chats_agentAppId_chatId_key" ON "agent_teams_chats"("agentAppId", "chatId");

-- RESTRICT: removing a group chat or a bot never cascades into these rows.
ALTER TABLE "agent_teams_chats" ADD CONSTRAINT "agent_teams_chats_agentAppId_fkey" FOREIGN KEY ("agentAppId") REFERENCES "agent_teams_apps"("id") ON DELETE RESTRICT ON UPDATE CASCADE;
