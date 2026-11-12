import { teamsAgentAppDeprovisionJob } from "@rakazo/adapter-kit";
import { destroyBot } from "@rakazo/adapters";
import type { Actor } from "@rakazo/contracts";
import type { createRepos } from "@rakazo/db";
import { getLogger } from "@rakazo/logging";
import type { RouterDeps } from "./router.js";

export function createBotRemoval(deps: RouterDeps, repos: ReturnType<typeof createRepos>) {
  const removeBot = async (
    actor: Actor,
    botId: string,
    options: Parameters<typeof destroyBot>[3],
  ) => {
    const bot = await repos.getBot(actor, botId, { includeArchived: true });
    await destroyBot(
      {
        prisma: deps.prisma,
        sandbox: deps.sandbox,
        home: deps.home,
        jobs: deps.jobs,
        artifacts: deps.artifacts,
        secretStore: deps.secrets,
        dataDir: deps.dataDir,
        messagingInstallations: deps.messagingInstallations,
      },
      bot,
      {
        operationId: "destroy",
        traceId: "destroy",
        spaceId: actor.spaceId,
        userId: actor.userId,
        botId: bot.id,
        signal: new AbortController().signal,
      },
      options,
    );
    // The agent's own Teams app outlives the bot row; a retried job removes it.
    // The same TEAMS_AGENT_APPS_ENABLED accessor the agent Teams app routes read.
    if (deps.env.teamsAgentAppsEnabled) {
      await deps.jobs
        .enqueue(teamsAgentAppDeprovisionJob({ botId: bot.id, agentName: bot.name }))
        .catch((error: unknown) => {
          getLogger().error("teams.agent-app-deprovision enqueue error", error);
        });
    }
  };
  return removeBot;
}
