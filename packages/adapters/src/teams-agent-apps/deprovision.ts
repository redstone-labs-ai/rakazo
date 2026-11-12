import type { AgentTeamsDeprovisionDeps } from "./deps.js";
import type { TokenProvider } from "./http.js";
import {
  ARM_SCOPE,
  asRecord,
  BOT_CONNECTOR_SCOPE,
  clientCredentialTokens,
  GRAPH_BASE,
  GRAPH_SCOPE,
  msCall,
  stringField,
} from "./http.js";
import { AGENT_RSC_PERMISSIONS } from "./manifest.js";
import { AZURE_BOT_API_VERSION, azureBotUrl, safeErrorMessage } from "./provision.js";
import type { AgentTeamsAppRecord, AgentTeamsChatRecord } from "./store.js";

/** The message the shared bot posts in each chat the agent leaves. */
export function agentHandoffMessage(agentName: string): string {
  return `${agentName} was removed. @Bot is here instead.`;
}

/**
 * Takes an agent's own Teams identity away. Idempotent and retryable: each step clears the id it
 * consumed, so a retry skips finished work. For each chat the agent is in, the shared Bot app is
 * installed first, posts a handoff, and only then is the agent app uninstalled. Then the catalog
 * app, Azure Bot, Entra app and Key Vault secret are deleted, and the row ends `removed`.
 * On failure the row stays `removing` with `lastError`, ready for the job queue to retry.
 */
export async function deprovisionAgentTeamsApp(
  botId: string,
  deps: AgentTeamsDeprovisionDeps,
  /** The agent's name for the handoff message; the bot row may already be gone. */
  options: { agentName?: string } = {},
): Promise<AgentTeamsAppRecord | null> {
  const now = deps.now ?? (() => new Date());
  const agentName = options.agentName?.trim() || "This agent";
  const found = await deps.store.findByBotId(botId);
  if (!found || found.status === "removed") return found;
  let app = await deps.store.update(found.id, { status: "removing" });
  try {
    for (const chat of await deps.store.activeChats(app.id)) {
      await handOffChat(app, chat, agentName, deps);
      await deps.store.markChatRemoved(chat.id, now());
    }
    if (app.catalogAppId) {
      await msCall(deps.fetch, {
        method: "DELETE",
        url: `${GRAPH_BASE}/appCatalogs/teamsApps/${encodeURIComponent(app.catalogAppId)}`,
        token: await deps.catalogToken(),
        allow: [404],
        label: "delete Teams catalog app",
      });
      app = await deps.store.update(app.id, { catalogAppId: null });
    }
    if (app.azureBotName) {
      await msCall(deps.fetch, {
        method: "DELETE",
        url: `${azureBotUrl(deps, app.azureBotName)}?api-version=${AZURE_BOT_API_VERSION}`,
        token: await deps.provisionerToken(ARM_SCOPE),
        allow: [404],
        label: "delete Azure Bot",
      });
      app = await deps.store.update(app.id, { azureBotName: null });
    }
    if (app.entraObjectId) {
      // Deleting the application also deletes its service principal and role grants.
      await msCall(deps.fetch, {
        method: "DELETE",
        url: `${GRAPH_BASE}/applications/${encodeURIComponent(app.entraObjectId)}`,
        token: await deps.provisionerToken(GRAPH_SCOPE),
        allow: [404],
        label: "delete Entra app",
      });
      // entraAppId stays as the record of which identity this was.
      app = await deps.store.update(app.id, { entraObjectId: null });
    }
    if (app.secretRef) {
      await deps.secrets.delete(app.secretRef);
      app = await deps.store.update(app.id, { secretRef: null });
    }
    return await deps.store.update(app.id, { status: "removed", lastError: null });
  } catch (error) {
    await deps.store.update(app.id, { status: "removing", lastError: safeErrorMessage(error) });
    throw error;
  }
}

async function handOffChat(
  app: AgentTeamsAppRecord,
  chat: AgentTeamsChatRecord,
  agentName: string,
  deps: AgentTeamsDeprovisionDeps,
): Promise<void> {
  const shared = deps.sharedBot;
  const sharedGraph = await shared.token(GRAPH_SCOPE);
  const chatUrl = `${GRAPH_BASE}/chats/${encodeURIComponent(chat.chatId)}`;

  // 1. The shared Bot app must be in the chat before the agent leaves it.
  const sharedInstall = await findInstall(deps, sharedGraph, chat.chatId, shared.catalogAppId);
  if (!sharedInstall) {
    await msCall(deps.fetch, {
      method: "POST",
      url: `${chatUrl}/installedApps`,
      token: sharedGraph,
      body: {
        "teamsApp@odata.bind": `${GRAPH_BASE}/appCatalogs/teamsApps/${encodeURIComponent(shared.catalogAppId)}`,
        consentedPermissionSet: {
          resourceSpecificPermissions: AGENT_RSC_PERMISSIONS.map((permissionValue) => ({
            permissionValue,
            permissionType: "Application",
          })),
        },
      },
      label: "install shared Bot app",
    });
  }

  // 2. Handoff message, from the shared bot.
  const base = shared.serviceUrl.endsWith("/") ? shared.serviceUrl : `${shared.serviceUrl}/`;
  await msCall(deps.fetch, {
    method: "POST",
    url: `${base}v3/conversations/${encodeURIComponent(chat.chatId)}/activities`,
    token: await shared.token(BOT_CONNECTOR_SCOPE),
    body: { type: "message", text: agentHandoffMessage(agentName) },
    label: "post handoff message",
  });

  // 3. Uninstall the agent app, as itself (its role covers only its own installs).
  if (!app.entraAppId || !app.secretRef) return;
  const agentToken = (deps.agentToken ?? defaultAgentToken(deps))(app.entraAppId, app.secretRef);
  const agentGraph = await agentToken(GRAPH_SCOPE);
  const installId =
    chat.installId ??
    (app.catalogAppId
      ? await findInstall(deps, agentGraph, chat.chatId, app.catalogAppId)
      : undefined);
  if (!installId) return;
  await msCall(deps.fetch, {
    method: "DELETE",
    url: `${chatUrl}/installedApps/${encodeURIComponent(installId)}`,
    token: agentGraph,
    allow: [404],
    label: "uninstall agent app",
  });
}

function defaultAgentToken(
  deps: AgentTeamsDeprovisionDeps,
): (entraAppId: string, secretRef: string) => TokenProvider {
  return (entraAppId, secretRef) =>
    clientCredentialTokens({
      tenantId: deps.config.tenantId,
      clientId: entraAppId,
      clientSecret: () => deps.secrets.get(secretRef),
      fetch: deps.fetch,
    });
}

/** The chat's install id of a catalog app, if that app is installed in the chat. */
async function findInstall(
  deps: Pick<AgentTeamsDeprovisionDeps, "fetch">,
  token: string,
  chatId: string,
  catalogAppId: string,
): Promise<string | undefined> {
  const filter = encodeURIComponent(`teamsApp/id eq '${catalogAppId}'`);
  const found = await msCall(deps.fetch, {
    method: "GET",
    url: `${GRAPH_BASE}/chats/${encodeURIComponent(chatId)}/installedApps?$expand=teamsApp&$filter=${filter}`,
    token,
    label: "read chat apps",
  });
  const first = asRecord(Array.isArray(found.json?.value) ? found.json.value[0] : undefined);
  return stringField(first, "id") || undefined;
}
