import { getLogger } from "@rakazo/logging";
import type { AgentTeamsAppsDeps } from "./deps.js";
import {
  ARM_BASE,
  ARM_SCOPE,
  asRecord,
  GRAPH_BASE,
  GRAPH_SCOPE,
  MICROSOFT_GRAPH_APP_ID,
  MicrosoftApiError,
  msCall,
  stringField,
} from "./http.js";
import { agentIcons, buildAgentTeamsManifest, buildManifestZip } from "./manifest.js";
import type { AgentTeamsAppRecord, AgentTeamsAppStore, AgentTeamsChatRecord } from "./store.js";

export const AZURE_BOT_API_VERSION = "2022-09-15";
/** Graph application permissions every agent app gets, so it can create and join chats itself. */
export const AGENT_GRAPH_APP_ROLES = [
  "Chat.Create",
  "TeamsAppInstallation.ReadWriteAndConsentSelfForChat.All",
] as const;
const DEFAULT_MANIFEST_VERSION = "1.0.0";
const LAST_ERROR_MAX = 500;

export interface ProvisionAgentTeamsAppInput {
  botId: string;
  name: string;
  description: string;
  /** A square PNG avatar; without one the icon is the name's initial on a color. */
  avatar?: Uint8Array;
  /** Hex accent color, e.g. the bot's `color`. */
  color?: string;
}

/** The Entra app's unique name: the idempotency key for app creation. */
export function agentEntraUniqueName(botId: string): string {
  return `rakazo-agent-${botId}`;
}

/** The Azure Bot resource name: 2-64 chars of letters, digits, `-` and `_`. */
export function agentAzureBotName(botId: string): string {
  return `rakazo-agent-${botId.toLowerCase().replace(/[^a-z0-9_-]/g, "")}`.slice(0, 64);
}

/** The Key Vault secret name: letters, digits and `-` only. */
export function agentSecretName(botId: string): string {
  return `teams-agent-${botId.toLowerCase().replace(/[^a-z0-9-]/g, "")}`.slice(0, 127);
}

/** An error message safe to persist: no tokens, no secret values, bounded. */
export function safeErrorMessage(error: unknown): string {
  const message =
    error instanceof MicrosoftApiError || error instanceof Error ? error.message : String(error);
  return message
    .replace(/Bearer\s+[\w.~+/=-]+/gi, "Bearer [redacted]")
    .replace(/(secret|password|token)(["':=\s]+)[^\s"',}]+/gi, "$1$2[redacted]")
    .slice(0, LAST_ERROR_MAX);
}

/**
 * Gives an agent its own Teams identity. Idempotent and resumable: every step is skipped when
 * the row already holds its result, so a retry after a failure picks up where it stopped.
 * Steps: Entra app → service principal → password in Key Vault → Azure Bot + Teams channel →
 * Graph app roles → manifest → org catalog publish.
 */
export async function provisionAgentTeamsApp(
  input: ProvisionAgentTeamsAppInput,
  deps: AgentTeamsAppsDeps,
): Promise<AgentTeamsAppRecord> {
  const now = deps.now ?? (() => new Date());
  let app = await deps.store.ensure(input.botId);
  if (app.status === "removing") {
    throw new Error("This agent's Teams app is being removed; provision it again afterwards.");
  }
  if (app.status === "published" || app.status === "visible") return app;
  if (app.status === "removed") {
    app = await deps.store.update(app.id, {
      status: "creating",
      entraAppId: null,
      entraObjectId: null,
      secretRef: null,
      azureBotName: null,
      catalogAppId: null,
      manifestVersion: null,
      publishedAt: null,
      firstInstalledAt: null,
      lastError: null,
    });
  } else if (app.status === "failed") {
    app = await deps.store.update(app.id, { status: "creating", lastError: null });
  }
  let published: AgentTeamsAppRecord;
  try {
    published = await runProvision(app, input, deps, now);
  } catch (error) {
    await deps.store.update(app.id, { status: "failed", lastError: safeErrorMessage(error) });
    throw error;
  }
  if (deps.onPublished) {
    await deps.onPublished(published).catch((error: unknown) => {
      getLogger().warn("agent Teams app published but platform refresh failed", {
        botId: published.botId,
        error: safeErrorMessage(error),
      });
    });
  }
  return published;
}

async function runProvision(
  start: AgentTeamsAppRecord,
  input: ProvisionAgentTeamsAppInput,
  deps: AgentTeamsAppsDeps,
  now: () => Date,
): Promise<AgentTeamsAppRecord> {
  const { store, config } = deps;
  let app = start;
  const graph = () => deps.provisionerToken(GRAPH_SCOPE);

  // 1. Entra app, single-tenant. The unique name makes a repeated create return the same app.
  if (!app.entraAppId || !app.entraObjectId) {
    const created = await upsertEntraApp(deps, await graph(), input);
    app = await store.update(app.id, {
      entraAppId: created.appId,
      entraObjectId: created.objectId,
    });
  }
  const entraAppId = required(app.entraAppId, "entraAppId");
  const entraObjectId = required(app.entraObjectId, "entraObjectId");

  // 2. Service principal (needed for tokens and app role grants).
  const servicePrincipalId = await ensureServicePrincipal(deps, await graph(), entraAppId);

  // 3. Password credential, stored only in Key Vault.
  if (!app.secretRef) {
    const added = await msCall(deps.fetch, {
      method: "POST",
      url: `${GRAPH_BASE}/applications/${encodeURIComponent(entraObjectId)}/addPassword`,
      token: await graph(),
      body: { passwordCredential: { displayName: "rakazo-teams-bot" } },
      label: "add Entra app password",
    });
    const secretText = stringField(added.json, "secretText");
    if (!secretText) throw new Error("add Entra app password returned no secret");
    const secretRef = agentSecretName(input.botId);
    await deps.secrets.set(secretRef, secretText);
    app = await store.update(app.id, { secretRef });
  }

  // 4. Azure Bot with the Teams channel. PUT is idempotent, so both run before the name is saved.
  if (!app.azureBotName) {
    const azureBotName = agentAzureBotName(input.botId);
    const arm = await deps.provisionerToken(ARM_SCOPE);
    const botUrl = azureBotUrl(deps, azureBotName);
    await msCall(deps.fetch, {
      method: "PUT",
      url: `${botUrl}?api-version=${AZURE_BOT_API_VERSION}`,
      token: arm,
      body: {
        location: "global",
        kind: "azurebot",
        sku: { name: "F0" },
        properties: {
          displayName: input.name,
          description: input.description.slice(0, 512),
          endpoint: config.messagingEndpoint,
          msaAppId: entraAppId,
          msaAppType: "SingleTenant",
          msaAppTenantId: config.tenantId,
        },
      },
      label: "create Azure Bot",
    });
    await msCall(deps.fetch, {
      method: "PUT",
      url: `${botUrl}/channels/MsTeamsChannel?api-version=${AZURE_BOT_API_VERSION}`,
      token: arm,
      body: {
        location: "global",
        properties: { channelName: "MsTeamsChannel", properties: { isEnabled: true } },
      },
      label: "enable Teams channel",
    });
    app = await store.update(app.id, { azureBotName, status: "registered" });
  }

  // 5. Graph app roles for the agent's own service principal. Cheap and idempotent; always checked.
  await grantAgentAppRoles(deps, await graph(), servicePrincipalId);

  // 6–7. Manifest and org catalog publish, with the delegated admin token.
  if (!app.catalogAppId) {
    const catalogToken = await deps.catalogToken();
    const version = config.manifestVersion ?? DEFAULT_MANIFEST_VERSION;
    const existing = await findCatalogApp(deps, catalogToken, entraAppId);
    let catalogAppId = existing;
    if (!catalogAppId) {
      const manifest = buildAgentTeamsManifest({
        appId: entraAppId,
        name: input.name,
        description: input.description,
        version,
        developer: config.developer,
        ...(input.color ? { accentColor: input.color } : {}),
      });
      const icons = agentIcons(input.name, {
        ...(input.avatar ? { avatar: input.avatar } : {}),
        ...(input.color ? { color: input.color } : {}),
      });
      const published = await msCall(deps.fetch, {
        method: "POST",
        url: `${GRAPH_BASE}/appCatalogs/teamsApps`,
        token: catalogToken,
        body: buildManifestZip(manifest, icons),
        contentType: "application/zip",
        label: "publish Teams catalog app",
      });
      catalogAppId = stringField(published.json, "id");
      if (!catalogAppId) throw new Error("publish Teams catalog app returned no id");
    }
    app = await store.update(app.id, {
      catalogAppId,
      manifestVersion: version,
      publishedAt: now(),
      status: "published",
      lastError: null,
    });
  } else if (app.status !== "published") {
    app = await store.update(app.id, {
      status: "published",
      publishedAt: app.publishedAt ?? now(),
      lastError: null,
    });
  }
  return app;
}

function required(value: string | null, field: string): string {
  if (!value) throw new Error(`agent Teams app has no ${field}`);
  return value;
}

export function azureBotUrl(
  deps: Pick<AgentTeamsAppsDeps, "config">,
  azureBotName: string,
): string {
  const { subscriptionId, resourceGroup } = deps.config;
  return `${ARM_BASE}/subscriptions/${encodeURIComponent(subscriptionId)}/resourceGroups/${encodeURIComponent(resourceGroup)}/providers/Microsoft.BotService/botServices/${encodeURIComponent(azureBotName)}`;
}

async function upsertEntraApp(
  deps: AgentTeamsAppsDeps,
  token: string,
  input: ProvisionAgentTeamsAppInput,
): Promise<{ appId: string; objectId: string }> {
  const url = `${GRAPH_BASE}/applications(uniqueName='${encodeURIComponent(agentEntraUniqueName(input.botId))}')`;
  const upserted = await msCall(deps.fetch, {
    method: "PATCH",
    url,
    token,
    headers: { prefer: "create-if-missing" },
    body: { displayName: input.name, signInAudience: "AzureADMyOrg" },
    label: "create Entra app",
  });
  let record = upserted.json;
  if (!stringField(record, "appId")) {
    record = (
      await msCall(deps.fetch, {
        method: "GET",
        url: `${url}?$select=id,appId`,
        token,
        label: "read Entra app",
      })
    ).json;
  }
  const appId = stringField(record, "appId");
  const objectId = stringField(record, "id");
  if (!appId || !objectId) throw new Error("create Entra app returned no app id");
  return { appId, objectId };
}

async function ensureServicePrincipal(
  deps: AgentTeamsAppsDeps,
  token: string,
  appId: string,
): Promise<string> {
  const found = await msCall(deps.fetch, {
    method: "GET",
    url: `${GRAPH_BASE}/servicePrincipals(appId='${encodeURIComponent(appId)}')?$select=id`,
    token,
    allow: [404],
    label: "read service principal",
  });
  const existing = stringField(found.json, "id");
  if (existing) return existing;
  const created = await msCall(deps.fetch, {
    method: "POST",
    url: `${GRAPH_BASE}/servicePrincipals`,
    token,
    body: { appId },
    label: "create service principal",
  });
  const id = stringField(created.json, "id");
  if (!id) throw new Error("create service principal returned no id");
  return id;
}

async function grantAgentAppRoles(
  deps: AgentTeamsAppsDeps,
  token: string,
  servicePrincipalId: string,
): Promise<void> {
  const graphSp = await msCall(deps.fetch, {
    method: "GET",
    url: `${GRAPH_BASE}/servicePrincipals(appId='${MICROSOFT_GRAPH_APP_ID}')?$select=id,appRoles`,
    token,
    label: "read Microsoft Graph service principal",
  });
  const resourceId = stringField(graphSp.json, "id");
  const roles = Array.isArray(graphSp.json?.appRoles) ? graphSp.json.appRoles : [];
  const roleIds = new Map<string, string>();
  for (const role of roles) {
    const record = asRecord(role);
    const value = stringField(record, "value");
    const id = stringField(record, "id");
    if (value && id) roleIds.set(value, id);
  }
  const assigned = await msCall(deps.fetch, {
    method: "GET",
    url: `${GRAPH_BASE}/servicePrincipals/${encodeURIComponent(servicePrincipalId)}/appRoleAssignments`,
    token,
    label: "read app role assignments",
  });
  const current = new Set(
    (Array.isArray(assigned.json?.value) ? assigned.json.value : [])
      .map((entry) => asRecord(entry))
      .filter((entry) => stringField(entry, "resourceId") === resourceId)
      .map((entry) => stringField(entry, "appRoleId")),
  );
  for (const value of AGENT_GRAPH_APP_ROLES) {
    const appRoleId = roleIds.get(value);
    if (!appRoleId) throw new Error(`Microsoft Graph has no app role ${value}`);
    if (current.has(appRoleId)) continue;
    await msCall(deps.fetch, {
      method: "POST",
      url: `${GRAPH_BASE}/servicePrincipals/${encodeURIComponent(servicePrincipalId)}/appRoleAssignments`,
      token,
      body: { principalId: servicePrincipalId, resourceId, appRoleId },
      label: `grant ${value}`,
    });
  }
}

/** The org catalog id of the app whose manifest id is `externalId`, if it is published. */
export async function findCatalogApp(
  deps: Pick<AgentTeamsAppsDeps, "fetch">,
  token: string,
  externalId: string,
): Promise<string | undefined> {
  const filter = encodeURIComponent(`externalId eq '${externalId}'`);
  const found = await msCall(deps.fetch, {
    method: "GET",
    url: `${GRAPH_BASE}/appCatalogs/teamsApps?$filter=${filter}`,
    token,
    label: "read Teams catalog app",
  });
  const first = asRecord(Array.isArray(found.json?.value) ? found.json.value[0] : undefined);
  return stringField(first, "id") || undefined;
}

/**
 * Records that the agent's app was installed in a chat. The first install proves the app is
 * visible in Teams, so it also sets `firstInstalledAt`.
 */
export async function recordChatInstall(
  agentAppId: string,
  chatId: string,
  installId: string | null,
  deps: { store: AgentTeamsAppStore; now?: () => Date },
): Promise<AgentTeamsChatRecord> {
  const at = (deps.now ?? (() => new Date()))();
  const app = await deps.store.findById(agentAppId);
  if (!app) throw new Error(`No agent Teams app ${agentAppId}`);
  const chat = await deps.store.upsertChat({ agentAppId, chatId, installId, installedAt: at });
  if (!app.firstInstalledAt) {
    await deps.store.update(agentAppId, {
      firstInstalledAt: at,
      ...(app.status === "published" ? { status: "visible" } : {}),
    });
  }
  return chat;
}

/** Marks the agent app as gone from one chat (uninstalled, or the chat was left). */
export async function markChatRemoved(
  agentAppId: string,
  chatId: string,
  deps: { store: AgentTeamsAppStore; now?: () => Date },
): Promise<void> {
  const chat = (await deps.store.activeChats(agentAppId)).find((row) => row.chatId === chatId);
  if (chat) await deps.store.markChatRemoved(chat.id, (deps.now ?? (() => new Date()))());
}
