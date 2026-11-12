import { getLogger } from "@rakazo/logging";
import { agentEntraAppBody, assignAgentGraphAppRoles } from "./app-role-guard.js";
import type { AgentTeamsAppsDeps } from "./deps.js";
import {
  ARM_BASE,
  ARM_SCOPE,
  asRecord,
  GRAPH_BASE,
  GRAPH_SCOPE,
  MicrosoftApiError,
  msCall,
  stringField,
} from "./http.js";
import { agentIcons, buildAgentTeamsManifest, buildManifestZip } from "./manifest.js";
import { createdRecently, withReplicationRetry } from "./replication.js";
import type {
  AgentTeamsAppPatch,
  AgentTeamsAppRecord,
  AgentTeamsAppStore,
  AgentTeamsChatRecord,
} from "./store.js";

export { AGENT_GRAPH_APP_ROLES } from "./app-role-guard.js";

export const AZURE_BOT_API_VERSION = "2022-09-15";
/** How long one provision or removal holds the row's lease. */
export const AGENT_TEAMS_APP_LOCK_MS = 15 * 60_000;
const DEFAULT_MANIFEST_VERSION = "1.0.0";
const LAST_ERROR_MAX = 500;
/** Statuses provisioning may write to. Anything else (removing, removed) means stop. */
const PROVISIONING_STATUSES = ["creating", "registered"] as const;

/** Another provision or removal holds this agent's row; the job queue retries later. */
export class AgentTeamsAppBusyError extends Error {
  constructor(botId: string) {
    super(`agent Teams app for ${botId} is busy with another operation; retry later`);
    this.name = "AgentTeamsAppBusyError";
  }
}

/** Removal began while provisioning ran; provisioning stopped without touching the status. */
export class AgentTeamsAppRemovalStartedError extends Error {
  constructor() {
    super("agent Teams app removal started; provisioning stopped");
    this.name = "AgentTeamsAppRemovalStartedError";
  }
}

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
 * Runs `fn` holding the row's operation lease, so provision and removal of one agent never
 * interleave. Throws `AgentTeamsAppBusyError` (retryable) when the lease is held.
 */
export async function withAgentAppLock<T>(
  deps: Pick<AgentTeamsAppsDeps, "store" | "now" | "lockMs">,
  app: AgentTeamsAppRecord,
  fn: () => Promise<T>,
): Promise<T> {
  const at = (deps.now ?? (() => new Date()))();
  const until = new Date(at.getTime() + (deps.lockMs ?? AGENT_TEAMS_APP_LOCK_MS));
  if (!(await deps.store.acquireLock(app.id, until, at))) {
    throw new AgentTeamsAppBusyError(app.botId);
  }
  try {
    return await fn();
  } finally {
    await deps.store.releaseLock(app.id).catch((error: unknown) => {
      getLogger().warn("agent Teams app lease release failed", {
        botId: app.botId,
        error: safeErrorMessage(error),
      });
    });
  }
}

/**
 * Gives an agent its own Teams identity. Idempotent and resumable: every step is skipped when
 * the row already holds its result, so a retry after a failure picks up where it stopped.
 * Steps: Entra app → service principal → password in Key Vault → Azure Bot + Teams channel →
 * Graph app roles (guarded) → manifest → org catalog publish. Runs under the row's lease, and
 * every write is a compare-and-set, so it never overwrites `removing` or `removed`.
 */
export async function provisionAgentTeamsApp(
  input: ProvisionAgentTeamsAppInput,
  deps: AgentTeamsAppsDeps,
): Promise<AgentTeamsAppRecord> {
  const now = deps.now ?? (() => new Date());
  const ensured = await deps.store.ensure(input.botId);
  const published = await withAgentAppLock(deps, ensured, async () => {
    let app = (await deps.store.findById(ensured.id)) ?? ensured;
    if (app.status === "removing") {
      throw new Error("This agent's Teams app is being removed; provision it again afterwards.");
    }
    if (app.status === "published" || app.status === "visible") return { app, fresh: false };
    if (app.status === "removed" || app.status === "failed") {
      const reset: AgentTeamsAppPatch =
        app.status === "removed"
          ? {
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
            }
          : { status: "creating", lastError: null };
      const restarted = await deps.store.updateIfStatus(app.id, [app.status], reset);
      if (!restarted) throw new AgentTeamsAppBusyError(input.botId);
      app = restarted;
    }
    try {
      return { app: await runProvision(app, input, deps, now), fresh: true };
    } catch (error) {
      // A removal that started meanwhile owns the status; only a live provision is marked failed.
      if (!(error instanceof AgentTeamsAppRemovalStartedError)) {
        await deps.store.updateIfStatus(app.id, PROVISIONING_STATUSES, {
          status: "failed",
          lastError: safeErrorMessage(error),
        });
      }
      throw error;
    }
  });
  if (published.fresh && deps.onPublished) {
    await deps.onPublished(published.app).catch((error: unknown) => {
      getLogger().warn("agent Teams app published but platform refresh failed", {
        botId: published.app.botId,
        error: safeErrorMessage(error),
      });
    });
  }
  return published.app;
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
  const save = async (patch: AgentTeamsAppPatch) => {
    const saved = await store.updateIfStatus(app.id, PROVISIONING_STATUSES, patch);
    if (!saved) throw new AgentTeamsAppRemovalStartedError();
    return saved;
  };

  // 1. Entra app, single-tenant, tagged. The unique name makes a repeated create return the same
  // app; the provisioner becomes its owner, which the role guard later checks.
  // Objects this run created (or the row created within the replication window) may not be
  // visible to Graph yet; calls on them retry replication-lag 404s only.
  let justCreated = createdRecently(app.createdAt, now());
  const retry = <T>(fn: () => Promise<T>) =>
    withReplicationRetry(fn, {
      recent: justCreated,
      ...(deps.sleep ? { sleep: deps.sleep } : {}),
      ...(deps.replicationDelaysMs ? { delays: deps.replicationDelaysMs } : {}),
    });
  if (!app.entraAppId || !app.entraObjectId) {
    const created = await upsertEntraApp(deps, await graph(), input);
    justCreated = true;
    app = await save({ entraAppId: created.appId, entraObjectId: created.objectId });
  }
  const entraAppId = required(app.entraAppId, "entraAppId");
  const entraObjectId = required(app.entraObjectId, "entraObjectId");

  // 2. Service principal (needed for tokens and app role grants).
  const sp = await retry(async () => ensureServicePrincipal(deps, await graph(), entraAppId));
  if (sp.created) justCreated = true;

  // 3. Password credential, stored only in Key Vault.
  if (!app.secretRef) {
    const added = await retry(async () =>
      msCall(deps.fetch, {
        method: "POST",
        url: `${GRAPH_BASE}/applications/${encodeURIComponent(entraObjectId)}/addPassword`,
        token: await graph(),
        body: { passwordCredential: { displayName: "rakazo-teams-bot" } },
        label: "add Entra app password",
      }),
    );
    const secretText = stringField(added.json, "secretText");
    if (!secretText) throw new Error("add Entra app password returned no secret");
    const secretRef = agentSecretName(input.botId);
    await deps.secrets.set(secretRef, secretText);
    app = await save({ secretRef });
  }

  // 4. Azure Bot with the Teams channel. PUT is idempotent, so both run before the name is saved.
  if (!app.azureBotName) {
    const azureBotName = agentAzureBotName(input.botId);
    const arm = await deps.provisionerToken(ARM_SCOPE);
    const botUrl = azureBotUrl(deps, azureBotName);
    await retry(() =>
      msCall(deps.fetch, {
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
      }),
    );
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
    app = await save({ azureBotName, status: "registered" });
  }

  // 5. Graph app roles for the agent's own service principal, only through the guard.
  await assignAgentGraphAppRoles(deps, input.botId, { justCreated });

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
    app = await save({
      catalogAppId,
      manifestVersion: version,
      publishedAt: now(),
      status: "published",
      lastError: null,
    });
  } else {
    app = await save({
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
    // Single-tenant, tagged, "Rakazo agent: <name>", and no requested API access.
    body: agentEntraAppBody(input.name),
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
): Promise<{ id: string; created: boolean }> {
  const found = await msCall(deps.fetch, {
    method: "GET",
    url: `${GRAPH_BASE}/servicePrincipals(appId='${encodeURIComponent(appId)}')?$select=id`,
    token,
    allow: [404],
    label: "read service principal",
  });
  const existing = stringField(found.json, "id");
  if (existing) return { id: existing, created: false };
  const created = await msCall(deps.fetch, {
    method: "POST",
    url: `${GRAPH_BASE}/servicePrincipals`,
    token,
    body: { appId },
    label: "create service principal",
  });
  const id = stringField(created.json, "id");
  if (!id) throw new Error("create service principal returned no id");
  return { id, created: true };
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
 * visible in Teams, so it also sets `firstInstalledAt`. An app being removed records nothing:
 * the removal has already taken its list of chats.
 */
export async function recordChatInstall(
  agentAppId: string,
  chatId: string,
  installId: string | null,
  deps: { store: AgentTeamsAppStore; now?: () => Date },
): Promise<AgentTeamsChatRecord | null> {
  const at = (deps.now ?? (() => new Date()))();
  const app = await deps.store.findById(agentAppId);
  if (!app) throw new Error(`No agent Teams app ${agentAppId}`);
  if (app.status === "removing" || app.status === "removed") return null;
  const chat = await deps.store.upsertChat({ agentAppId, chatId, installId, installedAt: at });
  if (!app.firstInstalledAt) {
    // Compare-and-set, so an install seen mid-removal never turns `removing` back into `visible`.
    const moved = await deps.store.updateIfStatus(agentAppId, ["published"], {
      firstInstalledAt: at,
      status: "visible",
    });
    if (!moved) {
      await deps.store.updateIfStatus(agentAppId, ["creating", "registered", "visible"], {
        firstInstalledAt: at,
      });
    }
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
