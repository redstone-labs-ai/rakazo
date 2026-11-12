import { recordSecurityEvent } from "@rakazo/db";
import type { AgentTeamsAppsDeps } from "./deps.js";
import {
  asRecord,
  GRAPH_BASE,
  GRAPH_SCOPE,
  MICROSOFT_GRAPH_APP_ID,
  MicrosoftApiError,
  msCall,
  stringField,
} from "./http.js";
import { createdRecently, withReplicationRetry } from "./replication.js";
import type { AgentTeamsAppRecord } from "./store.js";

/**
 * The guard around the provisioner's most dangerous Graph permission, AppRoleAssignment.ReadWrite.All.
 * That permission could grant any app role on any API to any service principal; this module is
 * the ONLY code that calls the app-role-assignment endpoint, and it refuses everything except:
 *
 * - resource: the Microsoft Graph service principal (appId 00000003-0000-0000-c000-000000000000),
 * - roles: exactly `Chat.Create` and `TeamsAppInstallation.ReadWriteAndConsentSelfForChat.All`,
 *   resolved by value from Graph's own `appRoles` (application permissions only; never a delegated
 *   oauth2 grant),
 * - principal: the service principal of an Entra app recorded in `agent_teams_apps.entraAppId`,
 *   created by this flow (tag, name prefix, single-tenant, no requested API access) and owned by
 *   the provisioner identity, as Graph confirms through `/applications/{id}/owners`,
 * - and never the provisioner itself, the shared Bot app (TEAMS_APP_ID) or Graph.
 *
 * Every static check runs before any Graph call; every assignment and refusal is logged as a
 * `security.event` with ids only, never a token.
 */

/** Graph application permissions an agent app may hold: create chats and install itself in them. */
export const AGENT_GRAPH_APP_ROLES = [
  "Chat.Create",
  "TeamsAppInstallation.ReadWriteAndConsentSelfForChat.All",
] as const;
export type AgentGraphAppRole = (typeof AGENT_GRAPH_APP_ROLES)[number];

/** Every Entra app this flow creates carries this tag. */
export const AGENT_APP_TAG = "rakazo-agent-teams-app";
/** …and a display name with this prefix, so it is identifiable in the Entra portal. */
export const AGENT_APP_DISPLAY_PREFIX = "Rakazo agent: ";
const DISPLAY_NAME_MAX = 120;

/** Statuses during which provisioning may grant roles. */
const GRANTABLE_STATUSES = ["creating", "registered"] as const;

/** Why the guard refused, for logs and errors; never carries a secret. */
export class AppRoleGuardRefusal extends Error {
  constructor(readonly reason: string) {
    super(`app role assignment refused: ${reason}`);
    this.name = "AppRoleGuardRefusal";
  }
}

type Check = { ok: true } | { ok: false; reason: string };
const ok: Check = { ok: true };
const refuse = (reason: string): Check => ({ ok: false, reason });

const sameId = (a: string, b: string) => a.trim().toLowerCase() === b.trim().toLowerCase();

/** What a caller asks to grant. `kind: "delegated"` is an oauth2 permission grant. */
export interface AppRoleRequest {
  kind: "application" | "delegated";
  resourceAppId: string;
  role: string;
}

/** The allowlist: Graph, application kind, one of the two roles. Pure. */
export function checkRoleRequest(request: AppRoleRequest): Check {
  if (request.kind !== "application") return refuse("only application roles may be granted");
  if (!sameId(request.resourceAppId, MICROSOFT_GRAPH_APP_ID)) {
    return refuse("resource is not Microsoft Graph");
  }
  if (!(AGENT_GRAPH_APP_ROLES as readonly string[]).includes(request.role)) {
    return refuse(`role ${request.role} is not allowed`);
  }
  return ok;
}

/** Identities that must never receive a role, whatever the store says. */
export interface ProtectedIdentities {
  /** The provisioner managed identity's client id. */
  provisionerClientId: string;
  /** The provisioner managed identity's service principal object id. */
  provisionerObjectId?: string;
  /** TEAMS_APP_ID, the shared Bot app. */
  sharedBotAppId: string;
}

/** The target's app id is recorded by this flow and is none of the protected identities. Pure. */
export function checkTargetAppId(
  record: Pick<AgentTeamsAppRecord, "entraAppId" | "entraObjectId" | "status">,
  protectedIds: ProtectedIdentities,
): Check {
  const appId = record.entraAppId?.trim() ?? "";
  if (!appId || !record.entraObjectId) return refuse("no Entra app is recorded for this agent");
  if (!(GRANTABLE_STATUSES as readonly string[]).includes(record.status)) {
    return refuse(`agent app is ${record.status}`);
  }
  if (sameId(appId, protectedIds.provisionerClientId)) return refuse("target is the provisioner");
  if (protectedIds.sharedBotAppId && sameId(appId, protectedIds.sharedBotAppId)) {
    return refuse("target is the shared Bot app");
  }
  if (sameId(appId, MICROSOFT_GRAPH_APP_ID)) return refuse("target is Microsoft Graph");
  return ok;
}

/** The Entra app body this flow creates. Single-tenant, tagged, asks for no API access. */
export function agentEntraAppBody(name: string): {
  displayName: string;
  signInAudience: "AzureADMyOrg";
  tags: string[];
  requiredResourceAccess: [];
} {
  const clean = name.replace(/\s+/g, " ").trim() || "agent";
  return {
    displayName: `${AGENT_APP_DISPLAY_PREFIX}${clean}`.slice(0, DISPLAY_NAME_MAX),
    signInAudience: "AzureADMyOrg",
    tags: [AGENT_APP_TAG],
    requiredResourceAccess: [],
  };
}

/** An Entra app as Graph returns it: the fields the guard reads. */
export interface EntraAppShape {
  id?: string;
  appId?: string;
  displayName?: string;
  signInAudience?: string;
  tags?: unknown;
  requiredResourceAccess?: unknown;
}

/** The app has the shape this flow creates, and is the one recorded. Pure. */
export function checkEntraAppShape(
  app: EntraAppShape,
  expected: { appId: string; objectId: string },
): Check {
  if (!app.appId || !sameId(app.appId, expected.appId))
    return refuse("Entra app id does not match the record");
  if (app.id && !sameId(app.id, expected.objectId))
    return refuse("Entra object id does not match the record");
  if (app.signInAudience !== "AzureADMyOrg") return refuse("Entra app is not single-tenant");
  if (!Array.isArray(app.tags) || !app.tags.includes(AGENT_APP_TAG)) {
    return refuse("Entra app was not created by this flow (tag missing)");
  }
  if (!app.displayName?.startsWith(AGENT_APP_DISPLAY_PREFIX)) {
    return refuse("Entra app was not created by this flow (name prefix missing)");
  }
  if (
    Array.isArray(app.requiredResourceAccess)
      ? app.requiredResourceAccess.length > 0
      : app.requiredResourceAccess !== undefined && app.requiredResourceAccess !== null
  ) {
    return refuse("Entra app requests API access");
  }
  return ok;
}

/** Graph lists the provisioner among the app's owners. Pure. */
export function checkOwnership(ownerIds: readonly string[], provisionerObjectId: string): Check {
  if (!provisionerObjectId) return refuse("provisioner object id is unknown");
  return ownerIds.some((id) => sameId(id, provisionerObjectId))
    ? ok
    : refuse("provisioner is not an owner of the Entra app");
}

/** The service principal belongs to the recorded app and is not the provisioner. Pure. */
export function checkServicePrincipal(
  sp: { id: string; appId: string },
  expectedAppId: string,
  protectedIds: ProtectedIdentities,
): Check {
  if (!sp.id) return refuse("service principal not found");
  if (!sameId(sp.appId, expectedAppId)) return refuse("service principal belongs to another app");
  if (protectedIds.provisionerObjectId && sameId(sp.id, protectedIds.provisionerObjectId)) {
    return refuse("target is the provisioner");
  }
  return ok;
}

/** Resolves the allowed roles' ids from Graph's own appRoles. Pure. */
export function resolveGraphRoles(
  graphSp: Record<string, unknown> | undefined,
):
  | { ok: true; resourceId: string; roles: Map<AgentGraphAppRole, string> }
  | { ok: false; reason: string } {
  const resourceId = stringField(graphSp, "id");
  if (!resourceId || !sameId(stringField(graphSp, "appId"), MICROSOFT_GRAPH_APP_ID)) {
    return { ok: false, reason: "resource is not the Microsoft Graph service principal" };
  }
  const roles = new Map<AgentGraphAppRole, string>();
  for (const entry of Array.isArray(graphSp?.appRoles) ? graphSp.appRoles : []) {
    const role = asRecord(entry);
    const value = stringField(role, "value");
    const id = stringField(role, "id");
    const memberTypes = Array.isArray(role?.allowedMemberTypes) ? role.allowedMemberTypes : [];
    if (!id || role?.isEnabled === false || !memberTypes.includes("Application")) continue;
    const allowed = AGENT_GRAPH_APP_ROLES.find((name) => name === value);
    if (allowed) roles.set(allowed, id);
  }
  for (const name of AGENT_GRAPH_APP_ROLES) {
    if (!roles.has(name))
      return { ok: false, reason: `Microsoft Graph has no application role ${name}` };
  }
  return { ok: true, resourceId, roles };
}

/** A security event as this module records it: ids and outcomes only. */
export interface AgentAppSecurityEvent {
  type: "teams.agent_app.role_assigned" | "teams.agent_app.role_refused";
  metadata: Record<string, string | null>;
}

export type GuardDeps = Pick<
  AgentTeamsAppsDeps,
  | "store"
  | "fetch"
  | "provisionerToken"
  | "identity"
  | "config"
  | "securityEvent"
  | "now"
  | "sleep"
  | "replicationDelaysMs"
>;

export interface AssignAgentGraphAppRolesOptions {
  /**
   * The caller just created the agent's Entra app or service principal, so Graph may not see it
   * yet. Reads and writes on it then retry replication-lag 404s; checks are unchanged.
   */
  justCreated?: boolean;
}

function record(deps: GuardDeps, event: AgentAppSecurityEvent): Promise<void> {
  const write =
    deps.securityEvent ??
    ((e: AgentAppSecurityEvent) => recordSecurityEvent({ type: e.type, metadata: e.metadata }));
  return write(event).catch(() => undefined);
}

/**
 * Grants the agent its two Graph application roles, through every guard. Idempotent: roles it
 * already holds are skipped. Throws `AppRoleGuardRefusal` on any refusal.
 */
export async function assignAgentGraphAppRoles(
  deps: GuardDeps,
  botId: string,
  options: AssignAgentGraphAppRolesOptions = {},
): Promise<void> {
  const base = {
    botId,
    entraAppId: null as string | null,
    servicePrincipalId: null as string | null,
  };
  const refused = async (reason: string): Promise<never> => {
    await record(deps, { type: "teams.agent_app.role_refused", metadata: { ...base, reason } });
    throw new AppRoleGuardRefusal(reason);
  };
  const enforce = async (check: Check) => {
    if (!check.ok) await refused(check.reason);
  };

  // Static checks: no Graph call happens unless every one passes.
  const app = await deps.store.findByBotId(botId);
  if (!app) return refused("no agent app is recorded for this bot");
  base.entraAppId = app.entraAppId;
  const protectedIds: ProtectedIdentities = {
    provisionerClientId: deps.identity.clientId,
    sharedBotAppId: deps.config.sharedBotAppId,
  };
  await enforce(checkTargetAppId(app, protectedIds));
  for (const role of AGENT_GRAPH_APP_ROLES) {
    await enforce(
      checkRoleRequest({ kind: "application", resourceAppId: MICROSOFT_GRAPH_APP_ID, role }),
    );
  }
  const entraAppId = app.entraAppId ?? "";
  const entraObjectId = app.entraObjectId ?? "";
  protectedIds.provisionerObjectId = await deps.identity.objectId();
  if (!protectedIds.provisionerObjectId) return refused("provisioner object id is unknown");

  const token = await deps.provisionerToken(GRAPH_SCOPE);
  const get = (url: string, label: string, allow?: number[]) =>
    msCall(deps.fetch, { method: "GET", url, token, label, ...(allow ? { allow } : {}) });
  // Only the agent's own just-created app and principal get replication-lag retries.
  const recent =
    options.justCreated === true ||
    createdRecently(app.createdAt, (deps.now ?? (() => new Date()))());
  const retry = <T>(fn: () => Promise<T>) =>
    withReplicationRetry(fn, {
      recent,
      ...(deps.sleep ? { sleep: deps.sleep } : {}),
      ...(deps.replicationDelaysMs ? { delays: deps.replicationDelaysMs } : {}),
    });

  // The app is the one recorded, has this flow's shape, and the provisioner owns it.
  const entraApp = await retry(() =>
    get(
      `${GRAPH_BASE}/applications/${encodeURIComponent(entraObjectId)}?$select=id,appId,displayName,signInAudience,tags,requiredResourceAccess`,
      "read agent Entra app",
    ),
  );
  const shape = entraApp.json;
  await enforce(
    checkEntraAppShape(
      {
        id: stringField(shape, "id"),
        appId: stringField(shape, "appId"),
        displayName: stringField(shape, "displayName"),
        signInAudience: stringField(shape, "signInAudience"),
        tags: shape?.tags,
        requiredResourceAccess: shape?.requiredResourceAccess,
      },
      { appId: entraAppId, objectId: entraObjectId },
    ),
  );
  const owners = await retry(() =>
    get(
      `${GRAPH_BASE}/applications/${encodeURIComponent(entraObjectId)}/owners?$select=id`,
      "read agent Entra app owners",
    ),
  );
  const ownerIds = (Array.isArray(owners.json?.value) ? owners.json.value : []).map((owner) =>
    stringField(asRecord(owner), "id"),
  );
  await enforce(checkOwnership(ownerIds, protectedIds.provisionerObjectId));

  // The principal is that app's own service principal.
  // While replicating, a 404 is retried; once the budget is spent (or when not recent) a 404
  // reads as "not found" and the check below refuses.
  const spUrl = `${GRAPH_BASE}/servicePrincipals(appId='${encodeURIComponent(entraAppId)}')?$select=id,appId`;
  const sp = await retry(() => get(spUrl, "read agent service principal")).catch(
    (error: unknown) => {
      if (error instanceof MicrosoftApiError && error.status === 404) {
        return { status: 404, json: undefined };
      }
      throw error;
    },
  );
  const principal = { id: stringField(sp.json, "id"), appId: stringField(sp.json, "appId") };
  base.servicePrincipalId = principal.id || null;
  await enforce(checkServicePrincipal(principal, entraAppId, protectedIds));

  // The resource is Graph, and role ids come from Graph's own application roles.
  const graphSp = await get(
    `${GRAPH_BASE}/servicePrincipals(appId='${MICROSOFT_GRAPH_APP_ID}')?$select=id,appId,appRoles`,
    "read Microsoft Graph service principal",
  );
  const resolved = resolveGraphRoles(graphSp.json);
  if (!resolved.ok) return refused(resolved.reason);

  const assignmentsUrl = `${GRAPH_BASE}/servicePrincipals/${encodeURIComponent(principal.id)}/appRoleAssignments`;
  const assigned = await retry(() => get(assignmentsUrl, "read app role assignments"));
  const current = new Set(
    (Array.isArray(assigned.json?.value) ? assigned.json.value : [])
      .map((entry) => asRecord(entry))
      .filter((entry) => sameId(stringField(entry, "resourceId"), resolved.resourceId))
      .map((entry) => stringField(entry, "appRoleId")),
  );
  for (const role of AGENT_GRAPH_APP_ROLES) {
    const appRoleId = resolved.roles.get(role) ?? "";
    if (current.has(appRoleId)) continue;
    await retry(() =>
      msCall(deps.fetch, {
        method: "POST",
        url: assignmentsUrl,
        token,
        body: { principalId: principal.id, resourceId: resolved.resourceId, appRoleId },
        label: `grant ${role}`,
      }),
    );
    await record(deps, {
      type: "teams.agent_app.role_assigned",
      metadata: { ...base, role, appRoleId, resourceId: resolved.resourceId },
    });
  }
}
