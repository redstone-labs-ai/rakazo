import type { AgentAppSecurityEvent } from "./app-role-guard.js";
import type { FetchLike, TokenProvider } from "./http.js";
import type { AgentManifestDeveloper } from "./manifest.js";
import type { AgentTeamsAppRecord, AgentTeamsAppStore } from "./store.js";

/** Where agent bot secrets live (Azure Key Vault in prod). Values never reach the database. */
export interface AgentTeamsSecretStore {
  set(name: string, value: string): Promise<void>;
  get(name: string): Promise<string>;
  /** Idempotent: a missing secret is not an error. */
  delete(name: string): Promise<void>;
}

export interface AgentTeamsAppsConfig {
  tenantId: string;
  subscriptionId: string;
  resourceGroup: string;
  /** The webhook every Teams bot posts to: the same URL the shared bot uses. */
  messagingEndpoint: string;
  developer: AgentManifestDeveloper;
  /** App package version for new manifests. */
  manifestVersion?: string;
  /** TEAMS_APP_ID, the shared Bot app: never a target of an app role grant. */
  sharedBotAppId: string;
}

/** The provisioner's own identity (a user-assigned managed identity in production). */
export interface AgentTeamsProvisionerIdentity {
  /** Client id (TEAMS_AGENT_APPS_IDENTITY_CLIENT_ID). */
  clientId: string;
  /** Its service principal's object id, which Graph lists as owner of every agent app. */
  objectId: () => Promise<string>;
}

export interface AgentTeamsAppsDeps {
  store: AgentTeamsAppStore;
  secrets: AgentTeamsSecretStore;
  fetch: FetchLike;
  config: AgentTeamsAppsConfig;
  /**
   * The provisioner's own tokens (managed identity), for Graph (apps, service principals,
   * guarded app role grants), ARM (Azure Bot) and Key Vault.
   */
  provisionerToken: TokenProvider;
  identity: AgentTeamsProvisionerIdentity;
  /** Records a guard `security.event`; defaults to `recordSecurityEvent`. */
  securityEvent?: (event: AgentAppSecurityEvent) => Promise<void>;
  /** Operation lease length; default 15 minutes. */
  lockMs?: number;
  /** A delegated Teams admin token for Graph: the org catalog takes no app-only publish. */
  catalogToken: () => Promise<string>;
  now?: () => Date;
  /**
   * Called after provisioning reaches `published` (the router's `refreshAgentTeamsPlatforms`),
   * so the new bot starts answering without a restart. A failure here is logged, not fatal.
   */
  onPublished?: (app: AgentTeamsAppRecord) => Promise<void>;
}

export interface AgentTeamsDeprovisionDeps extends AgentTeamsAppsDeps {
  /** The shared "Bot" Teams app that takes over each chat the agent leaves. */
  sharedBot: {
    /** TEAMS_APP_CATALOG_ID. */
    catalogAppId: string;
    /** The shared bot's client-credential tokens (TEAMS_APP_ID / TEAMS_APP_PASSWORD). */
    token: TokenProvider;
    /** Bot Connector service URL (TEAMS_SERVICE_URL). */
    serviceUrl: string;
  };
  /** The agent app's own tokens. Defaults to its Entra app id plus the Key Vault secret. */
  agentToken?: (entraAppId: string, secretRef: string) => TokenProvider;
}
