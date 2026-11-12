import type { PrismaClient } from "@rakazo/db";
import { getLogger } from "@rakazo/logging";
import type { AgentTeamsDeprovisionDeps, AgentTeamsSecretStore } from "./deps.js";
import type { FetchLike, TokenProvider } from "./http.js";
import {
  asRecord,
  clientCredentialTokens,
  MicrosoftApiError,
  msCall,
  stringField,
} from "./http.js";
import { prismaAgentTeamsAppStore } from "./store.js";

export const KEY_VAULT_SCOPE = "https://vault.azure.net/.default";
const KEY_VAULT_API_VERSION = "7.4";
const DEFAULT_SERVICE_URL = "https://smba.trafficmanager.net/teams/";
/** Delegated scopes the catalog token is refreshed with. */
const CATALOG_SCOPES = "https://graph.microsoft.com/AppCatalog.ReadWrite.All offline_access";

export function teamsAgentAppsEnabled(env: Record<string, string | undefined>): boolean {
  return env.TEAMS_AGENT_APPS_ENABLED === "true";
}

/** Azure Key Vault secrets over REST, with the provisioner's token. */
export function keyVaultSecretStore(options: {
  vaultUrl: string;
  token: TokenProvider;
  fetch: FetchLike;
}): AgentTeamsSecretStore {
  const base = options.vaultUrl.replace(/\/+$/, "");
  const url = (name: string) =>
    `${base}/secrets/${encodeURIComponent(name)}?api-version=${KEY_VAULT_API_VERSION}`;
  return {
    async set(name, value) {
      await msCall(options.fetch, {
        method: "PUT",
        url: url(name),
        token: await options.token(KEY_VAULT_SCOPE),
        body: { value, contentType: "teams-agent-app-password" },
        label: "store Key Vault secret",
      });
    },
    async get(name) {
      const found = await msCall(options.fetch, {
        method: "GET",
        url: url(name),
        token: await options.token(KEY_VAULT_SCOPE),
        label: "read Key Vault secret",
      });
      const value = stringField(found.json, "value");
      if (!value) throw new Error(`Key Vault secret ${name} is empty`);
      return value;
    },
    async delete(name) {
      await msCall(options.fetch, {
        method: "DELETE",
        url: url(name),
        token: await options.token(KEY_VAULT_SCOPE),
        allow: [404],
        label: "delete Key Vault secret",
      });
    },
  };
}

/**
 * A delegated Teams admin Graph token, refreshed from a refresh token kept in Key Vault. A
 * rotated refresh token is written back, so the next call keeps working.
 */
export function refreshTokenCatalogToken(options: {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  refreshTokenSecret: string;
  secrets: AgentTeamsSecretStore;
  fetch: FetchLike;
}): () => Promise<string> {
  return async () => {
    const refreshToken = await options.secrets.get(options.refreshTokenSecret);
    const response = await options.fetch(
      `https://login.microsoftonline.com/${encodeURIComponent(options.tenantId)}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "content-type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          grant_type: "refresh_token",
          client_id: options.clientId,
          client_secret: options.clientSecret,
          refresh_token: refreshToken,
          scope: CATALOG_SCOPES,
        }).toString(),
      },
    );
    const text = await response.text().catch(() => "");
    let body: Record<string, unknown> | undefined;
    try {
      body = asRecord(JSON.parse(text));
    } catch {
      body = undefined;
    }
    const token = stringField(body, "access_token");
    if (!response.ok || !token) {
      const code = stringField(body, "error") || undefined;
      throw new MicrosoftApiError(
        `catalog admin token refresh failed (${response.status}${code ? ` ${code}` : ""})`,
        response.status,
        code,
      );
    }
    const rotated = stringField(body, "refresh_token");
    if (rotated && rotated !== refreshToken) {
      await options.secrets.set(options.refreshTokenSecret, rotated);
    }
    return token;
  };
}

/**
 * Provisioning dependencies from env, or undefined when the flag is off or config is missing.
 * Reads: TEAMS_AGENT_APPS_ENABLED, TEAMS_APP_TENANT_ID, TEAMS_AGENT_APPS_CLIENT_ID,
 * TEAMS_AGENT_APPS_CLIENT_SECRET, TEAMS_AGENT_APPS_SUBSCRIPTION_ID,
 * TEAMS_AGENT_APPS_RESOURCE_GROUP, TEAMS_AGENT_APPS_KEY_VAULT_URL,
 * TEAMS_AGENT_APPS_CATALOG_REFRESH_TOKEN_SECRET, TEAMS_MESSAGING_ENDPOINT, TEAMS_APP_ID,
 * TEAMS_APP_PASSWORD, TEAMS_APP_CATALOG_ID, TEAMS_SERVICE_URL and the optional
 * TEAMS_AGENT_APPS_DEVELOPER_{NAME,WEBSITE_URL,PRIVACY_URL,TERMS_URL}.
 */
export function teamsAgentAppsFromEnv(
  env: Record<string, string | undefined>,
  prisma: PrismaClient,
  options: { fetch?: FetchLike } = {},
): AgentTeamsDeprovisionDeps | undefined {
  if (!teamsAgentAppsEnabled(env)) return undefined;
  const read = (key: string) => env[key]?.trim() ?? "";
  const requiredKeys = [
    "TEAMS_APP_TENANT_ID",
    "TEAMS_AGENT_APPS_CLIENT_ID",
    "TEAMS_AGENT_APPS_CLIENT_SECRET",
    "TEAMS_AGENT_APPS_SUBSCRIPTION_ID",
    "TEAMS_AGENT_APPS_RESOURCE_GROUP",
    "TEAMS_AGENT_APPS_KEY_VAULT_URL",
    "TEAMS_AGENT_APPS_CATALOG_REFRESH_TOKEN_SECRET",
    "TEAMS_MESSAGING_ENDPOINT",
    "TEAMS_APP_ID",
    "TEAMS_APP_PASSWORD",
    "TEAMS_APP_CATALOG_ID",
  ];
  const missing = requiredKeys.filter((key) => !read(key));
  if (missing.length) {
    getLogger().warn("TEAMS_AGENT_APPS_ENABLED is on but config is missing", { missing });
    return undefined;
  }
  const fetchImpl: FetchLike = options.fetch ?? ((input, init) => fetch(input, init));
  const tenantId = read("TEAMS_APP_TENANT_ID");
  const clientId = read("TEAMS_AGENT_APPS_CLIENT_ID");
  const clientSecret = read("TEAMS_AGENT_APPS_CLIENT_SECRET");
  const provisionerToken = clientCredentialTokens({
    tenantId,
    clientId,
    clientSecret: async () => clientSecret,
    fetch: fetchImpl,
  });
  const secrets = keyVaultSecretStore({
    vaultUrl: read("TEAMS_AGENT_APPS_KEY_VAULT_URL"),
    token: provisionerToken,
    fetch: fetchImpl,
  });
  const sharedPassword = read("TEAMS_APP_PASSWORD");
  return {
    store: prismaAgentTeamsAppStore(prisma),
    secrets,
    fetch: fetchImpl,
    config: {
      tenantId,
      subscriptionId: read("TEAMS_AGENT_APPS_SUBSCRIPTION_ID"),
      resourceGroup: read("TEAMS_AGENT_APPS_RESOURCE_GROUP"),
      messagingEndpoint: read("TEAMS_MESSAGING_ENDPOINT"),
      developer: {
        name: read("TEAMS_AGENT_APPS_DEVELOPER_NAME") || "Redstone",
        websiteUrl: read("TEAMS_AGENT_APPS_DEVELOPER_WEBSITE_URL") || "https://managebot.ai",
        privacyUrl:
          read("TEAMS_AGENT_APPS_DEVELOPER_PRIVACY_URL") || "https://managebot.ai/privacy",
        termsOfUseUrl: read("TEAMS_AGENT_APPS_DEVELOPER_TERMS_URL") || "https://managebot.ai/terms",
      },
    },
    provisionerToken,
    catalogToken: refreshTokenCatalogToken({
      tenantId,
      clientId,
      clientSecret,
      refreshTokenSecret: read("TEAMS_AGENT_APPS_CATALOG_REFRESH_TOKEN_SECRET"),
      secrets,
      fetch: fetchImpl,
    }),
    sharedBot: {
      catalogAppId: read("TEAMS_APP_CATALOG_ID"),
      token: clientCredentialTokens({
        tenantId,
        clientId: read("TEAMS_APP_ID").toLowerCase(),
        clientSecret: async () => sharedPassword,
        fetch: fetchImpl,
      }),
      serviceUrl: read("TEAMS_SERVICE_URL") || DEFAULT_SERVICE_URL,
    },
  };
}
