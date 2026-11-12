export type FetchLike = (input: string, init?: RequestInit) => Promise<Response>;

/** An access token for one scope, e.g. `https://graph.microsoft.com/.default`. */
export type TokenProvider = (scope: string) => Promise<string>;

export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
export const GRAPH_SCOPE = "https://graph.microsoft.com/.default";
export const ARM_BASE = "https://management.azure.com";
export const ARM_SCOPE = "https://management.azure.com/.default";
export const BOT_CONNECTOR_SCOPE = "https://api.botframework.com/.default";
/** Microsoft Graph's own service principal app id, the resource of Graph app roles. */
export const MICROSOFT_GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000";

/** Tokens are renewed five minutes before they expire. */
export const TOKEN_EARLY_RENEW_MS = 5 * 60_000;

/**
 * A Graph, ARM or Bot Connector failure. The message carries the status and the service error
 * code and message, never a request body, so it is safe to store as `lastError`.
 */
export class MicrosoftApiError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly code?: string,
    /** Graph's `error.innerError.code`, when present. */
    readonly innerCode?: string,
  ) {
    super(message);
    this.name = "MicrosoftApiError";
  }
}

export function asRecord(value: unknown): Record<string, unknown> | undefined {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined;
}

export function stringField(record: Record<string, unknown> | undefined, key: string): string {
  const value = record?.[key];
  return typeof value === "string" ? value : "";
}

async function readJson(response: Response): Promise<Record<string, unknown> | undefined> {
  const text = await response.text().catch(() => "");
  if (!text) return undefined;
  try {
    return asRecord(JSON.parse(text));
  } catch {
    return undefined;
  }
}

export interface MsRequest {
  method: "GET" | "POST" | "PUT" | "PATCH" | "DELETE";
  url: string;
  token: string;
  /** JSON body, or raw bytes sent with `contentType`. */
  body?: unknown;
  contentType?: string;
  headers?: Record<string, string>;
  /** Statuses that count as success with no body, e.g. 404 on an idempotent delete. */
  allow?: readonly number[];
  /** A label for errors: "create Entra app", "publish catalog app", …. */
  label: string;
}

export interface MsResponse {
  status: number;
  json: Record<string, unknown> | undefined;
}

/** One request; a failure throws `MicrosoftApiError`. No step retries on its own. */
export async function msCall(fetchImpl: FetchLike, request: MsRequest): Promise<MsResponse> {
  const headers: Record<string, string> = {
    authorization: `Bearer ${request.token}`,
    ...request.headers,
  };
  let body: BodyInit | undefined;
  if (request.body instanceof Uint8Array) {
    headers["content-type"] = request.contentType ?? "application/octet-stream";
    body = new Uint8Array(request.body);
  } else if (request.body !== undefined) {
    headers["content-type"] = "application/json";
    body = JSON.stringify(request.body);
  }
  const response = await fetchImpl(request.url, {
    method: request.method,
    headers,
    ...(body !== undefined ? { body } : {}),
  });
  const json = await readJson(response);
  if (response.ok) return { status: response.status, json };
  if (request.allow?.includes(response.status)) return { status: response.status, json: undefined };
  const error = asRecord(json?.error);
  const code = stringField(error, "code") || undefined;
  const detail = stringField(error, "message") || response.statusText || "request failed";
  const inner = asRecord(error?.innerError) ?? asRecord(error?.innererror);
  const innerCode = stringField(inner, "code") || undefined;
  throw new MicrosoftApiError(
    `${request.label} failed (${response.status}${code ? ` ${code}` : ""}): ${detail}`,
    response.status,
    code,
    innerCode,
  );
}

interface CachedToken {
  token: string;
  expiresAt: number;
}

/**
 * Caches one token per scope until `TOKEN_EARLY_RENEW_MS` before it expires. Concurrent callers
 * share one request; a failed request leaves no cache entry.
 */
function cachedTokens(
  request: (scope: string) => Promise<CachedToken>,
  now: () => number,
): TokenProvider {
  const cache = new Map<string, Promise<CachedToken>>();
  return async (scope) => {
    const cached = cache.get(scope);
    if (cached) {
      const entry = await cached.catch(() => undefined);
      if (entry && entry.expiresAt - TOKEN_EARLY_RENEW_MS > now()) return entry.token;
    }
    const pending = request(scope);
    cache.set(scope, pending);
    pending.catch(() => {
      if (cache.get(scope) === pending) cache.delete(scope);
    });
    return (await pending).token;
  };
}

/**
 * Client-credential tokens for one app (an agent's own Entra app, or the shared Bot), cached per
 * scope.
 */
export function clientCredentialTokens(options: {
  tenantId: string;
  clientId: string;
  clientSecret: () => Promise<string>;
  fetch: FetchLike;
  now?: () => number;
}): TokenProvider {
  const now = options.now ?? Date.now;
  return cachedTokens(async (scope) => {
    const response = await options.fetch(
      `https://login.microsoftonline.com/${encodeURIComponent(options.tenantId)}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "content-type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          grant_type: "client_credentials",
          client_id: options.clientId,
          client_secret: await options.clientSecret(),
          scope,
        }).toString(),
      },
    );
    const body = await readJson(response);
    const token = stringField(body, "access_token");
    if (!response.ok || !token) {
      const code = stringField(body, "error") || undefined;
      throw new MicrosoftApiError(
        `token request failed (${response.status}${code ? ` ${code}` : ""})`,
        response.status,
        code,
      );
    }
    const expiresIn = typeof body?.expires_in === "number" ? body.expires_in : 3600;
    return { token, expiresAt: now() + expiresIn * 1000 };
  }, now);
}

/** The only resources the provisioner's managed identity asks tokens for. */
export const MANAGED_IDENTITY_RESOURCES = [
  "https://graph.microsoft.com",
  "https://management.azure.com",
  "https://vault.azure.net",
] as const;
const MANAGED_IDENTITY_API_VERSION = "2019-08-01";

/** `https://graph.microsoft.com/.default` → `https://graph.microsoft.com`. */
export function scopeResource(scope: string): string {
  return scope.replace(/\/\.default$/, "").replace(/\/+$/, "");
}

/**
 * Tokens for a user-assigned managed identity in Azure Container Apps (or App Service): a GET to
 * the platform's IDENTITY_ENDPOINT with its IDENTITY_HEADER. No client secret exists anywhere.
 * Cached per resource until five minutes before expiry. Only Graph, ARM and Key Vault are asked.
 */
export function managedIdentityTokens(options: {
  /** IDENTITY_ENDPOINT, set by the platform. */
  endpoint: string;
  /** IDENTITY_HEADER, set by the platform. */
  header: string;
  /** The user-assigned identity's client id (TEAMS_AGENT_APPS_IDENTITY_CLIENT_ID). */
  clientId: string;
  fetch: FetchLike;
  now?: () => number;
}): TokenProvider {
  const now = options.now ?? Date.now;
  const allowed = new Set<string>(MANAGED_IDENTITY_RESOURCES);
  return cachedTokens(async (scope) => {
    const resource = scopeResource(scope);
    if (!allowed.has(resource)) {
      throw new Error(`managed identity tokens are not issued for ${resource}`);
    }
    const query = new URLSearchParams({
      "api-version": MANAGED_IDENTITY_API_VERSION,
      resource,
      client_id: options.clientId,
    });
    const separator = options.endpoint.includes("?") ? "&" : "?";
    const response = await options.fetch(`${options.endpoint}${separator}${query.toString()}`, {
      method: "GET",
      headers: { "X-IDENTITY-HEADER": options.header },
    });
    const body = await readJson(response);
    const token = stringField(body, "access_token");
    if (!response.ok || !token) {
      const code = stringField(body, "error") || undefined;
      throw new MicrosoftApiError(
        `managed identity token request failed (${response.status}${code ? ` ${code}` : ""})`,
        response.status,
        code,
      );
    }
    return { token, expiresAt: managedIdentityExpiry(body, now()) };
  }, now);
}

/** `expires_on` (epoch seconds, a string or number) or `expires_in`; an hour when neither. */
function managedIdentityExpiry(body: Record<string, unknown> | undefined, at: number): number {
  const expiresOn = Number(body?.expires_on);
  if (Number.isFinite(expiresOn) && expiresOn > 0) return expiresOn * 1000;
  const expiresIn = Number(body?.expires_in);
  if (Number.isFinite(expiresIn) && expiresIn > 0) return at + expiresIn * 1000;
  return at + 3600_000;
}

/**
 * The `oid` claim of an access token: the object id of the principal it was issued to. Only
 * read from the provisioner's own freshly issued token, never from a caller.
 */
export function tokenObjectId(token: string): string {
  const payload = token.split(".")[1];
  if (!payload) throw new Error("access token is not a JWT");
  let claims: Record<string, unknown> | undefined;
  try {
    claims = asRecord(JSON.parse(Buffer.from(payload, "base64url").toString("utf8")));
  } catch {
    claims = undefined;
  }
  const oid = stringField(claims, "oid");
  if (!oid) throw new Error("access token has no oid claim");
  return oid;
}
