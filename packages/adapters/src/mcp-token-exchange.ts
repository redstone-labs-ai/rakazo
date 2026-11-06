import type { OAuthTokens } from "@modelcontextprotocol/sdk/shared/auth.js";
import { getLogger } from "@rakazo/logging";
import type { McpGrantIssuer } from "./mcp-oauth.js";

const TOKEN_EXCHANGE = "urn:ietf:params:oauth:grant-type:token-exchange";
const ACCESS_TOKEN_TYPE = "urn:ietf:params:oauth:token-type:access_token";
const REQUEST_TIMEOUT_MS = 15_000;
const MAX_ATTEMPTS = 3;

/** An authorization server that issues MCP grants for the person's SSO access token (RFC 8693). */
export type McpTokenExchangeConfig = {
  /** The authorization server's issuer, as the MCP servers' resource metadata names it. */
  issuer: string;
  tokenEndpoint: string;
  clientId: string;
  clientSecret: string;
  /** The scope of the SSO access token the server accepts as the subject token. */
  subjectScope: string;
  /** A response field naming the person's role at the service, reported to `onRole`. */
  roleField?: string;
};

/** Issuer URLs compared as resource metadata and configuration may spell them. */
function sameIssuer(left: string, right: string): boolean {
  const normalize = (value: string) => {
    try {
      return new URL(value).href.replace(/\/+$/, "");
    } catch {
      return value;
    }
  };
  return normalize(left) === normalize(right);
}

export type McpTokenExchangeOptions = {
  /** An SSO access token for the person with the given scope, if they have one. */
  subjectToken: (userId: string, scope: string) => Promise<string | undefined>;
  /** The person's role at the service, from a response's `roleField`. */
  onRole?: (userId: string, issuer: string, role: string) => Promise<void>;
  fetch?: typeof fetch;
  /** Waits between attempts after a rate limit; replaced in tests. */
  sleep?: (ms: number) => Promise<void>;
};

/**
 * Issues MCP grants by exchanging the person's SSO access token at authorization servers the
 * deployment holds a confidential client with. Requests to one token endpoint run one at a
 * time, and a rate-limited request waits and tries again, since every request comes from the
 * same deployment.
 */
export class McpTokenExchange implements McpGrantIssuer {
  private readonly queues = new Map<string, Promise<unknown>>();

  constructor(
    private readonly configs: McpTokenExchangeConfig[],
    private readonly options: McpTokenExchangeOptions,
  ) {}

  private configFor(authorizationServerUrl: string) {
    return this.configs.find((config) => sameIssuer(config.issuer, authorizationServerUrl));
  }

  clientFor(authorizationServerUrl: string) {
    const config = this.configFor(authorizationServerUrl);
    return config
      ? {
          clientId: config.clientId,
          clientSecret: config.clientSecret,
          authMethod: "client_secret_post" as const,
        }
      : undefined;
  }

  async issue(input: {
    userId: string;
    authorizationServerUrl: string;
    resource?: string;
  }): Promise<OAuthTokens | undefined> {
    const config = this.configFor(input.authorizationServerUrl);
    if (!config) return undefined;
    const subject = await this.options.subjectToken(input.userId, config.subjectScope);
    if (!subject) return undefined;
    const result = await this.serially(config.tokenEndpoint, () =>
      this.exchange(config, subject, input.resource),
    );
    if (config.roleField && typeof result.response[config.roleField] === "string")
      await this.options
        .onRole?.(input.userId, config.issuer, result.response[config.roleField] as string)
        .catch((error: unknown) => getLogger().error("MCP token exchange role not stored", error));
    return result.tokens;
  }

  private serially<T>(key: string, task: () => Promise<T>): Promise<T> {
    const previous = this.queues.get(key) ?? Promise.resolve();
    const next = previous.catch(() => undefined).then(task);
    const tail = next.catch(() => undefined);
    this.queues.set(key, tail);
    void tail.then(() => {
      if (this.queues.get(key) === tail) this.queues.delete(key);
    });
    return next;
  }

  private async exchange(
    config: McpTokenExchangeConfig,
    subjectToken: string,
    resource: string | undefined,
  ): Promise<{ tokens: OAuthTokens; response: Record<string, unknown> }> {
    const body = new URLSearchParams({
      grant_type: TOKEN_EXCHANGE,
      subject_token: subjectToken,
      subject_token_type: ACCESS_TOKEN_TYPE,
      client_id: config.clientId,
      client_secret: config.clientSecret,
    });
    if (resource) body.set("resource", resource);
    const fetchFn = this.options.fetch ?? fetch;
    const sleep =
      this.options.sleep ?? ((ms: number) => new Promise((resolve) => setTimeout(resolve, ms)));
    for (let attempt = 1; ; attempt++) {
      const response = await fetchFn(config.tokenEndpoint, {
        method: "POST",
        headers: {
          "content-type": "application/x-www-form-urlencoded",
          accept: "application/json",
        },
        body,
        redirect: "error",
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
      if (response.status === 429 && attempt < MAX_ATTEMPTS) {
        const retryAfter = Number(response.headers.get("retry-after"));
        await sleep(
          Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter * 1000 : 2000 * attempt,
        );
        continue;
      }
      const json: unknown = await response.json().catch(() => undefined);
      const fields = (json && typeof json === "object" ? json : {}) as Record<string, unknown>;
      if (!response.ok || typeof fields.access_token !== "string") {
        // The error code and description are the server's protocol answer, never a credential.
        const code = typeof fields.error === "string" ? fields.error : `HTTP ${response.status}`;
        const description =
          typeof fields.error_description === "string" ? `: ${fields.error_description}` : "";
        throw new Error(`MCP token exchange refused: ${code}${description}`.slice(0, 300));
      }
      const tokens: OAuthTokens = {
        access_token: fields.access_token,
        token_type: typeof fields.token_type === "string" ? fields.token_type : "Bearer",
        ...(typeof fields.expires_in === "number" ? { expires_in: fields.expires_in } : {}),
        ...(typeof fields.refresh_token === "string"
          ? { refresh_token: fields.refresh_token }
          : {}),
        ...(typeof fields.scope === "string" ? { scope: fields.scope } : {}),
      };
      return { tokens, response: fields };
    }
  }
}
