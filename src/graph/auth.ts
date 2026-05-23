import type { NolendarConfig } from "../domain/config.js";

export const DEFAULT_GRAPH_SCOPES = ["Calendars.Read", "User.Read"];

export interface GraphAuthBaseConfig {
  tenantId: string;
  scopes: string[];
  clientId?: string;
}

export interface GraphDeviceCodeAuthConfig extends GraphAuthBaseConfig {
  mode: "device_code";
}

export interface GraphInteractiveBrowserAuthConfig extends GraphAuthBaseConfig {
  mode: "interactive_browser";
}

export interface GraphAuthorizationCodeAuthConfig extends GraphAuthBaseConfig {
  mode: "auth_code";
  clientId: string;
  clientSecret: string;
  redirectUri: string;
}

export interface GraphStaticAccessTokenAuthConfig extends GraphAuthBaseConfig {
  mode: "static_access_token";
  accessToken: string;
}

export type GraphAuthConfig =
  | GraphDeviceCodeAuthConfig
  | GraphInteractiveBrowserAuthConfig
  | GraphAuthorizationCodeAuthConfig
  | GraphStaticAccessTokenAuthConfig;

export class GraphAuthError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "GraphAuthError";
  }
}

export function resolveGraphAuthConfig(config: NolendarConfig, env: NodeJS.ProcessEnv = process.env): GraphAuthConfig {
  const accessToken = env.MICROSOFT_ACCESS_TOKEN?.trim();

  if (accessToken) {
    return {
      mode: "static_access_token",
      tenantId: config.microsoft.tenant,
      clientId: undefined,
      scopes: parseScopes(env.MICROSOFT_GRAPH_SCOPES),
      accessToken,
    };
  }

  const clientId = env.MICROSOFT_CLIENT_ID;

  if (config.microsoft.authMode === "auth_code") {
    if (!clientId) {
      throw new GraphAuthError("Missing `MICROSOFT_CLIENT_ID` for Microsoft Graph authorization-code authentication.");
    }

    const clientSecret = env.MICROSOFT_CLIENT_SECRET;

    if (!clientSecret) {
      throw new GraphAuthError("Missing `MICROSOFT_CLIENT_SECRET` for Microsoft Graph authorization-code authentication.");
    }

    return {
      mode: "auth_code",
      tenantId: config.microsoft.tenant,
      clientId,
      clientSecret,
      redirectUri: env.MICROSOFT_REDIRECT_URI ?? "http://localhost:8787/auth/callback",
      scopes: parseScopes(env.MICROSOFT_GRAPH_SCOPES),
    };
  }

  return {
    mode: config.microsoft.authMode,
    tenantId: config.microsoft.tenant,
    clientId,
    scopes: parseScopes(env.MICROSOFT_GRAPH_SCOPES),
  };
}

function parseScopes(rawScopes?: string): string[] {
  if (!rawScopes) {
    return DEFAULT_GRAPH_SCOPES;
  }

  const scopes = rawScopes
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);

  if (scopes.length === 0) {
    throw new GraphAuthError("`MICROSOFT_GRAPH_SCOPES` must contain at least one scope if provided.");
  }

  return scopes;
}
