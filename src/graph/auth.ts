import type { NolendarConfig } from "../domain/config.js";

export const DEFAULT_GRAPH_SCOPES = ["Calendars.Read", "User.Read"];

export interface GraphAuthBaseConfig {
  tenantId: string;
  clientId: string;
  scopes: string[];
}

export interface GraphDeviceCodeAuthConfig extends GraphAuthBaseConfig {
  mode: "device_code";
}

export interface GraphAuthorizationCodeAuthConfig extends GraphAuthBaseConfig {
  mode: "auth_code";
  clientSecret: string;
  redirectUri: string;
}

export type GraphAuthConfig = GraphDeviceCodeAuthConfig | GraphAuthorizationCodeAuthConfig;

export class GraphAuthError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "GraphAuthError";
  }
}

export function resolveGraphAuthConfig(config: NolendarConfig, env: NodeJS.ProcessEnv = process.env): GraphAuthConfig {
  const clientId = env.MICROSOFT_CLIENT_ID;

  if (!clientId) {
    throw new GraphAuthError(
      `Missing \`MICROSOFT_CLIENT_ID\` for Microsoft Graph ${formatMode(config.microsoft.authMode)} authentication.`,
    );
  }

  if (config.microsoft.authMode === "auth_code") {
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
    mode: "device_code",
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

function formatMode(mode: NolendarConfig["microsoft"]["authMode"]): string {
  return mode === "auth_code" ? "authorization-code" : "device-code";
}
