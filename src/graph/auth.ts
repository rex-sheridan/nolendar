import type { NolendarConfig } from "../domain/config.js";

export const DEFAULT_GRAPH_SCOPES = ["Calendars.Read", "User.Read"];

export interface GraphAuthConfig {
  tenantId: string;
  clientId: string;
  scopes: string[];
}

export class GraphAuthError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "GraphAuthError";
  }
}

export function resolveGraphAuthConfig(config: NolendarConfig, env: NodeJS.ProcessEnv = process.env): GraphAuthConfig {
  const clientId = env.MICROSOFT_CLIENT_ID;

  if (!clientId) {
    throw new GraphAuthError("Missing `MICROSOFT_CLIENT_ID` for Microsoft Graph device-code authentication.");
  }

  return {
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
