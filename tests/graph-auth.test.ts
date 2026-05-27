import { describe, expect, it } from "vitest";

import type { NolendarConfig } from "../src/domain/config.js";
import { DEFAULT_GRAPH_SCOPES, GraphAuthError, resolveGraphAuthConfig } from "../src/graph/auth.js";

const CONFIG: NolendarConfig = {
  microsoft: {
    tenant: "organizations",
    authMode: "device_code",
  },
  notion: {
    databaseId: "db_123",
  },
  calendars: [
    {
      id: "primary",
    },
  ],
  filters: {
    ignoreDeclined: true,
    requireAttendees: false,
    ignorePersonal: false,
    ignoreOptionalAttendance: false,
  },
  mapping: {
    title: "Name",
    due: "Due",
    eventId: "Outlook Event ID",
    changeKey: "Outlook ChangeKey",
  },
  sync: {
    lookahead: "today",
    statePath: "/tmp/.nolendar/state.json",
  },
};

describe("resolveGraphAuthConfig", () => {
  it("requires a client id for device code mode", () => {
    expect(() => resolveGraphAuthConfig(CONFIG, {})).toThrowError(
      new GraphAuthError("Missing `MICROSOFT_CLIENT_ID` for Microsoft Graph device_code authentication."),
    );
  });

  it("uses default scopes when none are provided", () => {
    expect(resolveGraphAuthConfig(CONFIG, { MICROSOFT_CLIENT_ID: "client-id" })).toEqual({
      mode: "device_code",
      tenantId: "organizations",
      clientId: "client-id",
      scopes: DEFAULT_GRAPH_SCOPES,
    });
  });

  it("parses comma-separated scopes from the environment", () => {
    expect(
      resolveGraphAuthConfig(CONFIG, {
        MICROSOFT_CLIENT_ID: "client-id",
        MICROSOFT_GRAPH_SCOPES: "Calendars.Read, User.Read",
      }),
    ).toEqual({
      mode: "device_code",
      tenantId: "organizations",
      clientId: "client-id",
      scopes: ["Calendars.Read", "User.Read"],
    });
  });

  it("requires a client secret for auth code mode", () => {
    expect(() =>
      resolveGraphAuthConfig(
        {
          ...CONFIG,
          microsoft: {
            tenant: "organizations",
            authMode: "auth_code",
          },
        },
        {
        },
      ),
    ).toThrowError(
      new GraphAuthError("Missing `MICROSOFT_CLIENT_ID` for Microsoft Graph authorization-code authentication."),
    );
  });

  it("requires a client secret once auth code has a client id", () => {
    expect(() =>
      resolveGraphAuthConfig(
        {
          ...CONFIG,
          microsoft: {
            tenant: "organizations",
            authMode: "auth_code",
          },
        },
        {
          MICROSOFT_CLIENT_ID: "client-id",
        },
      ),
    ).toThrowError(
      new GraphAuthError("Missing `MICROSOFT_CLIENT_SECRET` for Microsoft Graph authorization-code authentication."),
    );
  });

  it("builds auth code settings when required environment is present", () => {
    expect(
      resolveGraphAuthConfig(
        {
          ...CONFIG,
          microsoft: {
            tenant: "organizations",
            authMode: "auth_code",
          },
        },
        {
          MICROSOFT_CLIENT_ID: "client-id",
          MICROSOFT_CLIENT_SECRET: "client-secret",
          MICROSOFT_REDIRECT_URI: "http://localhost:8787/auth/callback",
        },
      ),
    ).toEqual({
      mode: "auth_code",
      tenantId: "organizations",
      clientId: "client-id",
      clientSecret: "client-secret",
      redirectUri: "http://localhost:8787/auth/callback",
      scopes: DEFAULT_GRAPH_SCOPES,
    });
  });

  it("requires a client id for interactive browser mode", () => {
    expect(() =>
      resolveGraphAuthConfig(
        {
          ...CONFIG,
          microsoft: {
            tenant: "organizations",
            authMode: "interactive_browser",
          },
        },
        {},
      ),
    ).toThrowError(
      new GraphAuthError("Missing `MICROSOFT_CLIENT_ID` for Microsoft Graph interactive_browser authentication."),
    );
  });

  it("prefers a raw access token when provided", () => {
    expect(
      resolveGraphAuthConfig(CONFIG, {
        MICROSOFT_ACCESS_TOKEN: "bearer-token",
      }),
    ).toEqual({
      mode: "static_access_token",
      tenantId: "organizations",
      clientId: undefined,
      scopes: DEFAULT_GRAPH_SCOPES,
      accessToken: "bearer-token",
    });
  });
});
