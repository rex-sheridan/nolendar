import { describe, expect, it } from "vitest";

import type { NolendarConfig } from "../src/domain/config.js";
import { DEFAULT_GRAPH_SCOPES, GraphAuthError, resolveGraphAuthConfig } from "../src/graph/auth.js";

const CONFIG: NolendarConfig = {
  microsoft: {
    tenant: "organizations",
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
  it("requires a client id", () => {
    expect(() => resolveGraphAuthConfig(CONFIG, {})).toThrowError(
      new GraphAuthError("Missing `MICROSOFT_CLIENT_ID` for Microsoft Graph device-code authentication."),
    );
  });

  it("uses default scopes when none are provided", () => {
    expect(resolveGraphAuthConfig(CONFIG, { MICROSOFT_CLIENT_ID: "client-id" })).toEqual({
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
      tenantId: "organizations",
      clientId: "client-id",
      scopes: ["Calendars.Read", "User.Read"],
    });
  });
});
