import path from "node:path";
import { describe, expect, it } from "vitest";

import { ConfigError, normalizeConfig } from "../src/config.js";

describe("normalizeConfig", () => {
  it("applies defaults for optional sections", () => {
    const config = normalizeConfig(
      {
        notion: {
          databaseId: "db_123",
        },
        calendars: [
          {
            id: "primary",
          },
        ],
      },
      "/tmp/nolendar.yml",
    );

    expect(config.microsoft.tenant).toBe("common");
    expect(config.sync.lookahead).toBe("today");
    expect(config.sync.statePath).toBe(path.resolve("/tmp", ".nolendar/state.json"));
    expect(config.mapping.eventId).toBe("Outlook Event ID");
    expect(config.filters.ignoreDeclined).toBe(true);
  });

  it("resolves configured state path relative to config file", () => {
    const config = normalizeConfig(
      {
        notion: {
          databaseId: "db_123",
        },
        calendars: [
          {
            id: "team",
          },
        ],
        sync: {
          statePath: "./data/state.json",
        },
      },
      "/Users/rex/workspace/nolendar/config/nolendar.yml",
    );

    expect(config.sync.statePath).toBe("/Users/rex/workspace/nolendar/config/data/state.json");
  });

  it("rejects an invalid lookahead window", () => {
    expect(() =>
      normalizeConfig({
        notion: {
          databaseId: "db_123",
        },
        calendars: [
          {
            id: "team",
          },
        ],
        sync: {
          lookahead: "30d",
        },
      }),
    ).toThrowError(new ConfigError("`sync.lookahead` must be one of: today, 24h, 7d."));
  });

  it("rejects missing notion database id", () => {
    expect(() =>
      normalizeConfig({
        notion: {},
        calendars: [
          {
            id: "team",
          },
        ],
      }),
    ).toThrowError(new ConfigError("`notion.databaseId` is required."));
  });
});
