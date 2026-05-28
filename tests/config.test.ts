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
    expect(config.microsoft.authMode).toBe("device_code");
    expect(config.sync.lookahead).toBe("today");
    expect(config.sync.statePath).toBe(path.resolve("/tmp", ".nolendar/state.json"));
    expect(config.mapping.eventId).toBe("Outlook Event ID");
    expect(config.filters.ignoreDeclined).toBe(true);
  });

  it("accepts a default assignee email", () => {
    const config = normalizeConfig({
      notion: {
        databaseId: "db_123",
        defaultAssigneeEmail: "me@example.com",
      },
      calendars: [
        {
          id: "primary",
        },
      ],
    });

    expect(config.notion.defaultAssigneeEmail).toBe("me@example.com");
  });

  it("accepts a people data source and participants mapping", () => {
    const config = normalizeConfig({
      notion: {
        databaseId: "db_123",
        peopleDataSource: {
          databaseId: "people_123",
        },
      },
      calendars: [
        {
          id: "primary",
        },
      ],
      mapping: {
        participants: "Participants",
      },
    });

    expect(config.notion.peopleDataSource).toEqual({
      databaseId: "people_123",
      nameProperty: "Name",
      emailProperty: "Email Address",
      maxAttendeesPerMeeting: 10,
    });
    expect(config.mapping.participants).toBe("Participants");
  });

  it("accepts a configured participant attendee limit", () => {
    const config = normalizeConfig({
      notion: {
        databaseId: "db_123",
        peopleDataSource: {
          databaseId: "people_123",
          maxAttendeesPerMeeting: 3,
        },
      },
      calendars: [
        {
          id: "primary",
        },
      ],
      mapping: {
        participants: "Participants",
      },
    });

    expect(config.notion.peopleDataSource?.maxAttendeesPerMeeting).toBe(3);
  });

  it("accepts a true data source template configuration", () => {
    const config = normalizeConfig({
      notion: {
        databaseId: "db_123",
        dataSourceTemplate: {
          type: "template_id",
          templateId: "template-123",
          timezone: "America/New_York",
        },
      },
      calendars: [
        {
          id: "primary",
        },
      ],
    });

    expect(config.notion.dataSourceTemplate).toEqual({
      type: "template_id",
      templateId: "template-123",
      timezone: "America/New_York",
    });
  });

  it("accepts configured generated page content sections", () => {
    const config = normalizeConfig({
      notion: {
        databaseId: "db_123",
        pageContent: {
          sections: ["meeting_link", "calendar_event", "meeting_details"],
        },
      },
      calendars: [
        {
          id: "primary",
        },
      ],
    });

    expect(config.notion.pageContent).toEqual({
      sections: ["meeting_link", "calendar_event", "meeting_details"],
    });
  });

  it("accepts an emoji page icon", () => {
    const config = normalizeConfig({
      notion: {
        databaseId: "db_123",
        pageIcon: {
          type: "emoji",
          emoji: "📝",
        },
      },
      calendars: [
        {
          id: "primary",
        },
      ],
    });

    expect(config.notion.pageIcon).toEqual({
      type: "emoji",
      emoji: "📝",
    });
  });

  it("accepts a native notion icon with color", () => {
    const config = normalizeConfig({
      notion: {
        databaseId: "db_123",
        pageIcon: {
          type: "icon",
          name: "calendar",
          color: "blue",
        },
      },
      calendars: [
        {
          id: "primary",
        },
      ],
    });

    expect(config.notion.pageIcon).toEqual({
      type: "icon",
      name: "calendar",
      color: "blue",
    });
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

  it("accepts relative lookahead windows", () => {
    const config = normalizeConfig({
      notion: {
        databaseId: "db_123",
      },
      calendars: [
        {
          id: "team",
        },
      ],
      sync: {
        lookahead: "5d",
      },
    });

    expect(config.sync.lookahead).toBe("5d");
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
          lookahead: "30x",
        },
      }),
    ).toThrowError(
      new ConfigError("`sync.lookahead` must be `today` or a relative range like `12h`, `5d`, `2w`, or `3m`."),
    );
  });

  it("rejects invalid participant attendee limits", () => {
    expect(() =>
      normalizeConfig({
        notion: {
          databaseId: "db_123",
          peopleDataSource: {
            databaseId: "people_123",
            maxAttendeesPerMeeting: -1,
          },
        },
        calendars: [
          {
            id: "team",
          },
        ],
        mapping: {
          participants: "Participants",
        },
      }),
    ).toThrowError(
      new ConfigError("`notion.peopleDataSource.maxAttendeesPerMeeting` must be a non-negative integer."),
    );
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

  it("rejects a participants mapping without a people data source", () => {
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
        mapping: {
          participants: "Participants",
        },
      }),
    ).toThrowError(new ConfigError("`notion.peopleDataSource` is required when `mapping.participants` is configured."));
  });

  it("rejects a people data source without a participants mapping", () => {
    expect(() =>
      normalizeConfig({
        notion: {
          databaseId: "db_123",
          peopleDataSource: {
            databaseId: "people_123",
          },
        },
        calendars: [
          {
            id: "team",
          },
        ],
      }),
    ).toThrowError(new ConfigError("`mapping.participants` is required when `notion.peopleDataSource` is configured."));
  });

  it("rejects a template_id data source template without templateId", () => {
    expect(() =>
      normalizeConfig({
        notion: {
          databaseId: "db_123",
          dataSourceTemplate: {
            type: "template_id",
          },
        },
        calendars: [
          {
            id: "team",
          },
        ],
      }),
    ).toThrowError(
      new ConfigError("`notion.dataSourceTemplate.templateId` is required when template type is `template_id`."),
    );
  });

  it("rejects simultaneous page-copy and true template configuration", () => {
    expect(() =>
      normalizeConfig({
        notion: {
          databaseId: "db_123",
          templatePageId: "page-123",
          dataSourceTemplate: {
            type: "default",
          },
        },
        calendars: [
          {
            id: "team",
          },
        ],
      }),
    ).toThrowError(new ConfigError("`notion.templatePageId` and `notion.dataSourceTemplate` are mutually exclusive."));
  });

  it("rejects an invalid generated page content section", () => {
    expect(() =>
      normalizeConfig({
        notion: {
          databaseId: "db_123",
          pageContent: {
            sections: ["notes", "bogus"],
          },
        },
        calendars: [
          {
            id: "team",
          },
        ],
      }),
    ).toThrowError(
      new ConfigError(
        "`notion.pageContent.sections[1]` must be one of: meeting_link, calendar_event, meeting_details, notes, action_items.",
      ),
    );
  });

  it("accepts the authorization code auth mode", () => {
    const config = normalizeConfig({
      microsoft: {
        tenant: "organizations",
        authMode: "auth_code",
      },
      notion: {
        databaseId: "db_123",
      },
      calendars: [
        {
          id: "team",
        },
      ],
    });

    expect(config.microsoft.authMode).toBe("auth_code");
  });

  it("accepts the interactive browser auth mode", () => {
    const config = normalizeConfig({
      microsoft: {
        tenant: "organizations",
        authMode: "interactive_browser",
      },
      notion: {
        databaseId: "db_123",
      },
      calendars: [
        {
          id: "team",
        },
      ],
    });

    expect(config.microsoft.authMode).toBe("interactive_browser");
  });

  it("rejects an invalid microsoft auth mode", () => {
    expect(() =>
      normalizeConfig({
        microsoft: {
          authMode: "magic",
        },
        notion: {
          databaseId: "db_123",
        },
        calendars: [
          {
            id: "team",
          },
        ],
      }),
    ).toThrowError(new ConfigError("`microsoft.authMode` must be one of: device_code, interactive_browser, auth_code."));
  });
});
