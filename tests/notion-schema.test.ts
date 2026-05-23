import { describe, expect, it } from "vitest";

import type { NolendarConfig } from "../src/domain/config.js";
import { buildRequiredNotionProperties, validateNotionSchema } from "../src/notion/schema.js";

const CONFIG: NolendarConfig = {
  microsoft: {
    tenant: "common",
    authMode: "device_code",
  },
  notion: {
    databaseId: "data-source-id",
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

describe("buildRequiredNotionProperties", () => {
  it("derives the required schema from config mapping", () => {
    expect(buildRequiredNotionProperties(CONFIG)).toEqual([
      { name: "Name", type: "title" },
      { name: "Due", type: "date" },
      { name: "Outlook Event ID", type: "rich_text" },
      { name: "Outlook ChangeKey", type: "rich_text" },
    ]);
  });
});

describe("validateNotionSchema", () => {
  it("accepts a compatible data source schema", () => {
    const result = validateNotionSchema(CONFIG, {
      id: "data-source-id",
      title: "Meetings",
      properties: {
        Name: { id: "title", name: "Name", type: "title" },
        Due: { id: "due", name: "Due", type: "date" },
        "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
        "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
      },
    });

    expect(result.valid).toBe(true);
    expect(result.missing).toEqual([]);
    expect(result.mismatched).toEqual([]);
  });

  it("reports missing and mismatched properties", () => {
    const result = validateNotionSchema(CONFIG, {
      id: "data-source-id",
      title: "Meetings",
      properties: {
        Name: { id: "title", name: "Name", type: "rich_text" },
        Due: { id: "due", name: "Due", type: "date" },
      },
    });

    expect(result.valid).toBe(false);
    expect(result.missing).toEqual([
      { name: "Outlook Event ID", type: "rich_text" },
      { name: "Outlook ChangeKey", type: "rich_text" },
    ]);
    expect(result.mismatched).toEqual([
      { name: "Name", expectedType: "title", actualType: "rich_text" },
    ]);
  });
});
