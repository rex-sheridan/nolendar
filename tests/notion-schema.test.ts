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
    peopleDataSource: {
      databaseId: "people-data-source-id",
      nameProperty: "Name",
      emailProperty: "Email Address",
    },
    defaultTags: ["meeting"],
    defaultAssigneeEmail: "me@example.com",
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
    eventLink: "Source URL",
    tags: "Tags",
    assignee: "Assignee",
    participants: "Participants",
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
      { name: "Source URL", type: "url" },
      { name: "Tags", type: "multi_select" },
      { name: "Assignee", type: "people" },
      { name: "Participants", type: "relation" },
    ]);
  });

  it("requires a status property when canceled meetings are mapped to status", () => {
    expect(
      buildRequiredNotionProperties({
        ...CONFIG,
        notion: {
          ...CONFIG.notion,
          canceledMeetings: {
            action: "set_status",
            statusProperty: "Status",
            statusValue: "Canceled",
          },
        },
      }),
    ).toContainEqual({ name: "Status", type: "status" });
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
        "Source URL": { id: "source-url", name: "Source URL", type: "url" },
        Tags: { id: "tags", name: "Tags", type: "multi_select" },
        Assignee: { id: "assignee", name: "Assignee", type: "people" },
        Participants: { id: "participants", name: "Participants", type: "relation" },
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
        "Email Address": { id: "email", name: "Email Address", type: "rich_text" },
        Due: { id: "due", name: "Due", type: "date" },
        "Source URL": { id: "source-url", name: "Source URL", type: "rich_text" },
        Participants: { id: "participants", name: "Participants", type: "people" },
      },
    });

    expect(result.valid).toBe(false);
    expect(result.missing).toEqual([
      { name: "Outlook Event ID", type: "rich_text" },
      { name: "Outlook ChangeKey", type: "rich_text" },
      { name: "Tags", type: "multi_select" },
      { name: "Assignee", type: "people" },
    ]);
    expect(result.mismatched).toEqual([
      { name: "Name", expectedType: "title", actualType: "rich_text" },
      { name: "Source URL", expectedType: "url", actualType: "rich_text" },
      { name: "Participants", expectedType: "relation", actualType: "people" },
    ]);
  });
});
