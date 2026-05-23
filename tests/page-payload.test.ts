import { describe, expect, it } from "vitest";

import type { NolendarConfig } from "../src/domain/config.js";
import type { Meeting } from "../src/domain/meeting.js";
import { buildMeetingProperties } from "../src/notion/page-payload.js";

const CONFIG: NolendarConfig = {
  microsoft: {
    tenant: "common",
    authMode: "device_code",
  },
  notion: {
    databaseId: "data-source-id",
    defaultTags: ["meeting", "outlook"],
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
  },
  sync: {
    lookahead: "today",
    statePath: "/tmp/.nolendar/state.json",
  },
};

const MEETING: Meeting = {
  id: "evt-1",
  changeKey: "ck-1",
  calendarId: "primary",
  title: "Planning",
  start: "2026-05-22T13:00:00.000Z",
  end: "2026-05-22T14:00:00.000Z",
  eventLink: "https://outlook.example/event",
  attendees: [],
  isCancelled: false,
  isRecurring: false,
};

describe("buildMeetingProperties", () => {
  it("maps event webLink into a configured url property and applies configured tags", () => {
    const properties = buildMeetingProperties(
      CONFIG,
      {
        id: "data-source-id",
        properties: {
          Name: { id: "title", name: "Name", type: "title" },
          Due: { id: "due", name: "Due", type: "date" },
          "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
          "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
          "Source URL": { id: "source-url", name: "Source URL", type: "url" },
          Tags: { id: "tags", name: "Tags", type: "multi_select" },
        },
      },
      MEETING,
    );

    expect(properties["Source URL"]).toEqual({
      url: "https://outlook.example/event",
    });
    expect(properties.Tags).toEqual({
      multi_select: [{ name: "meeting" }, { name: "outlook" }],
    });
  });
});
