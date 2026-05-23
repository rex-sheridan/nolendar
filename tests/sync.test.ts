import { describe, expect, it, vi } from "vitest";

import type { NolendarConfig } from "../src/domain/config.js";
import type { Meeting } from "../src/domain/meeting.js";
import { syncMeetingsToNotion } from "../src/sync.js";

const CONFIG: NolendarConfig = {
  microsoft: {
    tenant: "common",
  },
  notion: {
    databaseId: "data-source-id",
  },
  calendars: [
    {
      id: "primary",
      name: "Primary",
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

const MEETING: Meeting = {
  id: "evt-1",
  changeKey: "ck-1",
  calendarId: "primary",
  calendarName: "Primary",
  title: "Planning",
  start: "2026-05-22T13:00:00.000Z",
  end: "2026-05-22T14:00:00.000Z",
  organizer: "Jordan",
  attendees: [{ email: "riley@example.com", optional: false }],
  meetingLink: "https://teams.example/join",
  eventLink: "https://outlook.example/event",
  agenda: "Discuss roadmap",
  isCancelled: false,
  isRecurring: false,
};

describe("syncMeetingsToNotion", () => {
  it("creates a page for an unseen meeting", async () => {
    const notion = {
      retrieveDataSource: vi.fn(async () => ({
        id: "data-source-id",
        title: "Meetings",
        properties: {
          Name: { id: "title", name: "Name", type: "title" },
          Due: { id: "due", name: "Due", type: "date" },
          "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
          "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
        },
      })),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => null),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
    };

    const result = await syncMeetingsToNotion(CONFIG, [MEETING], notion);

    expect(result).toEqual({
      created: 1,
      updated: 0,
      skipped: 0,
      filtered: 0,
      dryRun: false,
    });
    expect(notion.createMeetingPage).toHaveBeenCalledTimes(1);
    expect(notion.updateMeetingPage).not.toHaveBeenCalled();
  });

  it("skips a page when the changeKey matches", async () => {
    const notion = {
      retrieveDataSource: vi.fn(async () => ({
        id: "data-source-id",
        title: "Meetings",
        properties: {
          Name: { id: "title", name: "Name", type: "title" },
          Due: { id: "due", name: "Due", type: "date" },
          "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
          "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
        },
      })),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => ({ id: "page-1", eventId: "evt-1", changeKey: "ck-1" })),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
    };

    const result = await syncMeetingsToNotion(CONFIG, [MEETING], notion);

    expect(result.skipped).toBe(1);
    expect(notion.createMeetingPage).not.toHaveBeenCalled();
    expect(notion.updateMeetingPage).not.toHaveBeenCalled();
  });

  it("updates an existing page when the changeKey differs", async () => {
    const notion = {
      retrieveDataSource: vi.fn(async () => ({
        id: "data-source-id",
        title: "Meetings",
        properties: {
          Name: { id: "title", name: "Name", type: "title" },
          Due: { id: "due", name: "Due", type: "date" },
          "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
          "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
        },
      })),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => ({ id: "page-1", eventId: "evt-1", changeKey: "old" })),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
    };

    const result = await syncMeetingsToNotion(CONFIG, [MEETING], notion);

    expect(result.updated).toBe(1);
    expect(notion.updateMeetingPage).toHaveBeenCalledTimes(1);
    expect(notion.createMeetingPage).not.toHaveBeenCalled();
  });

  it("supports dry-run mode without mutating Notion", async () => {
    const notion = {
      retrieveDataSource: vi.fn(async () => ({
        id: "data-source-id",
        title: "Meetings",
        properties: {
          Name: { id: "title", name: "Name", type: "title" },
          Due: { id: "due", name: "Due", type: "date" },
          "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
          "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
        },
      })),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => null),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
    };

    const result = await syncMeetingsToNotion(CONFIG, [MEETING], notion, { dryRun: true });

    expect(result).toEqual({
      created: 1,
      updated: 0,
      skipped: 0,
      filtered: 0,
      dryRun: true,
    });
    expect(notion.createMeetingPage).not.toHaveBeenCalled();
    expect(notion.updateMeetingPage).not.toHaveBeenCalled();
  });
});
