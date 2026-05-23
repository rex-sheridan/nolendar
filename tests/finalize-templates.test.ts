import { describe, expect, it, vi } from "vitest";

import type { NolendarConfig } from "../src/domain/config.js";
import type { Meeting } from "../src/domain/meeting.js";
import { finalizeTemplatedMeetingPages } from "../src/finalize-templates.js";

const CONFIG: NolendarConfig = {
  microsoft: {
    tenant: "common",
    authMode: "device_code",
  },
  notion: {
    databaseId: "data-source-id",
    dataSourceTemplate: {
      type: "default",
    },
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
  start: "2026-05-23T13:00:00.000Z",
  end: "2026-05-23T14:00:00.000Z",
  attendees: [],
  isCancelled: false,
  isRecurring: false,
};

describe("finalizeTemplatedMeetingPages", () => {
  it("finalizes unfinalized templated pages and marks existing generated content without appending duplicates", async () => {
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
      getDefaultAssigneeUserId: vi.fn(async () => undefined),
      getTemplateBlocks: vi.fn(async () => []),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi
        .fn()
        .mockResolvedValueOnce({ id: "page-1", eventId: "evt-1", changeKey: "ck-1" })
        .mockResolvedValueOnce({ id: "page-2", eventId: "evt-2", changeKey: "ck-2" })
        .mockResolvedValueOnce({ id: "page-3", eventId: "evt-3", changeKey: "ck-3" }),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi
        .fn()
        .mockResolvedValueOnce("appended")
        .mockResolvedValueOnce("marked_existing"),
    };

    const result = await finalizeTemplatedMeetingPages(
      CONFIG,
      [
        MEETING,
        { ...MEETING, id: "evt-2", changeKey: "ck-2", title: "Review" },
        { ...MEETING, id: "evt-3", changeKey: "ck-3", title: "Retro" },
      ],
      notion,
    );

    expect(result).toEqual({
      finalized: 2,
      markedExisting: 1,
      missingPage: 0,
      skippedCancelled: 0,
      filtered: 0,
    });
  });
});
