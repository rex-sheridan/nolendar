import { describe, expect, it } from "vitest";

import type { NolendarConfig } from "../src/domain/config.js";
import { listMeetingContentsForDay, resolveDayWindow } from "../src/meeting-contents.js";

const CONFIG: NolendarConfig = {
  microsoft: {
    tenant: "common",
    authMode: "device_code",
  },
  notion: {
    databaseId: "meetings",
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

describe("resolveDayWindow", () => {
  const clock = {
    now: () => new Date("2026-05-22T15:30:00.000Z"),
  };

  it("defaults relative days to full UTC day windows", () => {
    expect(resolveDayWindow("today", clock)).toEqual({
      label: "2026-05-22",
      start: "2026-05-22T00:00:00.000Z",
      end: "2026-05-23T00:00:00.000Z",
    });
    expect(resolveDayWindow("tomorrow", clock)).toEqual({
      label: "2026-05-23",
      start: "2026-05-23T00:00:00.000Z",
      end: "2026-05-24T00:00:00.000Z",
    });
  });

  it("accepts ISO dates and signed day offsets", () => {
    expect(resolveDayWindow("2026-06-01", clock).start).toBe("2026-06-01T00:00:00.000Z");
    expect(resolveDayWindow("+2d", clock).start).toBe("2026-05-24T00:00:00.000Z");
    expect(resolveDayWindow("-1d", clock).start).toBe("2026-05-21T00:00:00.000Z");
  });
});

describe("listMeetingContentsForDay", () => {
  it("prints compact Outlook meeting bodies", async () => {
    const lines = await listMeetingContentsForDay(
      CONFIG,
      {
        source: "outlook",
        day: "2026-05-22",
        detail: "compact",
      },
      {
        meetingSource: {
          async listMeetings({ window }) {
            expect(window).toEqual({
              label: "2026-05-22",
              start: "2026-05-22T00:00:00.000Z",
              end: "2026-05-23T00:00:00.000Z",
            });

            return [
              {
                id: "event-1",
                changeKey: "ck-1",
                calendarId: "primary",
                calendarName: "Primary",
                title: "Planning",
                start: "2026-05-22T13:00:00.000Z",
                end: "2026-05-22T13:30:00.000Z",
                details: "Discuss launch readiness.",
                attendees: [],
                isCancelled: false,
                isRecurring: false,
              },
            ];
          },
        },
        clock: {
          now: () => new Date("2026-05-22T15:30:00.000Z"),
        },
      },
    );

    expect(lines).toContain("# Meeting contents (outlook)");
    expect(lines).toContain("Title: Planning");
    expect(lines).toContain("  Discuss launch readiness.");
    expect(lines).not.toContain("Properties:");
  });

  it("prints full normalized Notion properties when requested", async () => {
    const lines = await listMeetingContentsForDay(
      CONFIG,
      {
        source: "notion",
        day: "today",
        detail: "full",
      },
      {
        notion: {
          async listMeetingPagesForWindow(args: {
            dataSourceId: string;
            datePropertyName: string;
            start: string;
            end: string;
          }) {
            expect(args).toEqual({
              dataSourceId: "meetings",
              datePropertyName: "Due",
              start: "2026-05-22T00:00:00.000Z",
              end: "2026-05-23T00:00:00.000Z",
            });

            return [
              {
                id: "page-1",
                url: "https://notion.so/page-1",
                properties: {
                  Name: "Planning",
                  Due: {
                    start: "2026-05-22T13:00:00.000Z",
                    end: "2026-05-22T13:30:00.000Z",
                  },
                },
                body: "Discuss launch readiness.",
              },
            ];
          },
        } as never,
        clock: {
          now: () => new Date("2026-05-22T15:30:00.000Z"),
        },
      },
    );

    expect(lines).toContain("# Meeting contents (notion)");
    expect(lines).toContain("Title: Planning");
    expect(lines).toContain("Properties:");
    expect(lines.join("\n")).toContain('"url": "https://notion.so/page-1"');
  });
});
