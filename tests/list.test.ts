import { describe, expect, it } from "vitest";

import type { NolendarConfig } from "../src/domain/config.js";
import type { Meeting } from "../src/domain/meeting.js";
import { listMeetings, resolveWindow } from "../src/list.js";

const BASE_CONFIG: NolendarConfig = {
  microsoft: {
    tenant: "common",
    authMode: "device_code",
  },
  notion: {
    databaseId: "db_123",
  },
  calendars: [
    {
      id: "primary",
      name: "Primary",
    },
    {
      id: "team",
      name: "Team",
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

describe("resolveWindow", () => {
  const clock = {
    now: () => new Date("2026-05-22T15:30:00.000Z"),
  };

  it("uses the full current UTC day for today", () => {
    expect(resolveWindow("today", clock)).toEqual({
      start: "2026-05-22T00:00:00.000Z",
      end: "2026-05-23T00:00:00.000Z",
    });
  });

  it("uses a rolling 24 hour window", () => {
    expect(resolveWindow("24h", clock)).toEqual({
      start: "2026-05-22T15:30:00.000Z",
      end: "2026-05-23T15:30:00.000Z",
    });
  });

  it("uses a rolling 7 day window", () => {
    expect(resolveWindow("7d", clock)).toEqual({
      start: "2026-05-22T15:30:00.000Z",
      end: "2026-05-29T15:30:00.000Z",
    });
  });
});

describe("listMeetings", () => {
  it("aggregates and sorts meetings from multiple calendars", async () => {
    const calls: string[] = [];

    const result = await listMeetings(BASE_CONFIG, "today", {
      clock: {
        now: () => new Date("2026-05-22T09:00:00.000Z"),
      },
      meetingSource: {
        async listMeetings({ calendar, window }) {
          calls.push(`${calendar.id}:${window.start}:${window.end}`);

          const fixtures: Record<string, Meeting[]> = {
            primary: [
              {
                id: "2",
                changeKey: "b",
                calendarId: "primary",
                calendarName: "Primary",
                title: "Standup",
                start: "2026-05-22T13:00:00.000Z",
                end: "2026-05-22T13:15:00.000Z",
                attendees: [],
                isCancelled: false,
                isRecurring: false,
              },
            ],
            team: [
              {
                id: "1",
                changeKey: "a",
                calendarId: "team",
                calendarName: "Team",
                title: "Planning",
                start: "2026-05-22T12:00:00.000Z",
                end: "2026-05-22T13:00:00.000Z",
                attendees: [],
                isCancelled: false,
                isRecurring: false,
              },
            ],
          };

          return fixtures[calendar.id] ?? [];
        },
      },
    });

    expect(calls).toEqual([
      "primary:2026-05-22T00:00:00.000Z:2026-05-23T00:00:00.000Z",
      "team:2026-05-22T00:00:00.000Z:2026-05-23T00:00:00.000Z",
    ]);
    expect(result.meetings.map((meeting) => meeting.title)).toEqual(["Planning", "Standup"]);
  });
});
