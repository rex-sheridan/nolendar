import { describe, expect, it } from "vitest";

import type { FiltersConfig } from "../src/domain/config.js";
import type { Meeting } from "../src/domain/meeting.js";
import { shouldSyncMeeting } from "../src/filters.js";

const BASE_FILTERS: FiltersConfig = {
  ignoreDeclined: true,
  requireAttendees: false,
  ignorePersonal: false,
  ignoreOptionalAttendance: false,
};

const BASE_MEETING: Meeting = {
  id: "evt-1",
  changeKey: "ck-1",
  calendarId: "primary",
  title: "Planning",
  start: "2026-05-22T13:00:00.000Z",
  end: "2026-05-22T14:00:00.000Z",
  attendees: [],
  isCancelled: false,
  isRecurring: false,
};

describe("shouldSyncMeeting", () => {
  it("filters declined meetings when configured", () => {
    expect(
      shouldSyncMeeting(
        {
          ...BASE_MEETING,
          responseStatus: "declined",
        },
        BASE_FILTERS,
      ),
    ).toBe(false);
  });

  it("filters meetings shorter than the configured minimum", () => {
    expect(
      shouldSyncMeeting(BASE_MEETING, {
        ...BASE_FILTERS,
        minDurationMinutes: 90,
      }),
    ).toBe(false);
  });

  it("filters meetings without attendees when required", () => {
    expect(
      shouldSyncMeeting(BASE_MEETING, {
        ...BASE_FILTERS,
        requireAttendees: true,
      }),
    ).toBe(false);
  });

  it("keeps eligible meetings", () => {
    expect(
      shouldSyncMeeting(
        {
          ...BASE_MEETING,
          attendees: [{ email: "a@example.com", optional: false }],
        },
        {
          ...BASE_FILTERS,
          requireAttendees: true,
          minDurationMinutes: 30,
        },
      ),
    ).toBe(true);
  });
});
