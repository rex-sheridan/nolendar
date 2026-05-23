import { describe, expect, it } from "vitest";

import { formatMeeting } from "../src/meeting-format.js";

describe("formatMeeting", () => {
  it("formats a normalized meeting into printable lines", () => {
    const lines = formatMeeting({
      id: "1",
      changeKey: "abc",
      calendarId: "team",
      calendarName: "Team",
      title: "Roadmap Review",
      start: "2026-05-22T12:00:00.000Z",
      end: "2026-05-22T13:00:00.000Z",
      organizer: "Casey",
      attendees: [
        {
          name: "Alex",
          optional: false,
        },
        {
          email: "sam@example.com",
          optional: true,
        },
      ],
      meetingLink: "https://teams.example/join",
      eventLink: "https://outlook.example/event",
      responseStatus: "accepted",
      isCancelled: false,
      isRecurring: true,
    });

    expect(lines).toEqual([
      "2026-05-22T12:00:00.000Z | Roadmap Review",
      "  Calendar: Team",
      "  Organizer: Casey",
      "  Attendees: Alex, sam@example.com (optional)",
      "  Meeting Link: https://teams.example/join",
      "  Event Link: https://outlook.example/event",
      "  Flags: recurring, accepted",
    ]);
  });
});
