import { describe, expect, it, vi } from "vitest";

import type { CalendarConfig } from "../src/domain/config.js";
import { GraphMeetingSource, normalizeGraphEvent } from "../src/graph/graph-meeting-source.js";

const CALENDAR: CalendarConfig = {
  id: "primary",
  name: "Primary",
};

describe("normalizeGraphEvent", () => {
  it("maps a Graph event into the normalized meeting shape", () => {
    const meeting = normalizeGraphEvent(
      {
        id: "evt-1",
        changeKey: "ck-1",
        subject: "1:1",
        start: {
          dateTime: "2026-05-22T09:00:00-04:00",
        },
        end: {
          dateTime: "2026-05-22T09:30:00-04:00",
        },
        organizer: {
          emailAddress: {
            name: "Jordan",
          },
        },
        attendees: [
          {
            emailAddress: {
              name: "Riley",
              address: "riley@example.com",
            },
            type: "required",
          },
          {
            emailAddress: {
              address: "morgan@example.com",
            },
            type: "optional",
          },
        ],
        onlineMeeting: {
          joinUrl: "https://teams.example/join",
        },
        webLink: "https://outlook.example/event",
        bodyPreview: "Discuss roadmap",
        responseStatus: {
          response: "accepted",
        },
        type: "occurrence",
      },
      CALENDAR,
    );

    expect(meeting).toEqual({
      id: "evt-1",
      changeKey: "ck-1",
      calendarId: "primary",
      calendarName: "Primary",
      title: "1:1",
      start: "2026-05-22T13:00:00.000Z",
      end: "2026-05-22T13:30:00.000Z",
      organizer: "Jordan",
      attendees: [
        {
          name: "Riley",
          email: "riley@example.com",
          optional: false,
        },
        {
          name: undefined,
          email: "morgan@example.com",
          optional: true,
        },
      ],
      meetingLink: "https://teams.example/join",
      eventLink: "https://outlook.example/event",
      agenda: "Discuss roadmap",
      responseStatus: "accepted",
      isCancelled: false,
      isRecurring: true,
    });
  });
});

describe("GraphMeetingSource", () => {
  it("fetches calendar view events for the requested window", async () => {
    const fetchMock = vi.fn(async () =>
      new Response(
        JSON.stringify({
          value: [
            {
              id: "evt-1",
              changeKey: "ck-1",
              subject: "Planning",
              start: {
                dateTime: "2026-05-22T10:00:00.000Z",
              },
              end: {
                dateTime: "2026-05-22T11:00:00.000Z",
              },
            },
          ],
        }),
        {
          status: 200,
          headers: {
            "Content-Type": "application/json",
          },
        },
      ),
    );
    const source = new GraphMeetingSource(
      {
        getAccessToken: async () => "token-123",
      },
      fetchMock as unknown as typeof fetch,
    );

    const meetings = await source.listMeetings({
      calendar: CALENDAR,
      window: {
        start: "2026-05-22T00:00:00.000Z",
        end: "2026-05-23T00:00:00.000Z",
      },
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    const firstCall = fetchMock.mock.calls[0];

    expect(firstCall).toBeDefined();
    const [requestUrl, init] = firstCall as unknown as [URL, RequestInit];
    expect(String(requestUrl)).toContain("/me/calendars/primary/calendarView");
    expect(String(requestUrl)).toContain("startDateTime=2026-05-22T00%3A00%3A00.000Z");
    expect(String(requestUrl)).toContain("endDateTime=2026-05-23T00%3A00%3A00.000Z");
    expect(init).toEqual({
      headers: {
        Authorization: "Bearer token-123",
        Accept: "application/json",
      },
    });
    expect(meetings).toHaveLength(1);
    expect(meetings[0]?.title).toBe("Planning");
  });
});
