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
          timeZone: "Eastern Standard Time",
        },
        end: {
          dateTime: "2026-05-22T09:30:00-04:00",
          timeZone: "Eastern Standard Time",
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
        body: {
          contentType: "html",
          content:
            "<html><body><p>Discuss roadmap</p><p><a href=\"https://teams.microsoft.com/l/meetup-join/123\">Join Teams</a></p></body></html>",
        },
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
      details: "Discuss roadmap\nJoin Teams",
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
                dateTime: "2026-05-22T10:00:00.0000000",
                timeZone: "UTC",
              },
              end: {
                dateTime: "2026-05-22T11:00:00.0000000",
                timeZone: "UTC",
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
        Prefer: 'outlook.timezone="UTC"',
      },
    });
    expect(meetings).toHaveLength(1);
    expect(meetings[0]?.title).toBe("Planning");
    expect(meetings[0]?.start).toBe("2026-05-22T10:00:00.000Z");
  });

  it("ignores series master events from calendar view results", async () => {
    const fetchMock = vi.fn(async () =>
      new Response(
        JSON.stringify({
          value: [
            {
              id: "series-master",
              changeKey: "master-1",
              subject: "Weekly Sync",
              type: "seriesMaster",
              recurrence: {},
              start: {
                dateTime: "2026-05-22T10:00:00.0000000",
                timeZone: "UTC",
              },
              end: {
                dateTime: "2026-05-22T11:00:00.0000000",
                timeZone: "UTC",
              },
            },
            {
              id: "occ-1",
              changeKey: "occ-1",
              subject: "Weekly Sync",
              type: "occurrence",
              start: {
                dateTime: "2026-05-22T10:00:00.0000000",
                timeZone: "UTC",
              },
              end: {
                dateTime: "2026-05-22T11:00:00.0000000",
                timeZone: "UTC",
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

    expect(meetings.map((meeting) => meeting.id)).toEqual(["occ-1"]);
  });

  it("extracts a Teams link from the event body when onlineMeeting is unavailable", () => {
    const meeting = normalizeGraphEvent(
      {
        id: "evt-2",
        changeKey: "ck-2",
        subject: "Review",
        start: {
          dateTime: "2026-05-22T10:00:00.0000000",
          timeZone: "UTC",
        },
        end: {
          dateTime: "2026-05-22T11:00:00.0000000",
          timeZone: "UTC",
        },
        body: {
          contentType: "html",
          content:
            "<p>Join here: <a href=\"https://teams.microsoft.com/l/meetup-join/abc\">Microsoft Teams Meeting</a></p>",
        },
      },
      CALENDAR,
    );

    expect(meeting.meetingLink).toBe("https://teams.microsoft.com/l/meetup-join/abc");
    expect(meeting.details).toBe("Join here: Microsoft Teams Meeting");
  });

  it("treats bare UTC Graph timestamps as UTC instead of local time", () => {
    const meeting = normalizeGraphEvent(
      {
        id: "evt-3",
        changeKey: "ck-3",
        subject: "Timezone check",
        start: {
          dateTime: "2026-05-23T15:00:00.0000000",
          timeZone: "UTC",
        },
        end: {
          dateTime: "2026-05-23T16:00:00.0000000",
          timeZone: "UTC",
        },
      },
      CALENDAR,
    );

    expect(meeting.start).toBe("2026-05-23T15:00:00.000Z");
    expect(meeting.end).toBe("2026-05-23T16:00:00.000Z");
  });

  it("follows paginated calendar view delta responses and returns the final delta link", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            value: [
              {
                id: "evt-1",
                changeKey: "ck-1",
                subject: "Planning",
                start: {
                  dateTime: "2026-05-22T10:00:00.0000000",
                  timeZone: "UTC",
                },
                end: {
                  dateTime: "2026-05-22T11:00:00.0000000",
                  timeZone: "UTC",
                },
              },
            ],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/calendars/primary/calendarView/delta?$skiptoken=page-2",
          }),
          {
            status: 200,
            headers: {
              "Content-Type": "application/json",
            },
          },
        ),
      )
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            value: [
              {
                id: "evt-2",
                changeKey: "ck-2",
                subject: "Review",
                start: {
                  dateTime: "2026-05-22T12:00:00.0000000",
                  timeZone: "UTC",
                },
                end: {
                  dateTime: "2026-05-22T13:00:00.0000000",
                  timeZone: "UTC",
                },
              },
            ],
            "@odata.deltaLink": "https://graph.microsoft.com/v1.0/me/calendars/primary/calendarView/delta?$deltatoken=done",
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

    const result = await source.listMeetingChanges({
      calendar: CALENDAR,
      window: {
        start: "2026-05-22T00:00:00.000Z",
        end: "2026-05-23T00:00:00.000Z",
      },
    });

    expect(fetchMock).toHaveBeenCalledTimes(2);
    const [firstRequestUrl, firstRequestInit] = fetchMock.mock.calls[0] as unknown as [URL, RequestInit];
    expect(String(firstRequestUrl)).toContain("/calendarView/delta");
    expect(String(firstRequestUrl)).not.toContain("%24orderby");
    expect(String(firstRequestUrl)).not.toContain("%24top");
    expect(firstRequestInit).toEqual({
      headers: {
        Authorization: "Bearer token-123",
        Accept: "application/json",
        Prefer: 'outlook.timezone="UTC", odata.maxpagesize=100',
      },
    });
    expect(result.meetings.map((meeting) => meeting.id)).toEqual(["evt-1", "evt-2"]);
    expect(result.deltaLink).toBe(
      "https://graph.microsoft.com/v1.0/me/calendars/primary/calendarView/delta?$deltatoken=done",
    );
  });

  it("fails fast when a delta response includes removed events", async () => {
    const fetchMock = vi.fn(async () =>
      new Response(
        JSON.stringify({
          value: [
            {
              id: "evt-removed",
              "@removed": {
                reason: "deleted",
              },
            },
          ],
          "@odata.deltaLink": "https://graph.microsoft.com/v1.0/me/calendars/primary/calendarView/delta?$deltatoken=done",
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

    await expect(
      source.listMeetingChanges({
        calendar: CALENDAR,
        window: {
          start: "2026-05-22T00:00:00.000Z",
          end: "2026-05-23T00:00:00.000Z",
        },
      }),
    ).rejects.toThrowError("Graph delta query returned removed events, which Nolendar does not handle yet.");
  });

  it("hydrates partial delta events before normalizing them", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            value: [
              {
                id: "evt-4",
                subject: "Hydrate me",
              },
            ],
            "@odata.deltaLink": "https://graph.microsoft.com/v1.0/me/calendars/primary/calendarView/delta?$deltatoken=done",
          }),
          {
            status: 200,
            headers: {
              "Content-Type": "application/json",
            },
          },
        ),
      )
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            id: "evt-4",
            changeKey: "ck-4",
            subject: "Hydrate me",
            start: {
              dateTime: "2026-05-22T14:00:00.0000000",
              timeZone: "UTC",
            },
            end: {
              dateTime: "2026-05-22T15:00:00.0000000",
              timeZone: "UTC",
            },
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

    const result = await source.listMeetingChanges({
      calendar: CALENDAR,
      window: {
        start: "2026-05-22T00:00:00.000Z",
        end: "2026-05-23T00:00:00.000Z",
      },
    });

    expect(fetchMock).toHaveBeenCalledTimes(2);
    const [hydrationUrl] = fetchMock.mock.calls[1] as unknown as [URL, RequestInit];
    expect(String(hydrationUrl)).toContain("/me/calendars/primary/events/evt-4");
    expect(result.meetings[0]?.changeKey).toBe("ck-4");
    expect(result.meetings[0]?.start).toBe("2026-05-22T14:00:00.000Z");
  });

  it("ignores series master events from delta results", async () => {
    const fetchMock = vi.fn(async () =>
      new Response(
        JSON.stringify({
          value: [
            {
              id: "series-master",
              changeKey: "master-1",
              subject: "Weekly Sync",
              type: "seriesMaster",
              recurrence: {},
              start: {
                dateTime: "2026-05-22T10:00:00.0000000",
                timeZone: "UTC",
              },
              end: {
                dateTime: "2026-05-22T11:00:00.0000000",
                timeZone: "UTC",
              },
            },
            {
              id: "occ-1",
              changeKey: "occ-1",
              subject: "Weekly Sync",
              type: "occurrence",
              start: {
                dateTime: "2026-05-22T10:00:00.0000000",
                timeZone: "UTC",
              },
              end: {
                dateTime: "2026-05-22T11:00:00.0000000",
                timeZone: "UTC",
              },
            },
          ],
          "@odata.deltaLink": "https://graph.microsoft.com/v1.0/me/calendars/primary/calendarView/delta?$deltatoken=done",
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

    const result = await source.listMeetingChanges({
      calendar: CALENDAR,
      window: {
        start: "2026-05-22T00:00:00.000Z",
        end: "2026-05-23T00:00:00.000Z",
      },
    });

    expect(result.meetings.map((meeting) => meeting.id)).toEqual(["occ-1"]);
  });

  it("retries throttled Graph requests using Retry-After", async () => {
    vi.useFakeTimers();
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            error: {
              code: "ApplicationThrottled",
              message: "Application is over its MailboxConcurrency limit.",
            },
          }),
          {
            status: 429,
            headers: {
              "Content-Type": "application/json",
              "Retry-After": "1",
            },
          },
        ),
      )
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            value: [
              {
                id: "evt-5",
                changeKey: "ck-5",
                subject: "Retried",
                start: {
                  dateTime: "2026-05-22T16:00:00.0000000",
                  timeZone: "UTC",
                },
                end: {
                  dateTime: "2026-05-22T17:00:00.0000000",
                  timeZone: "UTC",
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

    const pending = source.listMeetings({
      calendar: CALENDAR,
      window: {
        start: "2026-05-22T00:00:00.000Z",
        end: "2026-05-23T00:00:00.000Z",
      },
    });

    await vi.advanceTimersByTimeAsync(1000);
    const meetings = await pending;

    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(meetings[0]?.id).toBe("evt-5");
    vi.useRealTimers();
  });
});
