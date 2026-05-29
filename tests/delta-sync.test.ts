import { describe, expect, it, vi } from "vitest";

import type { NolendarConfig } from "../src/domain/config.js";
import type { Meeting } from "../src/domain/meeting.js";
import { syncCalendarChangesToNotion } from "../src/delta-sync.js";

const CONFIG: NolendarConfig = {
  microsoft: {
    tenant: "common",
    authMode: "device_code",
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
  start: "2026-05-23T13:00:00.000Z",
  end: "2026-05-23T14:00:00.000Z",
  attendees: [],
  isCancelled: false,
  isRecurring: false,
};

describe("syncCalendarChangesToNotion", () => {
  it("reuses a stored delta link when the saved window still matches", async () => {
    const meetingSource = {
      listMeetingChanges: vi.fn(async () => ({
        meetings: [MEETING],
        removedEventIds: [],
        deltaLink: "delta-2",
      })),
    };
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
      findPageByEventId: vi.fn(async () => null),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };
    const loadState = vi.fn(async () => ({
      version: 1 as const,
      calendars: {
        primary: {
          lookahead: "today" as const,
          window: {
            start: "2026-05-23T00:00:00.000Z",
            end: "2026-05-24T00:00:00.000Z",
          },
          deltaLink: "delta-1",
          updatedAt: "2026-05-23T12:00:00.000Z",
        },
      },
    }));
    const saveState = vi.fn(async () => undefined);

    const result = await syncCalendarChangesToNotion(CONFIG, notion, {
      meetingSource,
      loadState,
      saveState,
      clock: {
        now: () => new Date("2026-05-23T15:00:00.000Z"),
      },
    });

    expect(meetingSource.listMeetingChanges).toHaveBeenCalledWith({
      calendar: CONFIG.calendars[0],
      window: {
        start: "2026-05-23T00:00:00.000Z",
        end: "2026-05-24T00:00:00.000Z",
      },
      deltaLink: "delta-1",
    });
    expect(saveState).toHaveBeenCalledWith(CONFIG.sync.statePath, {
      version: 1,
      calendars: {
        primary: {
          lookahead: "today",
          window: {
            start: "2026-05-23T00:00:00.000Z",
            end: "2026-05-24T00:00:00.000Z",
          },
          deltaLink: "delta-2",
          updatedAt: "2026-05-23T15:00:00.000Z",
        },
      },
    });
    expect(result.created).toBe(1);
    expect(result.archived).toBe(0);
  });

  it("falls back to a fresh calendar view delta when the saved window no longer matches", async () => {
    const meetingSource = {
      listMeetingChanges: vi.fn(async () => ({
        meetings: [MEETING],
        removedEventIds: [],
        deltaLink: "delta-fresh",
      })),
    };
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
      findPageByEventId: vi.fn(async () => null),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    await syncCalendarChangesToNotion(CONFIG, notion, {
      meetingSource,
      loadState: async () => ({
        version: 1,
        calendars: {
          primary: {
            lookahead: "today",
            window: {
              start: "2026-05-22T00:00:00.000Z",
              end: "2026-05-23T00:00:00.000Z",
            },
            deltaLink: "stale-delta",
            updatedAt: "2026-05-22T12:00:00.000Z",
          },
        },
      }),
      saveState: vi.fn(async () => undefined),
      clock: {
        now: () => new Date("2026-05-23T15:00:00.000Z"),
      },
    });

    expect(meetingSource.listMeetingChanges).toHaveBeenCalledWith({
      calendar: CONFIG.calendars[0],
      window: {
        start: "2026-05-23T00:00:00.000Z",
        end: "2026-05-24T00:00:00.000Z",
      },
      deltaLink: undefined,
    });
  });

  it("does not persist state during dry-run mode", async () => {
    const saveState = vi.fn(async () => undefined);

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
      findPageByEventId: vi.fn(async () => null),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    await syncCalendarChangesToNotion(CONFIG, notion, {
      meetingSource: {
        listMeetingChanges: vi.fn(async () => ({
          meetings: [MEETING],
          removedEventIds: [],
          deltaLink: "delta-1",
        })),
      },
      loadState: vi.fn(async () => ({
        version: 1 as const,
        calendars: {},
      })),
      saveState,
      clock: {
        now: () => new Date("2026-05-23T15:00:00.000Z"),
      },
      syncOptions: {
        dryRun: true,
      },
    });

    expect(saveState).not.toHaveBeenCalled();
  });

  it("adds archive counts from cancelled meetings returned by delta sync", async () => {
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
        .mockResolvedValueOnce({ id: "page-1", eventId: "evt-1", changeKey: "old" })
        .mockResolvedValueOnce({ id: "page-2", eventId: "evt-removed", changeKey: "old" }),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    const result = await syncCalendarChangesToNotion(CONFIG, notion, {
      meetingSource: {
        listMeetingChanges: vi.fn(async () => ({
          meetings: [
            {
              ...MEETING,
              isCancelled: true,
            },
          ],
          removedEventIds: ["evt-removed"],
          deltaLink: "delta-1",
        })),
      },
      loadState: vi.fn(async () => ({
        version: 1 as const,
        calendars: {},
      })),
      saveState: vi.fn(async () => undefined),
      clock: {
        now: () => new Date("2026-05-23T15:00:00.000Z"),
      },
    });

    expect(result.archived).toBe(2);
    expect(notion.archivePage).toHaveBeenCalledTimes(2);
  });

  it("syncs calendars sequentially instead of in parallel", async () => {
    const order: string[] = [];
    let activeCalls = 0;
    let maxActiveCalls = 0;
    const config: NolendarConfig = {
      ...CONFIG,
      calendars: [
        { id: "primary", name: "Primary" },
        { id: "secondary", name: "Secondary" },
      ],
    };
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
      findPageByEventId: vi.fn(async () => null),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    await syncCalendarChangesToNotion(config, notion, {
      meetingSource: {
        listMeetingChanges: vi.fn(async ({ calendar }) => {
          order.push(`${calendar.id}:start`);
          activeCalls += 1;
          maxActiveCalls = Math.max(maxActiveCalls, activeCalls);
          await Promise.resolve();
          activeCalls -= 1;
          order.push(`${calendar.id}:end`);
          return {
            meetings: [{ ...MEETING, calendarId: calendar.id, calendarName: calendar.name }],
            removedEventIds: [],
            deltaLink: `delta-${calendar.id}`,
          };
        }),
      },
      loadState: vi.fn(async () => ({
        version: 1 as const,
        calendars: {},
      })),
      saveState: vi.fn(async () => undefined),
      clock: {
        now: () => new Date("2026-05-23T15:00:00.000Z"),
      },
    });

    expect(maxActiveCalls).toBe(1);
    expect(order).toEqual(["primary:start", "primary:end", "secondary:start", "secondary:end"]);
  });

  it("archives matching Notion pages for removed events", async () => {
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
        .mockResolvedValueOnce({ id: "page-removed", eventId: "evt-removed", changeKey: "old" })
        .mockResolvedValueOnce(null),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    const result = await syncCalendarChangesToNotion(CONFIG, notion, {
      meetingSource: {
        listMeetingChanges: vi.fn(async () => ({
          meetings: [MEETING],
          removedEventIds: ["evt-removed", "evt-missing"],
          deltaLink: "delta-1",
        })),
      },
      loadState: vi.fn(async () => ({
        version: 1 as const,
        calendars: {},
      })),
      saveState: vi.fn(async () => undefined),
      clock: {
        now: () => new Date("2026-05-23T15:00:00.000Z"),
      },
    });

    expect(notion.archivePage).toHaveBeenCalledWith("page-removed");
    expect(result.archived).toBe(1);
  });

  it("sets the configured status for removed events when cancelled meetings use status mapping", async () => {
    const config: NolendarConfig = {
      ...CONFIG,
      notion: {
        ...CONFIG.notion,
        canceledMeetings: {
          action: "set_status",
          statusProperty: "Status",
          statusValue: "Canceled",
        },
      },
    };
    const notion = {
      retrieveDataSource: vi.fn(async () => ({
        id: "data-source-id",
        title: "Meetings",
        properties: {
          Name: { id: "title", name: "Name", type: "title" },
          Due: { id: "due", name: "Due", type: "date" },
          "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
          "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
          Status: { id: "status", name: "Status", type: "status" },
        },
      })),
      getDefaultAssigneeUserId: vi.fn(async () => undefined),
      getTemplateBlocks: vi.fn(async () => []),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi
        .fn()
        .mockResolvedValueOnce({ id: "page-removed", eventId: "evt-removed", changeKey: "old" })
        .mockResolvedValueOnce(null),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      setPageStatus: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    const result = await syncCalendarChangesToNotion(config, notion, {
      meetingSource: {
        listMeetingChanges: vi.fn(async () => ({
          meetings: [MEETING],
          removedEventIds: ["evt-removed", "evt-missing"],
          deltaLink: "delta-1",
        })),
      },
      loadState: vi.fn(async () => ({
        version: 1 as const,
        calendars: {},
      })),
      saveState: vi.fn(async () => undefined),
      clock: {
        now: () => new Date("2026-05-23T15:00:00.000Z"),
      },
    });

    expect(result.updated).toBe(1);
    expect(result.archived).toBe(0);
    expect(notion.setPageStatus).toHaveBeenCalledWith({
      pageId: "page-removed",
      propertyName: "Status",
      statusName: "Canceled",
    });
    expect(notion.archivePage).not.toHaveBeenCalled();
  });

  it("sets the configured status for Notion pages missing from a fresh Outlook window", async () => {
    const config: NolendarConfig = {
      ...CONFIG,
      notion: {
        ...CONFIG.notion,
        canceledMeetings: {
          action: "set_status",
          statusProperty: "Status",
          statusValue: "Canceled",
        },
      },
    };
    const notion = {
      retrieveDataSource: vi.fn(async () => ({
        id: "data-source-id",
        title: "Meetings",
        properties: {
          Name: { id: "title", name: "Name", type: "title" },
          Due: { id: "due", name: "Due", type: "date" },
          "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
          "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
          Status: { id: "status", name: "Status", type: "status" },
        },
      })),
      getDefaultAssigneeUserId: vi.fn(async () => undefined),
      getTemplateBlocks: vi.fn(async () => []),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => null),
      listMeetingPagePropertiesForWindow: vi.fn(async () => [
        {
          id: "page-missing",
          properties: {
            "Outlook Event ID": "evt-missing-from-outlook",
            Status: "Scheduled",
          },
          body: "",
        },
      ]),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      setPageStatus: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    const result = await syncCalendarChangesToNotion(config, notion, {
      meetingSource: {
        listMeetingChanges: vi.fn(async () => ({
          meetings: [],
          removedEventIds: [],
          deltaLink: "delta-1",
        })),
      },
      loadState: vi.fn(async () => ({
        version: 1 as const,
        calendars: {},
      })),
      saveState: vi.fn(async () => undefined),
      clock: {
        now: () => new Date("2026-05-23T15:00:00.000Z"),
      },
    });

    expect(result.updated).toBe(1);
    expect(result.archived).toBe(0);
    expect(notion.listMeetingPagePropertiesForWindow).toHaveBeenCalledWith({
      dataSourceId: "data-source-id",
      datePropertyName: "Due",
      start: "2026-05-23T00:00:00.000Z",
      end: "2026-05-24T00:00:00.000Z",
    });
    expect(notion.setPageStatus).toHaveBeenCalledWith({
      pageId: "page-missing",
      propertyName: "Status",
      statusName: "Canceled",
    });
  });

  it("preserves the Notion client receiver when reconciling missing pages", async () => {
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
      findPageByEventId: vi.fn(async () => ({
        id: "page-1",
        eventId: "evt-1",
        changeKey: "ck-1",
      })),
      listMeetingPagePropertiesForWindow: vi.fn(async function (this: { marker?: string }) {
        if (this.marker !== "bound-notion") {
          throw new Error("Notion method receiver was not preserved.");
        }

        return [];
      }),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
      marker: "bound-notion",
    };

    await expect(
      syncCalendarChangesToNotion(CONFIG, notion, {
        meetingSource: {
          listMeetingChanges: vi.fn(async () => ({
            meetings: [MEETING],
            removedEventIds: [],
            deltaLink: "delta-1",
          })),
        },
        loadState: vi.fn(async () => ({
          version: 1 as const,
          calendars: {},
        })),
        saveState: vi.fn(async () => undefined),
        clock: {
          now: () => new Date("2026-05-23T15:00:00.000Z"),
        },
      }),
    ).resolves.toMatchObject({
      skipped: 1,
    });
  });
});
