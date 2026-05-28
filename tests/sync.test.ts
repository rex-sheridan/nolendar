import { describe, expect, it, vi } from "vitest";

import type { NolendarConfig } from "../src/domain/config.js";
import type { Meeting } from "../src/domain/meeting.js";
import { syncMeetingsToNotion } from "../src/sync.js";

const CONFIG: NolendarConfig = {
  microsoft: {
    tenant: "common",
    authMode: "device_code",
  },
  notion: {
    databaseId: "data-source-id",
    defaultTags: ["meeting"],
    defaultAssigneeEmail: "me@example.com",
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
  it("treats recurring occurrences as separate syncable meetings by event id", async () => {
    const notion = {
      retrieveDataSource: vi.fn(async () => ({
        id: "data-source-id",
        title: "Meetings",
        properties: {
          Name: { id: "title", name: "Name", type: "title" },
          Due: { id: "due", name: "Due", type: "date" },
          "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
          "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
          "Source URL": { id: "source-url", name: "Source URL", type: "url" },
          Tags: { id: "tags", name: "Tags", type: "multi_select" },
          Assignee: { id: "assignee", name: "Assignee", type: "people" },
        },
      })),
      getDefaultAssigneeUserId: vi.fn(async () => "user-1"),
      getTemplateBlocks: vi.fn(async () => []),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => null),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      setPageStatus: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    const result = await syncMeetingsToNotion(
      CONFIG,
      [
        {
          ...MEETING,
          id: "occ-1",
          title: "Weekly Sync",
          isRecurring: true,
        },
        {
          ...MEETING,
          id: "occ-2",
          title: "Weekly Sync",
          start: "2026-05-29T13:00:00.000Z",
          end: "2026-05-29T14:00:00.000Z",
          isRecurring: true,
        },
      ],
      notion,
    );

    expect(result.created).toBe(2);
    expect(notion.createMeetingPage).toHaveBeenCalledTimes(2);
  });

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
          "Source URL": { id: "source-url", name: "Source URL", type: "url" },
          Tags: { id: "tags", name: "Tags", type: "multi_select" },
          Assignee: { id: "assignee", name: "Assignee", type: "people" },
        },
      })),
      getDefaultAssigneeUserId: vi.fn(async () => "user-1"),
      getTemplateBlocks: vi.fn(async () => []),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => null),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    const result = await syncMeetingsToNotion(CONFIG, [MEETING], notion);

    expect(result).toEqual({
      created: 1,
      updated: 0,
      archived: 0,
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
          "Source URL": { id: "source-url", name: "Source URL", type: "url" },
          Tags: { id: "tags", name: "Tags", type: "multi_select" },
          Assignee: { id: "assignee", name: "Assignee", type: "people" },
        },
      })),
      getDefaultAssigneeUserId: vi.fn(async () => "user-1"),
      getTemplateBlocks: vi.fn(async () => []),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => ({ id: "page-1", eventId: "evt-1", changeKey: "ck-1" })),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
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
          "Source URL": { id: "source-url", name: "Source URL", type: "url" },
          Tags: { id: "tags", name: "Tags", type: "multi_select" },
          Assignee: { id: "assignee", name: "Assignee", type: "people" },
        },
      })),
      getDefaultAssigneeUserId: vi.fn(async () => "user-1"),
      getTemplateBlocks: vi.fn(async () => []),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => ({ id: "page-1", eventId: "evt-1", changeKey: "old" })),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
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
          "Source URL": { id: "source-url", name: "Source URL", type: "url" },
          Tags: { id: "tags", name: "Tags", type: "multi_select" },
          Assignee: { id: "assignee", name: "Assignee", type: "people" },
        },
      })),
      getDefaultAssigneeUserId: vi.fn(async () => "user-1"),
      getTemplateBlocks: vi.fn(async () => []),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => null),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    const result = await syncMeetingsToNotion(CONFIG, [MEETING], notion, { dryRun: true });

    expect(result).toEqual({
      created: 1,
      updated: 0,
      archived: 0,
      skipped: 0,
      filtered: 0,
      dryRun: true,
    });
    expect(notion.createMeetingPage).not.toHaveBeenCalled();
    expect(notion.updateMeetingPage).not.toHaveBeenCalled();
  });

  it("force-updates an existing page when the changeKey matches", async () => {
    const notion = {
      retrieveDataSource: vi.fn(async () => ({
        id: "data-source-id",
        title: "Meetings",
        properties: {
          Name: { id: "title", name: "Name", type: "title" },
          Due: { id: "due", name: "Due", type: "date" },
          "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
          "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
          "Source URL": { id: "source-url", name: "Source URL", type: "url" },
          Tags: { id: "tags", name: "Tags", type: "multi_select" },
          Assignee: { id: "assignee", name: "Assignee", type: "people" },
        },
      })),
      getDefaultAssigneeUserId: vi.fn(async () => "user-1"),
      getTemplateBlocks: vi.fn(async () => []),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => ({ id: "page-1", eventId: "evt-1", changeKey: "ck-1" })),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    const result = await syncMeetingsToNotion(CONFIG, [MEETING], notion, { forceUpdate: true });

    expect(result).toEqual({
      created: 0,
      updated: 1,
      archived: 0,
      skipped: 0,
      filtered: 0,
      dryRun: false,
    });
    expect(notion.updateMeetingPage).toHaveBeenCalledTimes(1);
    expect(notion.createMeetingPage).not.toHaveBeenCalled();
  });

  it("archives an existing page for a cancelled meeting instead of updating it", async () => {
    const notion = {
      retrieveDataSource: vi.fn(async () => ({
        id: "data-source-id",
        title: "Meetings",
        properties: {
          Name: { id: "title", name: "Name", type: "title" },
          Due: { id: "due", name: "Due", type: "date" },
          "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
          "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
          "Source URL": { id: "source-url", name: "Source URL", type: "url" },
          Tags: { id: "tags", name: "Tags", type: "multi_select" },
          Assignee: { id: "assignee", name: "Assignee", type: "people" },
        },
      })),
      getDefaultAssigneeUserId: vi.fn(async () => "user-1"),
      getTemplateBlocks: vi.fn(async () => []),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => ({ id: "page-1", eventId: "evt-1", changeKey: "old" })),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    const result = await syncMeetingsToNotion(
      CONFIG,
      [
        {
          ...MEETING,
          isCancelled: true,
        },
      ],
      notion,
    );

    expect(result).toEqual({
      created: 0,
      updated: 0,
      archived: 1,
      skipped: 0,
      filtered: 0,
      dryRun: false,
    });
    expect(notion.archivePage).toHaveBeenCalledWith("page-1");
    expect(notion.createMeetingPage).not.toHaveBeenCalled();
    expect(notion.updateMeetingPage).not.toHaveBeenCalled();
  });

  it("sets a configured status for a cancelled meeting when archiving is disabled", async () => {
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
          "Source URL": { id: "source-url", name: "Source URL", type: "url" },
          Tags: { id: "tags", name: "Tags", type: "multi_select" },
          Assignee: { id: "assignee", name: "Assignee", type: "people" },
          Status: { id: "status", name: "Status", type: "status" },
        },
      })),
      getDefaultAssigneeUserId: vi.fn(async () => "user-1"),
      getTemplateBlocks: vi.fn(async () => []),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => ({ id: "page-1", eventId: "evt-1", changeKey: "old" })),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      setPageStatus: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    const result = await syncMeetingsToNotion(
      config,
      [
        {
          ...MEETING,
          isCancelled: true,
        },
      ],
      notion,
    );

    expect(result).toEqual({
      created: 0,
      updated: 1,
      archived: 0,
      skipped: 0,
      filtered: 0,
      dryRun: false,
    });
    expect(notion.setPageStatus).toHaveBeenCalledWith({
      pageId: "page-1",
      propertyName: "Status",
      statusName: "Canceled",
    });
    expect(notion.archivePage).not.toHaveBeenCalled();
    expect(notion.createMeetingPage).not.toHaveBeenCalled();
    expect(notion.updateMeetingPage).not.toHaveBeenCalled();
  });

  it("handles cancelled meetings before normal filters are applied", async () => {
    const config: NolendarConfig = {
      ...CONFIG,
      filters: {
        ...CONFIG.filters,
        ignoreDeclined: true,
        ignorePatterns: [".*Planning.*"],
      },
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
          "Source URL": { id: "source-url", name: "Source URL", type: "url" },
          Tags: { id: "tags", name: "Tags", type: "multi_select" },
          Assignee: { id: "assignee", name: "Assignee", type: "people" },
          Status: { id: "status", name: "Status", type: "status" },
        },
      })),
      getDefaultAssigneeUserId: vi.fn(async () => "user-1"),
      getTemplateBlocks: vi.fn(async () => []),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => ({ id: "page-1", eventId: "evt-1", changeKey: "old" })),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      setPageStatus: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    const result = await syncMeetingsToNotion(
      config,
      [
        {
          ...MEETING,
          isCancelled: true,
          responseStatus: "declined",
        },
      ],
      notion,
    );

    expect(result).toEqual({
      created: 0,
      updated: 1,
      archived: 0,
      skipped: 0,
      filtered: 0,
      dryRun: false,
    });
    expect(notion.setPageStatus).toHaveBeenCalledWith({
      pageId: "page-1",
      propertyName: "Status",
      statusName: "Canceled",
    });
    expect(notion.archivePage).not.toHaveBeenCalled();
  });

  it("reports verbose decisions for cancelled status updates", async () => {
    const decisions: string[] = [];
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
          "Source URL": { id: "source-url", name: "Source URL", type: "url" },
          Tags: { id: "tags", name: "Tags", type: "multi_select" },
          Assignee: { id: "assignee", name: "Assignee", type: "people" },
          Status: { id: "status", name: "Status", type: "status" },
        },
      })),
      getDefaultAssigneeUserId: vi.fn(async () => "user-1"),
      getTemplateBlocks: vi.fn(async () => []),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => ({ id: "page-1", eventId: "evt-1", changeKey: "old" })),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      setPageStatus: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    await syncMeetingsToNotion(
      config,
      [
        {
          ...MEETING,
          isCancelled: true,
        },
      ],
      notion,
      {
        onDecision: (message) => decisions.push(message),
      },
    );

    expect(decisions).toContain(
      'sync decision: title="Planning" eventId=evt-1 pageId=page-1 decision=set_status_cancelled property=Status value=Canceled',
    );
  });

  it("does not create a new page for a cancelled meeting that has not been synced before", async () => {
    const notion = {
      retrieveDataSource: vi.fn(async () => ({
        id: "data-source-id",
        title: "Meetings",
        properties: {
          Name: { id: "title", name: "Name", type: "title" },
          Due: { id: "due", name: "Due", type: "date" },
          "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
          "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
          "Source URL": { id: "source-url", name: "Source URL", type: "url" },
          Tags: { id: "tags", name: "Tags", type: "multi_select" },
          Assignee: { id: "assignee", name: "Assignee", type: "people" },
        },
      })),
      getDefaultAssigneeUserId: vi.fn(async () => "user-1"),
      getTemplateBlocks: vi.fn(async () => []),
      ensureProperties: vi.fn(async () => undefined),
      findPageByEventId: vi.fn(async () => null),
      createMeetingPage: vi.fn(async () => ({ id: "page-1" })),
      updateMeetingPage: vi.fn(async () => undefined),
      archivePage: vi.fn(async () => undefined),
      finalizeMeetingPageContent: vi.fn(async () => "appended"),
    };

    const result = await syncMeetingsToNotion(
      CONFIG,
      [
        {
          ...MEETING,
          isCancelled: true,
        },
      ],
      notion,
    );

    expect(result).toEqual({
      created: 0,
      updated: 0,
      archived: 0,
      skipped: 1,
      filtered: 0,
      dryRun: false,
    });
    expect(notion.archivePage).not.toHaveBeenCalled();
    expect(notion.createMeetingPage).not.toHaveBeenCalled();
    expect(notion.updateMeetingPage).not.toHaveBeenCalled();
  });
});
