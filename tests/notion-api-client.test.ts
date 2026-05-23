import { describe, expect, it, vi } from "vitest";

import { ApiNotionClient } from "../src/notion/api-notion-client.js";

describe("ApiNotionClient.getDefaultAssigneeUserId", () => {
  it("resolves a people property assignee from defaultAssigneeEmail", async () => {
    const client = new ApiNotionClient("token", {
      dataSources: {
        retrieve: vi.fn(),
        update: vi.fn(),
        query: vi.fn(),
      },
      pages: {
        create: vi.fn(),
        update: vi.fn(),
      },
      users: {
        me: vi.fn(async () => ({
          id: "bot-user",
          type: "bot",
        })),
        list: vi.fn(async () => ({
          results: [
            {
              id: "user-123",
              type: "person",
              person: {
                email: "me@example.com",
              },
            },
          ],
          has_more: false,
          next_cursor: null,
        })),
      },
    });

    await expect(client.getDefaultAssigneeUserId("me@example.com")).resolves.toBe("user-123");
  });

  it("falls back to the authenticated user when email lookup does not match", async () => {
    const client = new ApiNotionClient("token", {
      dataSources: {
        retrieve: vi.fn(),
        update: vi.fn(),
        query: vi.fn(),
      },
      pages: {
        create: vi.fn(),
        update: vi.fn(),
      },
      users: {
        me: vi.fn(async () => ({
          id: "person-456",
          type: "person",
        })),
        list: vi.fn(async () => ({
          results: [],
          has_more: false,
          next_cursor: null,
        })),
      },
    });

    await expect(client.getDefaultAssigneeUserId("missing@example.com")).resolves.toBe("person-456");
  });
});

describe("ApiNotionClient page icon writes", () => {
  it("sends an emoji page icon when configured", async () => {
    const pages = {
      create: vi.fn(async () => ({ id: "page-1" })),
      update: vi.fn(async () => undefined),
    };
    const client = new ApiNotionClient("token", {
      dataSources: {
        retrieve: vi.fn(),
        update: vi.fn(),
        query: vi.fn(),
      },
      pages,
      users: {
        me: vi.fn(),
        list: vi.fn(),
      },
    });

    await client.createMeetingPage({
      config: {
        microsoft: { tenant: "common", authMode: "device_code" },
        notion: {
          databaseId: "data-source-id",
          pageIcon: {
            type: "emoji",
            emoji: "📝",
          },
        },
        calendars: [],
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
      },
      dataSource: {
        id: "data-source-id",
        properties: {
          Name: { id: "title", name: "Name", type: "title" },
          Due: { id: "due", name: "Due", type: "date" },
          "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
          "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
        },
      },
      meeting: {
        id: "evt-1",
        changeKey: "ck-1",
        calendarId: "primary",
        title: "Planning",
        start: "2026-05-22T13:00:00.000Z",
        end: "2026-05-22T14:00:00.000Z",
        attendees: [],
        isCancelled: false,
        isRecurring: false,
      },
    });

    expect(pages.create).toHaveBeenCalledWith(
      expect.objectContaining({
        icon: {
          type: "emoji",
          emoji: "📝",
        },
      }),
    );
  });

  it("sends a native notion icon with color when configured", async () => {
    const pages = {
      create: vi.fn(async () => ({ id: "page-1" })),
      update: vi.fn(async () => undefined),
    };
    const client = new ApiNotionClient("token", {
      dataSources: {
        retrieve: vi.fn(),
        update: vi.fn(),
        query: vi.fn(),
      },
      pages,
      users: {
        me: vi.fn(),
        list: vi.fn(),
      },
    });

    await client.updateMeetingPage({
      pageId: "page-1",
      config: {
        microsoft: { tenant: "common", authMode: "device_code" },
        notion: {
          databaseId: "data-source-id",
          pageIcon: {
            type: "icon",
            name: "calendar",
            color: "blue",
          },
        },
        calendars: [],
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
      },
      dataSource: {
        id: "data-source-id",
        properties: {
          Name: { id: "title", name: "Name", type: "title" },
          Due: { id: "due", name: "Due", type: "date" },
          "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
          "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
        },
      },
      meeting: {
        id: "evt-1",
        changeKey: "ck-1",
        calendarId: "primary",
        title: "Planning",
        start: "2026-05-22T13:00:00.000Z",
        end: "2026-05-22T14:00:00.000Z",
        attendees: [],
        isCancelled: false,
        isRecurring: false,
      },
    });

    expect(pages.update).toHaveBeenCalledWith(
      expect.objectContaining({
        icon: {
          type: "icon",
          icon: {
            name: "calendar",
            color: "blue",
          },
        },
      }),
    );
  });

  it("creates missing participant people pages and relates existing matches by email", async () => {
    const dataSources = {
      retrieve: vi.fn(),
      update: vi.fn(),
      query: vi
        .fn()
        .mockResolvedValueOnce({
          results: [
            {
              id: "person-existing",
              object: "page",
            },
          ],
        })
        .mockResolvedValueOnce({
          results: [],
        }),
    };
    const pages = {
      create: vi
        .fn()
        .mockResolvedValueOnce({ id: "person-created" })
        .mockResolvedValueOnce({ id: "meeting-page-1" }),
      update: vi.fn(async () => undefined),
    };
    const client = new ApiNotionClient("token", {
      dataSources,
      pages,
      users: {
        me: vi.fn(),
        list: vi.fn(),
      },
    });

    await client.createMeetingPage({
      config: {
        microsoft: { tenant: "common", authMode: "device_code" },
        notion: {
          databaseId: "meetings-data-source-id",
          peopleDataSource: {
            databaseId: "people-data-source-id",
            nameProperty: "Name",
            emailProperty: "Email Address",
          },
        },
        calendars: [],
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
          participants: "Participants",
        },
        sync: {
          lookahead: "today",
          statePath: "/tmp/.nolendar/state.json",
        },
      },
      dataSource: {
        id: "meetings-data-source-id",
        properties: {
          Name: { id: "title", name: "Name", type: "title" },
          Due: { id: "due", name: "Due", type: "date" },
          "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
          "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
          Participants: { id: "participants", name: "Participants", type: "relation" },
        },
      },
      meeting: {
        id: "evt-1",
        changeKey: "ck-1",
        calendarId: "primary",
        title: "Planning",
        start: "2026-05-22T13:00:00.000Z",
        end: "2026-05-22T14:00:00.000Z",
        attendees: [
          { name: "Alex Existing", email: "existing@example.com", optional: false },
          { name: "Casey New", email: "new@example.com", optional: false },
        ],
        isCancelled: false,
        isRecurring: false,
      },
    });

    expect(dataSources.query).toHaveBeenNthCalledWith(
      1,
      expect.objectContaining({
        data_source_id: "people-data-source-id",
        filter: {
          property: "Email Address",
          email: {
            equals: "existing@example.com",
          },
        },
      }),
    );
    expect(dataSources.query).toHaveBeenNthCalledWith(
      2,
      expect.objectContaining({
        data_source_id: "people-data-source-id",
        filter: {
          property: "Email Address",
          email: {
            equals: "new@example.com",
          },
        },
      }),
    );
    expect(pages.create).toHaveBeenNthCalledWith(
      1,
      expect.objectContaining({
        parent: {
          data_source_id: "people-data-source-id",
        },
        properties: {
          Name: {
            title: [{ type: "text", text: { content: "Casey New" } }],
          },
          "Email Address": {
            email: "new@example.com",
          },
        },
      }),
    );
    expect(pages.create).toHaveBeenNthCalledWith(
      2,
      expect.objectContaining({
        properties: expect.objectContaining({
          Participants: {
            relation: [{ id: "person-existing" }, { id: "person-created" }],
          },
        }),
      }),
    );
  });
});
