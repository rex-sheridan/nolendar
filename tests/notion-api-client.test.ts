import { describe, expect, it, vi } from "vitest";

import type { NolendarConfig } from "../src/domain/config.js";
import type { Meeting } from "../src/domain/meeting.js";
import { ApiNotionClient } from "../src/notion/api-notion-client.js";

describe("ApiNotionClient.getDefaultAssigneeUserId", () => {
  it("resolves a people property assignee from defaultAssigneeEmail", async () => {
    const client = new ApiNotionClient("token", {
      blocks: {
        children: {
          append: vi.fn(async () => undefined),
          list: vi.fn(),
        },
      },
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
      blocks: {
        children: {
          append: vi.fn(async () => undefined),
          list: vi.fn(),
        },
      },
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
  it("records timing entries for Notion API calls when a reporter is configured", async () => {
    const timingReporter = {
      record: vi.fn(),
    };
    const pages = {
      create: vi.fn(async () => ({ id: "page-1" })),
      update: vi.fn(async () => undefined),
    };
    const client = new ApiNotionClient(
      "token",
      {
        blocks: {
          children: {
            append: vi.fn(async () => undefined),
            list: vi.fn(async () => ({
              results: [],
              has_more: false,
              next_cursor: null,
            })),
          },
        },
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
      },
      timingReporter,
    );

    await client.createMeetingPage({
      config: {
        microsoft: { tenant: "common", authMode: "device_code" },
        notion: {
          databaseId: "data-source-id",
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

    expect(timingReporter.record).toHaveBeenCalledWith(
      expect.objectContaining({
        service: "notion",
        operation: "pages.create",
        detail: "parent_data_source_id=data-source-id template_blocks=0 children=4",
        status: "ok",
      }),
    );
  });

  it("finalizes native Notion template pages by appending generated meeting blocks once", async () => {
    const timingReporter = {
      record: vi.fn(),
    };
    const blocks = {
      children: {
        append: vi.fn(async () => undefined),
        list: vi.fn(async () => ({
          results: [],
          has_more: false,
          next_cursor: null,
        })),
      },
    };
    const client = new ApiNotionClient(
      "token",
      {
        blocks,
        dataSources: {
          retrieve: vi.fn(),
          update: vi.fn(),
          query: vi.fn(),
        },
        pages: {
          create: vi.fn(async () => ({ id: "page-1" })),
          update: vi.fn(async () => undefined),
        },
        users: {
          me: vi.fn(),
          list: vi.fn(),
        },
      },
      timingReporter,
    );

    const config: NolendarConfig = {
      microsoft: { tenant: "common", authMode: "device_code" },
      notion: {
        databaseId: "data-source-id",
        dataSourceTemplate: {
          type: "default",
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
    };
    const dataSource = {
      id: "data-source-id",
      properties: {
        Name: { id: "title", name: "Name", type: "title" },
        Due: { id: "due", name: "Due", type: "date" },
        "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
        "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
      },
    } as const;
    const meeting: Meeting = {
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

    await client.finalizeMeetingPageContent({
      pageId: "page-1",
      config,
      dataSource,
      meeting,
    });

    expect(blocks.children.append).toHaveBeenCalledWith({
      block_id: "page-1",
      children: expect.any(Array),
    });
    expect(blocks.children.list).toHaveBeenCalledWith({
      block_id: "page-1",
      start_cursor: undefined,
      page_size: 100,
    });
    expect(timingReporter.record).toHaveBeenCalledWith(
      expect.objectContaining({
        service: "notion",
        operation: "blocks.children.append",
        detail: "block_id=page-1 children=4",
      }),
    );
  });

  it("sends an emoji page icon when configured", async () => {
    const pages = {
      create: vi.fn(async () => ({ id: "page-1" })),
      update: vi.fn(async () => undefined),
    };
    const client = new ApiNotionClient("token", {
      blocks: {
        children: {
          append: vi.fn(async () => undefined),
          list: vi.fn(),
        },
      },
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
      blocks: {
        children: {
          append: vi.fn(async () => undefined),
          list: vi.fn(),
        },
      },
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
      blocks: {
        children: {
          append: vi.fn(async () => undefined),
          list: vi.fn(),
        },
      },
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

  it("limits participant People relation resolution per meeting", async () => {
    const dataSources = {
      retrieve: vi.fn(),
      update: vi.fn(),
      query: vi.fn(async () => ({
        results: [],
      })),
    };
    const pages = {
      create: vi
        .fn()
        .mockResolvedValueOnce({ id: "person-created" })
        .mockResolvedValueOnce({ id: "meeting-page-1" }),
      update: vi.fn(async () => undefined),
    };
    const client = new ApiNotionClient("token", {
      blocks: {
        children: {
          append: vi.fn(async () => undefined),
          list: vi.fn(),
        },
      },
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
            maxAttendeesPerMeeting: 1,
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
          { name: "Alex", email: "alex@example.com", optional: false },
          { name: "Casey", email: "casey@example.com", optional: false },
        ],
        isCancelled: false,
        isRecurring: false,
      },
    });

    expect(dataSources.query).toHaveBeenCalledTimes(1);
    expect(dataSources.query).toHaveBeenCalledWith(
      expect.objectContaining({
        filter: {
          property: "Email Address",
          email: {
            equals: "alex@example.com",
          },
        },
      }),
    );
    expect(pages.create).toHaveBeenNthCalledWith(
      2,
      expect.objectContaining({
        properties: expect.objectContaining({
          Participants: {
            relation: [{ id: "person-created" }],
          },
        }),
      }),
    );
  });

  it("prepends configured template blocks when creating a meeting page", async () => {
    const blocks = {
      children: {
        append: vi.fn(async () => undefined),
        list: vi.fn(async ({ block_id }: { block_id: string }) => {
          if (block_id === "template-page-id") {
            return {
              results: [
                {
                  id: "block-1",
                  object: "block",
                  type: "heading_1",
                  has_children: false,
                  heading_1: {
                    rich_text: [
                      {
                        type: "text",
                        text: {
                          content: "Template Heading",
                        },
                      },
                    ],
                  },
                },
              ],
              has_more: false,
              next_cursor: null,
            };
          }

          return {
            results: [],
            has_more: false,
            next_cursor: null,
          };
        }),
      },
    };
    const pages = {
      create: vi.fn(async () => ({ id: "page-1" })),
      update: vi.fn(async () => undefined),
    };
    const client = new ApiNotionClient("token", {
      blocks,
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
          templatePageId: "template-page-id",
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
        children: expect.arrayContaining([
          expect.objectContaining({
            type: "heading_1",
            heading_1: expect.objectContaining({
              rich_text: [
                expect.objectContaining({
                  text: expect.objectContaining({
                    content: "Template Heading",
                  }),
                }),
              ],
            }),
          }),
        ]),
      }),
    );
    expect(blocks.children.list).toHaveBeenCalledWith({
      block_id: "template-page-id",
      start_cursor: undefined,
      page_size: 100,
    });
  });

  it("uses the default Notion data source template and defers generated meeting blocks until finalization", async () => {
    const blocks = {
      children: {
        append: vi.fn(async () => undefined),
        list: vi.fn(),
      },
    };
    const pages = {
      create: vi.fn(async () => ({ id: "page-1" })),
      update: vi.fn(async () => undefined),
    };
    const client = new ApiNotionClient("token", {
      blocks,
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
          dataSourceTemplate: {
            type: "default",
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
        meetingLink: "https://teams.microsoft.com/l/meetup-join/test",
        eventLink: "https://outlook.office.com/calendar/item/test",
        details: "Agenda and body content",
        isCancelled: false,
        isRecurring: false,
      },
    });

    expect(pages.create).toHaveBeenCalledWith(
      expect.objectContaining({
        parent: {
          data_source_id: "data-source-id",
        },
        template: {
          type: "default",
        },
        children: undefined,
      }),
    );
    expect(blocks.children.append).not.toHaveBeenCalled();
  });

  it("uses a specific Notion data source template id when configured", async () => {
    const blocks = {
      children: {
        append: vi.fn(async () => undefined),
        list: vi.fn(),
      },
    };
    const pages = {
      create: vi.fn(async () => ({ id: "page-1" })),
      update: vi.fn(async () => undefined),
    };
    const client = new ApiNotionClient("token", {
      blocks,
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
          dataSourceTemplate: {
            type: "template_id",
            templateId: "template-123",
            timezone: "America/New_York",
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
        template: {
          type: "template_id",
          template_id: "template-123",
          timezone: "America/New_York",
        },
      }),
    );
    expect(blocks.children.append).not.toHaveBeenCalled();
  });
});
