import { describe, expect, it, vi } from "vitest";

import type { NolendarConfig } from "../src/domain/config.js";
import type { Meeting } from "../src/domain/meeting.js";
import { ApiNotionClient } from "../src/notion/api-notion-client.js";

describe("ApiNotionClient.ensureProperties", () => {
  it("creates relation properties with the configured target data source", async () => {
    const dataSources = {
      retrieve: vi.fn(),
      update: vi.fn(async () => undefined),
      query: vi.fn(),
    };
    const client = new ApiNotionClient("token", {
      blocks: {
        children: {
          append: vi.fn(async () => undefined),
          list: vi.fn(),
        },
      },
      dataSources,
      pages: {
        create: vi.fn(),
        update: vi.fn(),
      },
      users: {
        me: vi.fn(),
        list: vi.fn(),
      },
    });

    await client.ensureProperties("meeting-id", [
      {
        name: "Participants",
        type: "relation",
        relationDataSourceId: "people-id",
      },
    ]);

    expect(dataSources.update).toHaveBeenCalledWith({
      data_source_id: "meeting-id",
      properties: {
        Participants: {
          relation: {
            data_source_id: "people-id",
            single_property: {},
          },
        },
      },
    });
  });
});

describe("ApiNotionClient.listDataSourceTemplates", () => {
  it("lists templates for a data source", async () => {
    const listTemplates = vi.fn(async () => ({
      templates: [
        { id: "template-2", name: "Review", is_default: false },
        { id: "template-1", name: "Default", is_default: true },
      ],
      has_more: false,
      next_cursor: null,
    }));
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
        listTemplates,
      },
      pages: {
        create: vi.fn(),
        update: vi.fn(),
      },
      users: {
        me: vi.fn(),
        list: vi.fn(),
      },
    });

    await expect(client.listDataSourceTemplates("meeting-id")).resolves.toEqual([
      { id: "template-1", name: "Default", isDefault: true },
      { id: "template-2", name: "Review", isDefault: false },
    ]);
    expect(listTemplates).toHaveBeenCalledWith({
      data_source_id: "meeting-id",
      page_size: 100,
      start_cursor: undefined,
    });
  });
});

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

describe("ApiNotionClient.setPageStatus", () => {
  it("updates a Notion status property by name", async () => {
    const pages = {
      create: vi.fn(),
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

    await client.setPageStatus({
      pageId: "page-1",
      propertyName: "Status",
      statusName: "Canceled",
    });

    expect(pages.update).toHaveBeenCalledWith({
      page_id: "page-1",
      properties: {
        Status: {
          status: {
            name: "Canceled",
          },
        },
      },
    });
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

describe("ApiNotionClient.listMeetingPagesForWindow", () => {
  it("queries pages by the meeting date property and renders readable page content", async () => {
    const dataSources = {
      retrieve: vi.fn(),
      update: vi.fn(),
      query: vi.fn(async () => ({
        results: [
          {
            object: "page",
            id: "page-1",
            url: "https://notion.so/page-1",
            properties: {
              Name: {
                type: "title",
                title: [{ plain_text: "Planning" }],
              },
              Due: {
                type: "date",
                date: {
                  start: "2026-05-22T13:00:00.000Z",
                  end: "2026-05-22T13:30:00.000Z",
                },
              },
              Status: {
                type: "status",
                status: {
                  name: "Done",
                },
              },
            },
          },
        ],
        has_more: false,
        next_cursor: null,
      })),
    };
    const blocks = {
      children: {
        append: vi.fn(async () => undefined),
        list: vi.fn(async () => ({
          results: [
            {
              type: "heading_2",
              heading_2: {
                rich_text: [{ plain_text: "Meeting Details" }],
              },
            },
            {
              type: "paragraph",
              paragraph: {
                rich_text: [{ plain_text: "Discuss launch readiness." }],
              },
            },
          ],
          has_more: false,
          next_cursor: null,
        })),
      },
    };
    const client = new ApiNotionClient("token", {
      blocks,
      dataSources,
      pages: {
        create: vi.fn(),
        update: vi.fn(),
      },
      users: {
        me: vi.fn(),
        list: vi.fn(),
      },
    });

    await expect(
      client.listMeetingPagesForWindow({
        dataSourceId: "meetings",
        datePropertyName: "Due",
        start: "2026-05-22T00:00:00.000Z",
        end: "2026-05-23T00:00:00.000Z",
      }),
    ).resolves.toEqual([
      {
        id: "page-1",
        url: "https://notion.so/page-1",
        properties: {
          Name: "Planning",
          Due: {
            start: "2026-05-22T13:00:00.000Z",
            end: "2026-05-22T13:30:00.000Z",
            timeZone: undefined,
          },
          Status: "Done",
        },
        body: "## Meeting Details\nDiscuss launch readiness.",
      },
    ]);
    expect(dataSources.query).toHaveBeenCalledWith(
      expect.objectContaining({
        data_source_id: "meetings",
        filter: {
          property: "Due",
          date: {
            on_or_after: "2026-05-22T00:00:00.000Z",
            before: "2026-05-23T00:00:00.000Z",
          },
        },
      }),
    );
    expect(blocks.children.list).toHaveBeenCalledWith(
      expect.objectContaining({
        block_id: "page-1",
      }),
    );
  });

  it("can query page properties without reading page blocks", async () => {
    const dataSources = {
      retrieve: vi.fn(),
      update: vi.fn(),
      query: vi.fn(async () => ({
        results: [
          {
            object: "page",
            id: "page-1",
            url: "https://notion.so/page-1",
            properties: {
              Name: {
                type: "title",
                title: [{ plain_text: "Planning" }],
              },
              "Outlook Event ID": {
                type: "rich_text",
                rich_text: [{ plain_text: "evt-1" }],
              },
            },
          },
        ],
        has_more: false,
        next_cursor: null,
      })),
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
    const timingReporter = {
      record: vi.fn(),
    };
    const client = new ApiNotionClient(
      "token",
      {
        blocks,
        dataSources,
        pages: {
          create: vi.fn(),
          update: vi.fn(),
        },
        users: {
          me: vi.fn(),
          list: vi.fn(),
        },
      },
      timingReporter,
    );

    await expect(
      client.listMeetingPagePropertiesForWindow({
        dataSourceId: "meetings",
        datePropertyName: "Due",
        start: "2026-05-22T00:00:00.000Z",
        end: "2026-05-23T00:00:00.000Z",
      }),
    ).resolves.toEqual([
      {
        id: "page-1",
        url: "https://notion.so/page-1",
        properties: {
          Name: "Planning",
          "Outlook Event ID": "evt-1",
        },
      },
    ]);
    expect(blocks.children.list).not.toHaveBeenCalled();
    expect(timingReporter.record).toHaveBeenCalledWith(
      expect.objectContaining({
        service: "notion",
        operation: "dataSources.query",
        count: 1,
      }),
    );
  });
});
