import { Client } from "@notionhq/client";

import type { ApiTimingReporter } from "../api-timing.js";
import type { NolendarConfig } from "../domain/config.js";
import type { Meeting } from "../domain/meeting.js";
import type {
  NotionDataSourceProperty,
  NotionDataSourceSchema,
  NotionDataSourceSummary,
  NotionDataSourceTemplateSummary,
  NotionMeetingPage,
  NotionMeetingPageProperties,
  NotionPageRecord,
  RequiredNotionProperty,
} from "../domain/notion.js";
import type { NotionClient } from "./client.js";
import { buildMeetingChildren, buildMeetingProperties } from "./page-payload.js";

const NOTION_API_VERSION = "2026-03-11";

interface NotionSdkClient {
  blocks: {
    children: {
      append(args: { block_id: string; children: unknown[] }): Promise<unknown>;
      list(args: { block_id: string; start_cursor?: string; page_size?: number }): Promise<unknown>;
    };
  };
  dataSources: {
    retrieve(args: { data_source_id: string }): Promise<unknown>;
    update(args: { data_source_id: string; properties: Record<string, unknown> }): Promise<unknown>;
    query(args: Record<string, unknown>): Promise<unknown>;
    listTemplates?(args: { data_source_id: string; start_cursor?: string; page_size?: number }): Promise<unknown>;
  };
  pages: {
    create(args: Record<string, unknown>): Promise<unknown>;
    update(args: Record<string, unknown>): Promise<unknown>;
  };
  users: {
    me(): Promise<unknown>;
    list(args?: { start_cursor?: string; page_size?: number }): Promise<unknown>;
  };
  search?(args: Record<string, unknown>): Promise<unknown>;
}

export class ApiNotionClient implements NotionClient {
  private readonly client: NotionSdkClient;
  private readonly timingReporter?: ApiTimingReporter;
  private readonly defaultAssigneeCache = new Map<string, string | undefined>();
  private readonly participantPageCache = new Map<string, string>();
  private readonly templateBlockCache = new Map<string, unknown[]>();

  constructor(authToken: string, client?: NotionSdkClient, timingReporter?: ApiTimingReporter) {
    this.client =
      client ??
      (new Client({
        auth: authToken,
        notionVersion: NOTION_API_VERSION,
      }) as unknown as NotionSdkClient);
    this.timingReporter = timingReporter;
  }

  async listDataSources(): Promise<NotionDataSourceSummary[]> {
    const search = this.client.search;

    if (!search) {
      return [];
    }

    const dataSources: NotionDataSourceSummary[] = [];
    let nextCursor: string | undefined;

    do {
      const response = (await this.timed(
        "search",
        `filter=data_source start_cursor=${nextCursor ?? "-"}`,
        () =>
        search({
          filter: {
            property: "object",
            value: "data_source",
          },
          page_size: 100,
          start_cursor: nextCursor,
        }),
      )) as {
        results?: Array<{
          id?: string;
          object?: string;
          title?: Array<{ plain_text?: string }>;
        }>;
        next_cursor?: string | null;
        has_more?: boolean;
      };

      dataSources.push(
        ...(response.results ?? [])
          .filter((entry) => entry.object === "data_source" && typeof entry.id === "string")
          .map((entry) => ({
            id: entry.id ?? "",
            title: entry.title?.map((title) => title.plain_text ?? "").join("") || undefined,
          })),
      );
      nextCursor = response.has_more ? (response.next_cursor ?? undefined) : undefined;
    } while (nextCursor);

    return dataSources.sort((left, right) => (left.title ?? left.id).localeCompare(right.title ?? right.id));
  }

  async retrieveDataSource(dataSourceId: string): Promise<NotionDataSourceSchema> {
    const response = (await this.timed("dataSources.retrieve", `data_source_id=${dataSourceId}`, () =>
      this.client.dataSources.retrieve({
        data_source_id: dataSourceId,
      }),
    )) as {
      id?: string;
      title?: Array<{ plain_text?: string }>;
      properties?: Record<string, { id?: string; name?: string; type?: string }>;
    };

    return {
      id: response.id ?? dataSourceId,
      title: response.title?.map((entry) => entry.plain_text ?? "").join(""),
      properties: normalizeProperties(response.properties),
    };
  }

  async listDataSourceTemplates(dataSourceId: string): Promise<NotionDataSourceTemplateSummary[]> {
    const listTemplates = this.client.dataSources.listTemplates;

    if (!listTemplates) {
      return [];
    }

    const templates: NotionDataSourceTemplateSummary[] = [];
    let nextCursor: string | undefined;

    do {
      const response = (await this.timed(
        "dataSources.listTemplates",
        `data_source_id=${dataSourceId} start_cursor=${nextCursor ?? "-"}`,
        () =>
        listTemplates({
          data_source_id: dataSourceId,
          page_size: 100,
          start_cursor: nextCursor,
        }),
      )) as {
        templates?: Array<{
          id?: string;
          name?: string;
          is_default?: boolean;
        }>;
        next_cursor?: string | null;
        has_more?: boolean;
      };

      templates.push(
        ...(response.templates ?? [])
          .filter((template) => typeof template.id === "string")
          .map((template) => ({
            id: template.id ?? "",
            name: template.name || template.id || "",
            isDefault: template.is_default ?? false,
          })),
      );
      nextCursor = response.has_more ? (response.next_cursor ?? undefined) : undefined;
    } while (nextCursor);

    return templates.sort(compareDataSourceTemplates);
  }

  async getDefaultAssigneeUserId(defaultAssigneeEmail?: string): Promise<string | undefined> {
    const cacheKey = defaultAssigneeEmail?.trim().toLowerCase() ?? "__self__";

    if (this.defaultAssigneeCache.has(cacheKey)) {
      return this.defaultAssigneeCache.get(cacheKey);
    }

    if (defaultAssigneeEmail) {
      const matchedUserId = await this.findUserIdByEmail(defaultAssigneeEmail);
      if (matchedUserId) {
        this.defaultAssigneeCache.set(cacheKey, matchedUserId);
        return matchedUserId;
      }
    }

    const response = (await this.timed("users.me", undefined, () => this.client.users.me())) as {
      id?: string;
      type?: string;
      bot?: {
        owner?: {
          type?: string;
          user?: {
            id?: string;
          };
        };
      };
    };

    const resolvedUserId =
      response.type === "person"
        ? response.id
        : response.bot?.owner?.type === "user"
          ? response.bot.owner.user?.id
          : undefined;

    this.defaultAssigneeCache.set(cacheKey, resolvedUserId);
    return resolvedUserId;
  }

  async getTemplateBlocks(templatePageId: string): Promise<unknown[]> {
    if (this.templateBlockCache.has(templatePageId)) {
      return this.templateBlockCache.get(templatePageId) ?? [];
    }

    const blocks = await this.listBlockChildren(templatePageId, {
      purpose: "template_copy",
    });
    const sanitizedBlocks = await Promise.all(blocks.map((block) => this.sanitizeBlockForCreate(block)));
    this.templateBlockCache.set(templatePageId, sanitizedBlocks);
    return sanitizedBlocks;
  }

  async ensureProperties(dataSourceId: string, properties: RequiredNotionProperty[]): Promise<void> {
    if (properties.length === 0) {
      return;
    }

    const payload: Record<string, unknown> = {};

    for (const property of properties) {
      payload[property.name] = propertyTypePayload(property);
    }

    await this.timed(
      "dataSources.update",
      `data_source_id=${dataSourceId} properties=${Object.keys(payload).sort().join(",")}`,
      () =>
      this.client.dataSources.update({
        data_source_id: dataSourceId,
        properties: payload,
      }),
    );
  }

  async findPageByEventId(args: {
    dataSourceId: string;
    eventIdPropertyName: string;
    changeKeyPropertyName: string;
    eventId: string;
  }): Promise<NotionPageRecord | null> {
    const response = (await this.timed(
      "dataSources.query",
      `data_source_id=${args.dataSourceId} filter=${args.eventIdPropertyName}:rich_text=${args.eventId}`,
      () =>
      this.client.dataSources.query({
        data_source_id: args.dataSourceId,
        result_type: "page",
        page_size: 1,
        filter: {
          property: args.eventIdPropertyName,
          rich_text: {
            equals: args.eventId,
          },
        },
      }),
    )) as {
      results?: Array<{
        id?: string;
        object?: string;
        properties?: Record<string, unknown>;
      }>;
    };

    const first = response.results?.find((entry) => entry.object === "page");

    if (!first?.id) {
      return null;
    }

    return {
      id: first.id,
      eventId: readRichTextProperty(first.properties?.[args.eventIdPropertyName]),
      changeKey: readRichTextProperty(first.properties?.[args.changeKeyPropertyName]),
    };
  }

  async listMeetingPagesForWindow(args: {
    dataSourceId: string;
    datePropertyName: string;
    start: string;
    end: string;
  }): Promise<NotionMeetingPage[]> {
    const pages = await this.listMeetingPagePropertiesForWindow(args);

    return Promise.all(
      pages.map(async (page) => ({
        ...page,
        body: await this.getPageBodyMarkdown(page.id),
      })),
    );
  }

  async listMeetingPagePropertiesForWindow(args: {
    dataSourceId: string;
    datePropertyName: string;
    start: string;
    end: string;
  }): Promise<NotionMeetingPageProperties[]> {
    const pages: Array<{
      id?: string;
      object?: string;
      url?: string;
      properties?: Record<string, unknown>;
    }> = [];
    let nextCursor: string | undefined;

    do {
      const response = (await this.timed(
        "dataSources.query",
        `data_source_id=${args.dataSourceId} filter=${args.datePropertyName}:date ${args.start}..${args.end}`,
        () =>
          this.client.dataSources.query({
            data_source_id: args.dataSourceId,
            result_type: "page",
            page_size: 100,
            start_cursor: nextCursor,
            filter: {
              property: args.datePropertyName,
              date: {
                on_or_after: args.start,
                before: args.end,
              },
            },
            sorts: [
              {
                property: args.datePropertyName,
                direction: "ascending",
              },
            ],
          }),
      )) as {
        results?: Array<{
          id?: string;
          object?: string;
          url?: string;
          properties?: Record<string, unknown>;
        }>;
        next_cursor?: string | null;
        has_more?: boolean;
      };

      pages.push(...(response.results ?? []).filter((entry) => entry.object === "page"));
      nextCursor = response.has_more ? (response.next_cursor ?? undefined) : undefined;
    } while (nextCursor);

    return pages.map((page) => ({
      id: page.id ?? "",
      url: page.url,
      properties: normalizePageProperties(page.properties ?? {}),
    }));
  }

  async createMeetingPage(args: {
    config: NolendarConfig;
    dataSource: NotionDataSourceSchema;
    meeting: Meeting;
  }): Promise<{ id: string }> {
    const assigneeUserId = args.config.mapping.assignee
      ? await this.getDefaultAssigneeUserId(args.config.notion.defaultAssigneeEmail)
      : undefined;
    const participantPageIds = await this.resolveParticipantPageIds(args.config, args.meeting);
    const properties = buildMeetingProperties(args.config, args.dataSource, args.meeting, {
      assigneeUserId,
      participantPageIds,
    });
    const meetingChildren = buildMeetingChildren(args.config, args.meeting);
    const templateBlocks = args.config.notion.templatePageId
      ? await this.getTemplateBlocks(args.config.notion.templatePageId)
      : [];
    const dataSourceTemplate = buildDataSourceTemplatePayload(args.config);
    const response = (await this.timed(
      "pages.create",
      describeCreateMeetingPage(args.dataSource.id, dataSourceTemplate, templateBlocks.length, meetingChildren.length),
      () =>
      this.client.pages.create({
        parent: {
          data_source_id: args.dataSource.id,
        },
        properties,
        icon: buildPageIcon(args.config),
        template: dataSourceTemplate,
        children: dataSourceTemplate ? undefined : [...templateBlocks, ...meetingChildren],
      }),
    )) as { id?: string };
    const pageId = response.id ?? "";

    return {
      id: pageId,
    };
  }

  async updateMeetingPage(args: {
    pageId: string;
    config: NolendarConfig;
    dataSource: NotionDataSourceSchema;
    meeting: Meeting;
  }): Promise<void> {
    const assigneeUserId = args.config.mapping.assignee
      ? await this.getDefaultAssigneeUserId(args.config.notion.defaultAssigneeEmail)
      : undefined;
    const participantPageIds = await this.resolveParticipantPageIds(args.config, args.meeting);
    await this.timed(
      "pages.update",
      `page_id=${args.pageId} properties=${Object.keys(buildMeetingProperties(args.config, args.dataSource, args.meeting, {
        assigneeUserId,
        participantPageIds,
      }))
        .sort()
        .join(",")}`,
      () =>
      this.client.pages.update({
        page_id: args.pageId,
        properties: buildMeetingProperties(args.config, args.dataSource, args.meeting, {
          assigneeUserId,
          participantPageIds,
        }),
        icon: buildPageIcon(args.config),
      }),
    );
  }

  async archivePage(pageId: string): Promise<void> {
    await this.timed("pages.update", `page_id=${pageId} archived=true`, () =>
      this.client.pages.update({
        page_id: pageId,
        archived: true,
      }),
    );
  }

  async setPageStatus(args: {
    pageId: string;
    propertyName: string;
    statusName: string;
  }): Promise<void> {
    await this.timed("pages.update", `page_id=${args.pageId} status=${args.propertyName}:${args.statusName}`, () =>
      this.client.pages.update({
        page_id: args.pageId,
        properties: {
          [args.propertyName]: {
            status: {
              name: args.statusName,
            },
          },
        },
      }),
    );
  }

  async finalizeMeetingPageContent(args: {
    pageId: string;
    config: NolendarConfig;
    dataSource: NotionDataSourceSchema;
    meeting: Meeting;
  }): Promise<"appended" | "marked_existing"> {
    const meetingChildren = buildMeetingChildren(args.config, args.meeting);

    if (await this.pageHasGeneratedMeetingContent(args.pageId, args.config, args.meeting)) {
      return "marked_existing";
    }

    await this.timed("blocks.children.append", `block_id=${args.pageId} children=${meetingChildren.length}`, () =>
      this.client.blocks.children.append({
        block_id: args.pageId,
        children: meetingChildren,
      }),
    );
    return "appended";
  }

  private async findUserIdByEmail(email: string): Promise<string | undefined> {
    const normalizedEmail = email.trim().toLowerCase();
    let nextCursor: string | undefined;

    do {
      const response = (await this.timed(
        "users.list",
        `start_cursor=${nextCursor ?? "-"} page_size=100`,
        () =>
        this.client.users.list({
          start_cursor: nextCursor,
          page_size: 100,
        }),
      )) as {
        results?: Array<{
          id?: string;
          type?: string;
          person?: {
            email?: string;
          };
        }>;
        next_cursor?: string | null;
        has_more?: boolean;
      };

      const match = response.results?.find(
        (user) => user.type === "person" && user.person?.email?.trim().toLowerCase() === normalizedEmail,
      );

      if (match?.id) {
        return match.id;
      }

      nextCursor = response.has_more ? (response.next_cursor ?? undefined) : undefined;
    } while (nextCursor);

    return undefined;
  }

  private async resolveParticipantPageIds(config: NolendarConfig, meeting: Meeting): Promise<string[]> {
    const peopleDataSource = config.notion.peopleDataSource;
    const participantsProperty = config.mapping.participants;

    if (!peopleDataSource || !participantsProperty) {
      return [];
    }

    const resolvedPageIds = new Set<string>();
    const maxAttendeesPerMeeting = peopleDataSource.maxAttendeesPerMeeting ?? 10;

    if (maxAttendeesPerMeeting === 0) {
      return [];
    }

    for (const attendee of meeting.attendees) {
      if (resolvedPageIds.size >= maxAttendeesPerMeeting) {
        break;
      }

      const email = attendee.email?.trim().toLowerCase();

      if (!email) {
        continue;
      }

      const cachedPageId = this.participantPageCache.get(email);
      if (cachedPageId) {
        resolvedPageIds.add(cachedPageId);
        continue;
      }

      const pageId =
        (await this.findPeoplePageByEmail(peopleDataSource.databaseId, peopleDataSource.emailProperty, email)) ??
        (await this.createPeoplePage(peopleDataSource.databaseId, peopleDataSource.nameProperty, peopleDataSource.emailProperty, {
          name: attendee.name?.trim(),
          email,
        }));

      this.participantPageCache.set(email, pageId);
      resolvedPageIds.add(pageId);
    }

    return Array.from(resolvedPageIds);
  }

  private async findPeoplePageByEmail(
    dataSourceId: string,
    emailPropertyName: string,
    email: string,
  ): Promise<string | undefined> {
    const response = (await this.timed(
      "dataSources.query",
      `data_source_id=${dataSourceId} filter=${emailPropertyName}:email=${email}`,
      () =>
      this.client.dataSources.query({
        data_source_id: dataSourceId,
        result_type: "page",
        page_size: 1,
        filter: {
          property: emailPropertyName,
          email: {
            equals: email,
          },
        },
      }),
    )) as {
      results?: Array<{
        id?: string;
        object?: string;
      }>;
    };

    return response.results?.find((entry) => entry.object === "page")?.id;
  }

  private async createPeoplePage(
    dataSourceId: string,
    namePropertyName: string,
    emailPropertyName: string,
    attendee: { name?: string; email: string },
  ): Promise<string> {
    const response = (await this.timed(
      "pages.create",
      `parent_data_source_id=${dataSourceId} title=${truncateForTiming(attendee.name || attendee.email)} email=${attendee.email}`,
      () =>
      this.client.pages.create({
        parent: {
          data_source_id: dataSourceId,
        },
        properties: {
          [namePropertyName]: {
            title: [textBlock(attendee.name || attendee.email)],
          },
          [emailPropertyName]: {
            email: attendee.email,
          },
        },
      }),
    )) as { id?: string };

    return response.id ?? "";
  }

  private async listBlockChildren(
    blockId: string,
    options: {
      purpose?: "template_copy" | "nested_block_copy";
    } = {},
  ): Promise<Array<Record<string, unknown>>> {
    const blocks: Array<Record<string, unknown>> = [];
    let nextCursor: string | undefined;

    do {
      const response = (await this.timed(
        "blocks.children.list",
        describeBlockChildrenList(blockId, nextCursor, options),
        () =>
        this.client.blocks.children.list({
          block_id: blockId,
          start_cursor: nextCursor,
          page_size: 100,
        }),
      )) as {
        results?: Array<Record<string, unknown>>;
        next_cursor?: string | null;
        has_more?: boolean;
      };

      blocks.push(...(response.results ?? []));
      nextCursor = response.has_more ? (response.next_cursor ?? undefined) : undefined;
    } while (nextCursor);

    return blocks;
  }

  private async sanitizeBlockForCreate(block: Record<string, unknown>): Promise<Record<string, unknown>> {
    const sanitized: Record<string, unknown> = {};

    for (const [key, value] of Object.entries(block)) {
      if (
        key === "id" ||
        key === "parent" ||
        key === "created_time" ||
        key === "last_edited_time" ||
        key === "created_by" ||
        key === "last_edited_by" ||
        key === "archived" ||
        key === "in_trash" ||
        key === "has_children"
      ) {
        continue;
      }

      sanitized[key] = value;
    }

    const blockType = typeof sanitized.type === "string" ? sanitized.type : undefined;
    if (!blockType) {
      return sanitized;
    }

    const typedPayload = sanitized[blockType];
    if (!typedPayload || typeof typedPayload !== "object" || Array.isArray(typedPayload)) {
      return sanitized;
    }

    const typedRecord = { ...(typedPayload as Record<string, unknown>) };

    if ((block.type as string | undefined) && (block as { has_children?: boolean }).has_children) {
      typedRecord.children = await this.listAndSanitizeNestedChildren(block.id as string);
    }

    sanitized[blockType] = typedRecord;
    return sanitized;
  }

  private async listAndSanitizeNestedChildren(blockId: string): Promise<unknown[]> {
    const childBlocks = await this.listBlockChildren(blockId, {
      purpose: "nested_block_copy",
    });
    return Promise.all(childBlocks.map((child) => this.sanitizeBlockForCreate(child)));
  }

  private async getPageBodyMarkdown(pageId: string): Promise<string> {
    if (!pageId) {
      return "";
    }

    const blocks = await this.listBlockChildren(pageId);
    const lines: string[] = [];

    for (const block of blocks) {
      lines.push(...(await this.renderBlockMarkdown(block)));
    }

    return lines.join("\n").replace(/\n{3,}/g, "\n\n").trim();
  }

  private async renderBlockMarkdown(block: Record<string, unknown>, indent = ""): Promise<string[]> {
    const type = typeof block.type === "string" ? block.type : undefined;
    const payload = type ? block[type] : undefined;
    const record = payload && typeof payload === "object" && !Array.isArray(payload) ? (payload as Record<string, unknown>) : {};
    const text = richTextToPlainText(record.rich_text);
    const lines: string[] = [];

    switch (type) {
      case "heading_1":
        lines.push(`# ${text}`);
        break;
      case "heading_2":
        lines.push(`## ${text}`);
        break;
      case "heading_3":
        lines.push(`### ${text}`);
        break;
      case "bulleted_list_item":
        lines.push(`${indent}- ${text}`);
        break;
      case "numbered_list_item":
        lines.push(`${indent}1. ${text}`);
        break;
      case "to_do":
        lines.push(`${indent}- [${record.checked ? "x" : " "}] ${text}`);
        break;
      case "quote":
        lines.push(`> ${text}`);
        break;
      case "code":
        lines.push("```", text, "```");
        break;
      case "divider":
        lines.push("---");
        break;
      case "child_page":
        lines.push(`[child page] ${typeof record.title === "string" ? record.title : ""}`.trim());
        break;
      case "paragraph":
        if (text) {
          lines.push(text);
        }
        break;
      default:
        if (text) {
          lines.push(text);
        }
        break;
    }

    if ((block as { has_children?: boolean }).has_children && typeof block.id === "string") {
      const children = await this.listBlockChildren(block.id);
      for (const child of children) {
        lines.push(...(await this.renderBlockMarkdown(child, `${indent}  `)));
      }
    }

    return lines;
  }

  private async pageHasGeneratedMeetingContent(pageId: string, config: NolendarConfig, meeting: Meeting): Promise<boolean> {
    const children = await this.listBlockChildren(pageId);
    const headings = new Set(
      children
        .map((child) => extractBlockHeading(child))
        .filter((value): value is string => Boolean(value)),
    );
    const sections = config.notion.pageContent?.sections ?? [
      "meeting_link",
      "calendar_event",
      "meeting_details",
      "notes",
      "action_items",
    ];

    for (const section of sections) {
      switch (section) {
        case "meeting_link":
          if (meeting.meetingLink && !headings.has("Meeting Link")) {
            return false;
          }
          break;
        case "calendar_event":
          if (meeting.eventLink && !headings.has("Calendar Event")) {
            return false;
          }
          break;
        case "meeting_details":
          if ((meeting.details || meeting.agenda) && !headings.has("Meeting Details")) {
            return false;
          }
          break;
        case "notes":
          if (!headings.has("Notes")) {
            return false;
          }
          break;
        case "action_items":
          if (!headings.has("Action items")) {
            return false;
          }
          break;
      }
    }

    return true;
  }

  private async timed<T>(operation: string, detail: string | undefined, fn: () => Promise<T>): Promise<T> {
    const startedAt = Date.now();

    try {
      const result = await fn();
      this.timingReporter?.record({
        service: "notion",
        operation,
        detail,
        status: "ok",
        count: getNotionResultCount(result),
        durationMs: Date.now() - startedAt,
      });
      return result;
    } catch (error) {
      this.timingReporter?.record({
        service: "notion",
        operation,
        detail,
        status: "error",
        durationMs: Date.now() - startedAt,
      });
      throw error;
    }
  }
}

function getNotionResultCount(result: unknown): number | undefined {
  if (!result || typeof result !== "object" || Array.isArray(result)) {
    return undefined;
  }

  const results = (result as { results?: unknown }).results;

  if (Array.isArray(results)) {
    return results.length;
  }

  return undefined;
}

function describeCreateMeetingPage(
  dataSourceId: string,
  dataSourceTemplate: Record<string, unknown> | undefined,
  templateBlockCount: number,
  meetingChildCount: number,
): string {
  if (dataSourceTemplate) {
    return `parent_data_source_id=${dataSourceId} template=${String(dataSourceTemplate.type)} generated_children=${meetingChildCount}`;
  }

  return `parent_data_source_id=${dataSourceId} template_blocks=${templateBlockCount} children=${templateBlockCount + meetingChildCount}`;
}

function describeBlockChildrenList(
  blockId: string,
  nextCursor: string | undefined,
  options: {
    purpose?: "template_copy" | "nested_block_copy";
  },
): string {
  const parts = [`block_id=${blockId}`];

  if (options.purpose) {
    parts.push(`purpose=${options.purpose}`);
  }

  parts.push(`start_cursor=${nextCursor ?? "-"}`);
  parts.push("page_size=100");

  return parts.join(" ");
}

function truncateForTiming(value: string, maxLength = 40): string {
  return value.length <= maxLength ? value : `${value.slice(0, maxLength - 3)}...`;
}

function extractBlockHeading(block: Record<string, unknown>): string | undefined {
  const type = typeof block.type === "string" ? block.type : undefined;

  if (type !== "heading_1" && type !== "heading_2" && type !== "heading_3") {
    return undefined;
  }

  const payload = block[type] as { rich_text?: Array<{ plain_text?: string; text?: { content?: string } }> } | undefined;
  const richText = payload?.rich_text;

  if (!Array.isArray(richText)) {
    return undefined;
  }

  return richText
    .map((entry) => entry.plain_text ?? entry.text?.content ?? "")
    .join("")
    .trim() || undefined;
}

function normalizeProperties(
  properties?: Record<string, { id?: string; name?: string; type?: string }>,
): Record<string, NotionDataSourceProperty> {
  const normalized: Record<string, NotionDataSourceProperty> = {};

  for (const [name, property] of Object.entries(properties ?? {})) {
    normalized[name] = {
      id: property.id ?? name,
      name: property.name ?? name,
      type: property.type ?? "unknown",
    };
  }

  return normalized;
}

function normalizePageProperties(properties: Record<string, unknown>): Record<string, unknown> {
  const normalized: Record<string, unknown> = {};

  for (const [name, property] of Object.entries(properties)) {
    normalized[name] = normalizePageProperty(property);
  }

  return normalized;
}

function normalizePageProperty(property: unknown): unknown {
  const record = property && typeof property === "object" && !Array.isArray(property) ? (property as Record<string, unknown>) : {};
  const type = typeof record.type === "string" ? record.type : undefined;

  switch (type) {
    case "title":
      return richTextToPlainText(record.title);
    case "rich_text":
      return richTextToPlainText(record.rich_text);
    case "date": {
      const date = record.date && typeof record.date === "object" ? (record.date as Record<string, unknown>) : undefined;
      return date
        ? {
            start: date.start,
            end: date.end,
            timeZone: date.time_zone,
          }
        : null;
    }
    case "email":
    case "url":
    case "phone_number":
    case "number":
    case "checkbox":
      return record[type];
    case "status": {
      const status = record.status && typeof record.status === "object" ? (record.status as Record<string, unknown>) : undefined;
      return status?.name ?? null;
    }
    case "select": {
      const select = record.select && typeof record.select === "object" ? (record.select as Record<string, unknown>) : undefined;
      return select?.name ?? null;
    }
    case "multi_select":
      return Array.isArray(record.multi_select)
        ? record.multi_select.map((entry) =>
            entry && typeof entry === "object" ? (entry as { name?: string }).name : undefined,
          ).filter(Boolean)
        : [];
    case "people":
      return Array.isArray(record.people)
        ? record.people.map((entry) => normalizePerson(entry)).filter((entry) => Object.keys(entry).length > 0)
        : [];
    case "relation":
      return Array.isArray(record.relation)
        ? record.relation.map((entry) => (entry && typeof entry === "object" ? (entry as { id?: string }).id : undefined)).filter(Boolean)
        : [];
    default:
      return record[type ?? ""] ?? null;
  }
}

function normalizePerson(value: unknown): Record<string, string> {
  const record = value && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : {};
  const person = record.person && typeof record.person === "object" ? (record.person as Record<string, unknown>) : {};
  const normalized: Record<string, string> = {};

  if (typeof record.id === "string") {
    normalized.id = record.id;
  }

  if (typeof record.name === "string") {
    normalized.name = record.name;
  }

  if (typeof person.email === "string") {
    normalized.email = person.email;
  }

  return normalized;
}

function richTextToPlainText(value: unknown): string {
  if (!Array.isArray(value)) {
    return "";
  }

  return value
    .map((entry) => {
      const record = entry && typeof entry === "object" ? (entry as Record<string, unknown>) : {};
      const plainText = record.plain_text;
      const text = record.text && typeof record.text === "object" ? (record.text as Record<string, unknown>) : undefined;
      const content = text?.content;
      const href = typeof record.href === "string" ? record.href : undefined;
      const label = typeof plainText === "string" ? plainText : typeof content === "string" ? content : "";

      return href && label ? `${label} (${href})` : label;
    })
    .join("")
    .trim();
}

function compareDataSourceTemplates(
  left: NotionDataSourceTemplateSummary,
  right: NotionDataSourceTemplateSummary,
): number {
  if (left.isDefault !== right.isDefault) {
    return left.isDefault ? -1 : 1;
  }

  return left.name.localeCompare(right.name);
}

function propertyTypePayload(property: RequiredNotionProperty): Record<string, unknown> {
  switch (property.type) {
    case "title":
      return { title: {} };
    case "date":
      return { date: {} };
    case "rich_text":
      return { rich_text: {} };
    case "email":
      return { email: {} };
    case "url":
      return { url: {} };
    case "status":
      return { status: {} };
    case "multi_select":
      return { multi_select: { options: [] } };
    case "people":
      return { people: {} };
    case "relation":
      if (!property.relationDataSourceId) {
        throw new Error("Relation properties require a target data source ID.");
      }

      return {
        relation: {
          data_source_id: property.relationDataSourceId,
          single_property: {},
        },
      };
  }
}

function readRichTextProperty(property: unknown): string {
  const richText = (property as { type?: string; rich_text?: Array<{ plain_text?: string }> } | undefined)?.rich_text;

  if (!Array.isArray(richText)) {
    return "";
  }

  return richText.map((entry) => entry.plain_text ?? "").join("");
}

function buildPageIcon(config: NolendarConfig): Record<string, unknown> | undefined {
  const pageIcon = config.notion.pageIcon;

  if (!pageIcon) {
    return undefined;
  }

  if (pageIcon.type === "emoji") {
    return {
      type: "emoji",
      emoji: pageIcon.emoji,
    };
  }

  return {
    type: "icon",
    icon: {
      name: pageIcon.name,
      color: pageIcon.color ?? "gray",
    },
  };
}

function textBlock(content: string): { type: "text"; text: { content: string } } {
  return {
    type: "text",
    text: {
      content,
    },
  };
}

function buildDataSourceTemplatePayload(config: NolendarConfig): Record<string, unknown> | undefined {
  const template = config.notion.dataSourceTemplate;

  if (!template) {
    return undefined;
  }

  if (template.type === "default") {
    return {
      type: "default",
      ...(template.timezone ? { timezone: template.timezone } : {}),
    };
  }

  return {
    type: "template_id",
    template_id: template.templateId,
    ...(template.timezone ? { timezone: template.timezone } : {}),
  };
}
