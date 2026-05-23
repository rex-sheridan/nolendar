import { Client } from "@notionhq/client";

import type { NolendarConfig } from "../domain/config.js";
import type { Meeting } from "../domain/meeting.js";
import type { NotionDataSourceProperty, NotionDataSourceSchema, NotionPageRecord, RequiredNotionProperty } from "../domain/notion.js";
import type { NotionClient } from "./client.js";
import { buildMeetingChildren, buildMeetingProperties } from "./page-payload.js";

const NOTION_API_VERSION = "2026-03-11";

interface NotionSdkClient {
  dataSources: {
    retrieve(args: { data_source_id: string }): Promise<unknown>;
    update(args: { data_source_id: string; properties: Record<string, unknown> }): Promise<unknown>;
    query(args: Record<string, unknown>): Promise<unknown>;
  };
  pages: {
    create(args: Record<string, unknown>): Promise<unknown>;
    update(args: Record<string, unknown>): Promise<unknown>;
  };
  users: {
    me(): Promise<unknown>;
    list(args?: { start_cursor?: string; page_size?: number }): Promise<unknown>;
  };
}

export class ApiNotionClient implements NotionClient {
  private readonly client: NotionSdkClient;
  private readonly defaultAssigneeCache = new Map<string, string | undefined>();

  constructor(authToken: string, client?: NotionSdkClient) {
    this.client =
      client ??
      (new Client({
        auth: authToken,
        notionVersion: NOTION_API_VERSION,
      }) as unknown as NotionSdkClient);
  }

  async retrieveDataSource(dataSourceId: string): Promise<NotionDataSourceSchema> {
    const response = (await this.client.dataSources.retrieve({
      data_source_id: dataSourceId,
    })) as {
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

    const response = (await this.client.users.me()) as {
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

  async ensureProperties(dataSourceId: string, properties: RequiredNotionProperty[]): Promise<void> {
    if (properties.length === 0) {
      return;
    }

    const payload: Record<string, unknown> = {};

    for (const property of properties) {
      payload[property.name] = propertyTypePayload(property.type);
    }

    await this.client.dataSources.update({
      data_source_id: dataSourceId,
      properties: payload,
    });
  }

  async findPageByEventId(args: {
    dataSourceId: string;
    eventIdPropertyName: string;
    changeKeyPropertyName: string;
    eventId: string;
  }): Promise<NotionPageRecord | null> {
    const response = (await this.client.dataSources.query({
      data_source_id: args.dataSourceId,
      result_type: "page",
      page_size: 1,
      filter: {
        property: args.eventIdPropertyName,
        rich_text: {
          equals: args.eventId,
        },
      },
    })) as {
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

  async createMeetingPage(args: {
    config: NolendarConfig;
    dataSource: NotionDataSourceSchema;
    meeting: Meeting;
  }): Promise<{ id: string }> {
    const assigneeUserId = args.config.mapping.assignee
      ? await this.getDefaultAssigneeUserId(args.config.notion.defaultAssigneeEmail)
      : undefined;
    const response = (await this.client.pages.create({
      parent: {
        data_source_id: args.dataSource.id,
      },
      properties: buildMeetingProperties(args.config, args.dataSource, args.meeting, {
        assigneeUserId,
      }),
      icon: buildPageIcon(args.config),
      children: buildMeetingChildren(args.meeting),
    })) as { id?: string };

    return {
      id: response.id ?? "",
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
    await this.client.pages.update({
      page_id: args.pageId,
      properties: buildMeetingProperties(args.config, args.dataSource, args.meeting, {
        assigneeUserId,
      }),
      icon: buildPageIcon(args.config),
    });
  }

  private async findUserIdByEmail(email: string): Promise<string | undefined> {
    const normalizedEmail = email.trim().toLowerCase();
    let nextCursor: string | undefined;

    do {
      const response = (await this.client.users.list({
        start_cursor: nextCursor,
        page_size: 100,
      })) as {
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

function propertyTypePayload(type: RequiredNotionProperty["type"]): Record<string, unknown> {
  switch (type) {
    case "title":
      return { title: {} };
    case "date":
      return { date: {} };
    case "rich_text":
      return { rich_text: {} };
    case "url":
      return { url: {} };
    case "multi_select":
      return { multi_select: { options: [] } };
    case "people":
      return { people: {} };
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
