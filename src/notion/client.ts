import type { Meeting } from "../domain/meeting.js";
import type { NolendarConfig } from "../domain/config.js";
import type { NotionDataSourceSchema, NotionMeetingPage, NotionPageRecord, RequiredNotionProperty } from "../domain/notion.js";

export interface NotionClient {
  retrieveDataSource(dataSourceId: string): Promise<NotionDataSourceSchema>;
  getDefaultAssigneeUserId(defaultAssigneeEmail?: string): Promise<string | undefined>;
  getTemplateBlocks(templatePageId: string): Promise<unknown[]>;
  ensureProperties(dataSourceId: string, properties: RequiredNotionProperty[]): Promise<void>;
  findPageByEventId(args: {
    dataSourceId: string;
    eventIdPropertyName: string;
    changeKeyPropertyName: string;
    eventId: string;
  }): Promise<NotionPageRecord | null>;
  listMeetingPagesForWindow?(args: {
    dataSourceId: string;
    datePropertyName: string;
    start: string;
    end: string;
  }): Promise<NotionMeetingPage[]>;
  createMeetingPage(args: {
    config: NolendarConfig;
    dataSource: NotionDataSourceSchema;
    meeting: Meeting;
  }): Promise<{ id: string }>;
  updateMeetingPage(args: {
    pageId: string;
    config: NolendarConfig;
    dataSource: NotionDataSourceSchema;
    meeting: Meeting;
  }): Promise<void>;
  archivePage(pageId: string): Promise<void>;
  finalizeMeetingPageContent(args: {
    pageId: string;
    config: NolendarConfig;
    dataSource: NotionDataSourceSchema;
    meeting: Meeting;
  }): Promise<string>;
}
