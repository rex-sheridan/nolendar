import type { Meeting } from "../domain/meeting.js";
import type { NolendarConfig } from "../domain/config.js";
import type {
  NotionDataSourceSchema,
  NotionDataSourceSummary,
  NotionDataSourceTemplateSummary,
  NotionMeetingPage,
  NotionMeetingPageProperties,
  NotionPageRecord,
  RequiredNotionProperty,
} from "../domain/notion.js";

export interface NotionClient {
  listDataSources?(): Promise<NotionDataSourceSummary[]>;
  listDataSourceTemplates?(dataSourceId: string): Promise<NotionDataSourceTemplateSummary[]>;
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
  listMeetingPagePropertiesForWindow?(args: {
    dataSourceId: string;
    datePropertyName: string;
    start: string;
    end: string;
  }): Promise<NotionMeetingPageProperties[]>;
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
  setPageStatus?(args: {
    pageId: string;
    propertyName: string;
    statusName: string;
  }): Promise<void>;
  archivePage(pageId: string): Promise<void>;
  finalizeMeetingPageContent(args: {
    pageId: string;
    config: NolendarConfig;
    dataSource: NotionDataSourceSchema;
    meeting: Meeting;
  }): Promise<string>;
}
