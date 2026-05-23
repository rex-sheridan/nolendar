import type { Meeting } from "../domain/meeting.js";
import type { NolendarConfig } from "../domain/config.js";
import type { NotionDataSourceSchema, NotionPageRecord, RequiredNotionProperty } from "../domain/notion.js";

export interface NotionClient {
  retrieveDataSource(dataSourceId: string): Promise<NotionDataSourceSchema>;
  ensureProperties(dataSourceId: string, properties: RequiredNotionProperty[]): Promise<void>;
  findPageByEventId(args: {
    dataSourceId: string;
    eventIdPropertyName: string;
    changeKeyPropertyName: string;
    eventId: string;
  }): Promise<NotionPageRecord | null>;
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
}
