export type NotionPropertyType =
  | "title"
  | "date"
  | "rich_text"
  | "email"
  | "url"
  | "status"
  | "multi_select"
  | "people"
  | "relation"
  | string;

export interface NotionDataSourceProperty {
  id: string;
  name: string;
  type: NotionPropertyType;
}

export interface NotionDataSourceSchema {
  id: string;
  title?: string;
  properties: Record<string, NotionDataSourceProperty>;
}

export interface NotionPageRecord {
  id: string;
  eventId: string;
  changeKey: string;
}

export interface NotionMeetingPage {
  id: string;
  url?: string;
  properties: Record<string, unknown>;
  body: string;
}

export interface NotionMeetingPageProperties {
  id: string;
  url?: string;
  properties: Record<string, unknown>;
}

export interface RequiredNotionProperty {
  name: string;
  type: "title" | "date" | "rich_text" | "email" | "url" | "status" | "multi_select" | "people" | "relation";
}
