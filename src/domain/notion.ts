export type NotionPropertyType = "title" | "date" | "rich_text" | "url" | "status" | string;

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

export interface RequiredNotionProperty {
  name: string;
  type: "title" | "date" | "rich_text";
}
