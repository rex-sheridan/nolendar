export type MicrosoftAuthMode = "device_code" | "interactive_browser" | "auth_code";

export type LookaheadWindow = "today" | `${number}${"h" | "d" | "w" | "m"}`;

export interface MicrosoftConfig {
  tenant: "common" | "organizations" | "consumers";
  authMode: MicrosoftAuthMode;
}

export interface NotionConfig {
  databaseId: string;
  templatePageId?: string;
  dataSourceTemplate?: NotionDataSourceTemplateConfig;
  pageContent?: NotionPageContentConfig;
  defaultTags?: string[];
  defaultAssigneeEmail?: string;
  pageIcon?: NotionPageIconConfig;
  peopleDataSource?: NotionPeopleDataSourceConfig;
}

export interface NotionDataSourceTemplateConfig {
  type: "default" | "template_id";
  templateId?: string;
  timezone?: string;
}

export interface NotionPageContentConfig {
  sections: NotionPageContentSection[];
}

export type NotionPageContentSection =
  | "meeting_link"
  | "calendar_event"
  | "meeting_details"
  | "notes"
  | "action_items";

export interface NotionPeopleDataSourceConfig {
  databaseId: string;
  nameProperty: string;
  emailProperty: string;
  maxAttendeesPerMeeting?: number;
}

export type NotionPageIconConfig =
  | {
      type: "emoji";
      emoji: string;
    }
  | {
      type: "icon";
      name: string;
      color?: string;
    };

export interface CalendarConfig {
  id: string;
  name?: string;
}

export interface FiltersConfig {
  ignoreDeclined: boolean;
  minDurationMinutes?: number;
  requireAttendees: boolean;
  ignorePersonal: boolean;
  ignoreOptionalAttendance: boolean;
}

export interface MappingConfig {
  title: string;
  due: string;
  eventId: string;
  changeKey: string;
  eventLink?: string;
  tags?: string;
  assignee?: string;
  participants?: string;
}

export interface SyncConfig {
  lookahead: LookaheadWindow;
  statePath: string;
}

export interface NolendarConfig {
  microsoft: MicrosoftConfig;
  notion: NotionConfig;
  calendars: CalendarConfig[];
  filters: FiltersConfig;
  mapping: MappingConfig;
  sync: SyncConfig;
}
