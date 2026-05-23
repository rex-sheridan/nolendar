export type MicrosoftAuthMode = "device_code" | "auth_code";

export type LookaheadWindow = "today" | "24h" | "7d";

export interface MicrosoftConfig {
  tenant: "common" | "organizations" | "consumers";
  authMode: MicrosoftAuthMode;
}

export interface NotionConfig {
  databaseId: string;
  templatePageId?: string;
}

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
