import { readFile } from "node:fs/promises";
import path from "node:path";
import { parse } from "yaml";

import { isValidLookaheadWindow } from "./lookahead.js";
import type {
  CalendarConfig,
  FiltersConfig,
  LookaheadWindow,
  MappingConfig,
  MicrosoftAuthMode,
  NolendarConfig,
  NotionConfig,
  NotionPageContentSection,
  RelativeWindow,
} from "./domain/config.js";

export class ConfigError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ConfigError";
  }
}

export async function loadConfig(configPath: string): Promise<NolendarConfig> {
  const raw = await readFile(configPath, "utf8");
  const parsed = parse(raw) as unknown;

  return normalizeConfig(parsed, configPath);
}

export function normalizeConfig(input: unknown, configPath = process.cwd()): NolendarConfig {
  const record = asRecord(input, "Configuration file must contain a YAML object.");
  const notion = normalizeNotion(record.notion);
  const mapping = normalizeMapping(record.mapping);
  const calendars = normalizeCalendars(record.calendars);
  const lookahead = normalizeLookahead(record.sync);

  validateParticipantConfig(notion, mapping);
  validateTemplateConfig(notion);

  return {
    microsoft: {
      tenant: normalizeTenant(record.microsoft),
      authMode: normalizeMicrosoftAuthMode(record.microsoft),
    },
    notion,
    calendars,
    filters: normalizeFilters(record.filters),
    mapping,
    sync: {
      lookahead,
      statePath: normalizeStatePath(record.sync, configPath),
    },
  };
}

function normalizeTenant(value: unknown): "common" | "organizations" | "consumers" {
  if (value === undefined) {
    return "common";
  }

  const record = asRecord(value, "`microsoft` must be an object.");
  const tenant = record.tenant;

  if (tenant === undefined) {
    return "common";
  }

  if (tenant === "common" || tenant === "organizations" || tenant === "consumers") {
    return tenant;
  }

  throw new ConfigError("`microsoft.tenant` must be one of: common, organizations, consumers.");
}

function normalizeMicrosoftAuthMode(value: unknown): MicrosoftAuthMode {
  if (value === undefined) {
    return "device_code";
  }

  const record = asRecord(value, "`microsoft` must be an object.");
  const authMode = record.authMode;

  if (authMode === undefined) {
    return "device_code";
  }

  if (authMode === "device_code" || authMode === "interactive_browser" || authMode === "auth_code") {
    return authMode;
  }

  throw new ConfigError("`microsoft.authMode` must be one of: device_code, interactive_browser, auth_code.");
}

function normalizeNotion(value: unknown): NotionConfig {
  const record = asRecord(value, "`notion` must be an object.");
  const databaseId = requireString(record.databaseId, "`notion.databaseId` is required.");
  const templatePageId = optionalString(record.templatePageId, "`notion.templatePageId` must be a string.");
  const dataSourceTemplate = normalizeDataSourceTemplate(record.dataSourceTemplate);
  const pageContent = normalizeNotionPageContent(record.pageContent);
  const defaultTags = optionalStringArray(record.defaultTags, "`notion.defaultTags` must be an array of strings.");
  const defaultAssigneeEmail = optionalString(
    record.defaultAssigneeEmail,
    "`notion.defaultAssigneeEmail` must be a string.",
  );
  const pageIcon = normalizeNotionPageIcon(record.pageIcon);
  const peopleDataSource = normalizePeopleDataSource(record.peopleDataSource);
  const canceledMeetings = normalizeCanceledMeetings(record.canceledMeetings);
  const completedMeetings = normalizeCompletedMeetings(record.completedMeetings, canceledMeetings);

  return {
    databaseId,
    templatePageId,
    dataSourceTemplate,
    pageContent,
    defaultTags,
    defaultAssigneeEmail,
    pageIcon,
    peopleDataSource,
    canceledMeetings,
    completedMeetings,
  };
}

function normalizeCalendars(value: unknown): CalendarConfig[] {
  if (!Array.isArray(value) || value.length === 0) {
    throw new ConfigError("`calendars` must be a non-empty array.");
  }

  return value.map((entry, index) => {
    const record = asRecord(entry, `\`calendars[${index}]\` must be an object.`);

    return {
      id: requireString(record.id, `\`calendars[${index}].id\` is required.`),
      name: optionalString(record.name, `\`calendars[${index}].name\` must be a string.`),
    };
  });
}

function normalizeFilters(value: unknown): FiltersConfig {
  const record = value === undefined ? {} : asRecord(value, "`filters` must be an object.");
  const minDurationMinutes = optionalNumber(
    record.minDurationMinutes,
    "`filters.minDurationMinutes` must be a positive number.",
  );
  const ignoreNames = optionalStringArray(
    record.ignoreNames,
    "`filters.ignoreNames` must be an array of non-empty strings.",
  );
  const ignorePatterns = optionalStringArray(
    record.ignorePatterns ?? record.ignoreNamePatterns,
    "`filters.ignorePatterns` must be an array of non-empty strings.",
  );

  if (minDurationMinutes !== undefined && minDurationMinutes <= 0) {
    throw new ConfigError("`filters.minDurationMinutes` must be a positive number.");
  }

  validateRegexPatterns(ignorePatterns, "filters.ignorePatterns");

  return {
    ignoreDeclined: optionalBoolean(record.ignoreDeclined, true, "`filters.ignoreDeclined` must be a boolean."),
    minDurationMinutes,
    requireAttendees: optionalBoolean(record.requireAttendees, false, "`filters.requireAttendees` must be a boolean."),
    ignorePersonal: optionalBoolean(record.ignorePersonal, false, "`filters.ignorePersonal` must be a boolean."),
    ignoreOptionalAttendance: optionalBoolean(
      record.ignoreOptionalAttendance,
      false,
      "`filters.ignoreOptionalAttendance` must be a boolean.",
    ),
    ignoreNames: ignoreNames ?? [],
    ignorePatterns: ignorePatterns ?? [],
  };
}

function validateRegexPatterns(patterns: string[] | undefined, path: string): void {
  patterns?.forEach((pattern, index) => {
    try {
      new RegExp(pattern);
    } catch (error) {
      const detail = error instanceof Error ? ` ${error.message}` : "";
      throw new ConfigError(`\`${path}[${index}]\` must be a valid regular expression.${detail}`);
    }
  });
}

function normalizeMapping(value: unknown): MappingConfig {
  const record = value === undefined ? {} : asRecord(value, "`mapping` must be an object.");

  return {
    title: optionalStringWithDefault(record.title, "Name", "`mapping.title` must be a string."),
    due: optionalStringWithDefault(record.due, "Due", "`mapping.due` must be a string."),
    eventId: optionalStringWithDefault(record.eventId, "Outlook Event ID", "`mapping.eventId` must be a string."),
    changeKey: optionalStringWithDefault(record.changeKey, "Outlook ChangeKey", "`mapping.changeKey` must be a string."),
    eventLink: optionalString(record.eventLink, "`mapping.eventLink` must be a string."),
    tags: optionalString(record.tags, "`mapping.tags` must be a string."),
    assignee: optionalString(record.assignee, "`mapping.assignee` must be a string."),
    participants: optionalString(record.participants, "`mapping.participants` must be a string."),
  };
}

function normalizeLookahead(value: unknown): LookaheadWindow {
  if (value === undefined) {
    return "today";
  }

  const record = asRecord(value, "`sync` must be an object.");
  const lookahead = record.lookahead;

  if (lookahead === undefined) {
    return "today";
  }

  if (typeof lookahead !== "string" || !isValidLookaheadWindow(lookahead)) {
    throw new ConfigError("`sync.lookahead` must be `today` or a relative range like `12h`, `5d`, `2w`, or `3m`.");
  }

  return lookahead;
}

function normalizeStatePath(value: unknown, configPath: string): string {
  const record = value === undefined ? {} : asRecord(value, "`sync` must be an object.");
  const statePath = optionalString(record.statePath, "`sync.statePath` must be a string.");
  const baseDir = path.dirname(configPath);

  return path.resolve(baseDir, statePath ?? ".nolendar/state.json");
}

function asRecord(value: unknown, message: string): Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new ConfigError(message);
  }

  return value as Record<string, unknown>;
}

function requireString(value: unknown, message: string): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new ConfigError(message);
  }

  return value;
}

function optionalString(value: unknown, message: string): string | undefined {
  if (value === undefined) {
    return undefined;
  }

  if (typeof value !== "string" || value.trim() === "") {
    throw new ConfigError(message);
  }

  return value;
}

function optionalStringWithDefault(value: unknown, fallback: string, message: string): string {
  if (value === undefined) {
    return fallback;
  }

  if (typeof value !== "string" || value.trim() === "") {
    throw new ConfigError(message);
  }

  return value;
}

function optionalBoolean(value: unknown, fallback: boolean, message: string): boolean {
  if (value === undefined) {
    return fallback;
  }

  if (typeof value !== "boolean") {
    throw new ConfigError(message);
  }

  return value;
}

function optionalNumber(value: unknown, message: string): number | undefined {
  if (value === undefined) {
    return undefined;
  }

  if (typeof value !== "number" || Number.isNaN(value)) {
    throw new ConfigError(message);
  }

  return value;
}

function optionalNonNegativeIntegerWithDefault(value: unknown, fallback: number, message: string): number {
  if (value === undefined) {
    return fallback;
  }

  if (typeof value !== "number" || !Number.isInteger(value) || value < 0) {
    throw new ConfigError(message);
  }

  return value;
}

function optionalStringArray(value: unknown, message: string): string[] | undefined {
  if (value === undefined) {
    return undefined;
  }

  if (!Array.isArray(value) || value.some((entry) => typeof entry !== "string" || entry.trim() === "")) {
    throw new ConfigError(message);
  }

  return value;
}

function normalizeNotionPageIcon(value: unknown): NotionConfig["pageIcon"] {
  if (value === undefined) {
    return undefined;
  }

  const record = asRecord(value, "`notion.pageIcon` must be an object.");
  const type = requireString(record.type, "`notion.pageIcon.type` is required.");

  if (type === "emoji") {
    return {
      type,
      emoji: requireString(record.emoji, "`notion.pageIcon.emoji` is required for emoji icons."),
    };
  }

  if (type === "icon") {
    return {
      type,
      name: requireString(record.name, "`notion.pageIcon.name` is required for native icons."),
      color: optionalString(record.color, "`notion.pageIcon.color` must be a string."),
    };
  }

  throw new ConfigError("`notion.pageIcon.type` must be one of: emoji, icon.");
}

function normalizePeopleDataSource(value: unknown): NotionConfig["peopleDataSource"] {
  if (value === undefined) {
    return undefined;
  }

  const record = asRecord(value, "`notion.peopleDataSource` must be an object.");

  return {
    databaseId: requireString(record.databaseId, "`notion.peopleDataSource.databaseId` is required."),
    nameProperty: optionalStringWithDefault(
      record.nameProperty,
      "Name",
      "`notion.peopleDataSource.nameProperty` must be a string.",
    ),
    emailProperty: optionalStringWithDefault(
      record.emailProperty,
      "Email Address",
      "`notion.peopleDataSource.emailProperty` must be a string.",
    ),
    maxAttendeesPerMeeting: optionalNonNegativeIntegerWithDefault(
      record.maxAttendeesPerMeeting,
      10,
      "`notion.peopleDataSource.maxAttendeesPerMeeting` must be a non-negative integer.",
    ),
  };
}

function normalizeCanceledMeetings(value: unknown): NotionConfig["canceledMeetings"] {
  if (value === undefined) {
    return {
      action: "archive",
    };
  }

  const record = asRecord(value, "`notion.canceledMeetings` must be an object.");
  const action = requireString(record.action, "`notion.canceledMeetings.action` is required.");

  if (action === "archive") {
    return {
      action,
    };
  }

  if (action === "set_status") {
    return {
      action,
      statusProperty: requireString(
        record.statusProperty,
        "`notion.canceledMeetings.statusProperty` is required when action is `set_status`.",
      ),
      statusValue: requireString(
        record.statusValue,
        "`notion.canceledMeetings.statusValue` is required when action is `set_status`.",
      ),
    };
  }

  throw new ConfigError("`notion.canceledMeetings.action` must be one of: archive, set_status.");
}

function normalizeCompletedMeetings(
  value: unknown,
  canceledMeetings: NotionConfig["canceledMeetings"],
): NotionConfig["completedMeetings"] {
  if (value === undefined) {
    return canceledMeetings?.action === "set_status"
      ? {
          statusProperty: canceledMeetings.statusProperty,
          doneStatusValue: "Done",
          canceledStatusValue: canceledMeetings.statusValue,
          lookback: "1d",
        }
      : undefined;
  }

  const record = asRecord(value, "`notion.completedMeetings` must be an object.");

  return {
    statusProperty: requireString(
      record.statusProperty,
      "`notion.completedMeetings.statusProperty` is required.",
    ),
    doneStatusValue: optionalString(
      record.doneStatusValue,
      "`notion.completedMeetings.doneStatusValue` must be a string.",
    ) ?? "Done",
    canceledStatusValue: optionalString(
      record.canceledStatusValue,
      "`notion.completedMeetings.canceledStatusValue` must be a string.",
    ) ?? (canceledMeetings?.action === "set_status" ? canceledMeetings.statusValue : "Canceled"),
    lookback: normalizeRelativeWindow(
      record.lookback,
      "1d",
      "`notion.completedMeetings.lookback` must be a relative range like `12h`, `5d`, `2w`, or `3m`.",
    ),
  };
}

function normalizeRelativeWindow(value: unknown, fallback: RelativeWindow, message: string): RelativeWindow {
  if (value === undefined) {
    return fallback;
  }

  if (typeof value !== "string" || value === "today" || !isValidLookaheadWindow(value)) {
    throw new ConfigError(message);
  }

  return value as RelativeWindow;
}

function normalizeDataSourceTemplate(value: unknown): NotionConfig["dataSourceTemplate"] {
  if (value === undefined) {
    return undefined;
  }

  const record = asRecord(value, "`notion.dataSourceTemplate` must be an object.");
  const type = requireString(record.type, "`notion.dataSourceTemplate.type` is required.");

  if (type !== "default" && type !== "template_id") {
    throw new ConfigError("`notion.dataSourceTemplate.type` must be one of: default, template_id.");
  }

  const templateId = optionalString(
    record.templateId,
    "`notion.dataSourceTemplate.templateId` must be a string.",
  );

  if (type === "template_id" && !templateId) {
    throw new ConfigError("`notion.dataSourceTemplate.templateId` is required when template type is `template_id`.");
  }

  if (type === "default" && templateId !== undefined) {
    throw new ConfigError("`notion.dataSourceTemplate.templateId` must not be set when template type is `default`.");
  }

  return {
    type,
    templateId,
    timezone: optionalString(record.timezone, "`notion.dataSourceTemplate.timezone` must be a string."),
  };
}

function normalizeNotionPageContent(value: unknown): NotionConfig["pageContent"] {
  if (value === undefined) {
    return {
      sections: ["meeting_link", "calendar_event", "meeting_details", "notes", "action_items"],
    };
  }

  const record = asRecord(value, "`notion.pageContent` must be an object.");
  const sections = normalizeNotionPageContentSections(record.sections);

  return {
    sections,
  };
}

function normalizeNotionPageContentSections(value: unknown): NotionPageContentSection[] {
  if (!Array.isArray(value) || value.length === 0) {
    throw new ConfigError("`notion.pageContent.sections` must be a non-empty array.");
  }

  return value.map((entry, index) => {
    if (
      entry === "meeting_link" ||
      entry === "calendar_event" ||
      entry === "meeting_details" ||
      entry === "notes" ||
      entry === "action_items"
    ) {
      return entry;
    }

    throw new ConfigError(
      `\`notion.pageContent.sections[${index}]\` must be one of: meeting_link, calendar_event, meeting_details, notes, action_items.`,
    );
  });
}

function validateParticipantConfig(notion: NotionConfig, mapping: MappingConfig): void {
  const hasPeopleDataSource = notion.peopleDataSource !== undefined;
  const hasParticipantsMapping = mapping.participants !== undefined;

  if (hasPeopleDataSource === hasParticipantsMapping) {
    return;
  }

  if (hasPeopleDataSource) {
    throw new ConfigError("`mapping.participants` is required when `notion.peopleDataSource` is configured.");
  }

  throw new ConfigError("`notion.peopleDataSource` is required when `mapping.participants` is configured.");
}

function validateTemplateConfig(notion: NotionConfig): void {
  if (notion.templatePageId && notion.dataSourceTemplate) {
    throw new ConfigError("`notion.templatePageId` and `notion.dataSourceTemplate` are mutually exclusive.");
  }
}
