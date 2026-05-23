import { readFile } from "node:fs/promises";
import path from "node:path";
import { parse } from "yaml";

import type {
  CalendarConfig,
  FiltersConfig,
  LookaheadWindow,
  MappingConfig,
  NolendarConfig,
  NotionConfig,
} from "./domain/config.js";

const VALID_LOOKAHEADS: ReadonlySet<LookaheadWindow> = new Set(["today", "24h", "7d"]);

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
  const calendars = normalizeCalendars(record.calendars);
  const lookahead = normalizeLookahead(record.sync);

  return {
    microsoft: {
      tenant: normalizeTenant(record.microsoft),
    },
    notion,
    calendars,
    filters: normalizeFilters(record.filters),
    mapping: normalizeMapping(record.mapping),
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

function normalizeNotion(value: unknown): NotionConfig {
  const record = asRecord(value, "`notion` must be an object.");
  const databaseId = requireString(record.databaseId, "`notion.databaseId` is required.");
  const templatePageId = optionalString(record.templatePageId, "`notion.templatePageId` must be a string.");

  return {
    databaseId,
    templatePageId,
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

  if (minDurationMinutes !== undefined && minDurationMinutes <= 0) {
    throw new ConfigError("`filters.minDurationMinutes` must be a positive number.");
  }

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
  };
}

function normalizeMapping(value: unknown): MappingConfig {
  const record = value === undefined ? {} : asRecord(value, "`mapping` must be an object.");

  return {
    title: optionalStringWithDefault(record.title, "Name", "`mapping.title` must be a string."),
    due: optionalStringWithDefault(record.due, "Due", "`mapping.due` must be a string."),
    eventId: optionalStringWithDefault(record.eventId, "Outlook Event ID", "`mapping.eventId` must be a string."),
    changeKey: optionalStringWithDefault(record.changeKey, "Outlook ChangeKey", "`mapping.changeKey` must be a string."),
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

  if (typeof lookahead !== "string" || !VALID_LOOKAHEADS.has(lookahead as LookaheadWindow)) {
    throw new ConfigError("`sync.lookahead` must be one of: today, 24h, 7d.");
  }

  return lookahead as LookaheadWindow;
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
