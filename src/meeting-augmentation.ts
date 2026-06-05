import type { NolendarConfig } from "./domain/config.js";
import type { NotionMeetingPageProperties } from "./domain/notion.js";
import type { DayWindow } from "./meeting-contents.js";
import { resolveDayWindow } from "./meeting-contents.js";
import type { NotionClient } from "./notion/client.js";

export interface MeetingAugmentationSection {
  title: string;
  content: string;
  time?: MeetingSectionTime;
}

export interface MeetingSectionTime {
  startMinutes: number;
  endMinutes?: number;
}

export interface ImportMeetingAugmentationResult {
  day: DayWindow;
  matched: Array<{
    title: string;
    pageId: string;
    url?: string;
    appended: boolean;
  }>;
  unmatched: MeetingAugmentationSection[];
  ambiguous: Array<{
    title: string;
    pageIds: string[];
  }>;
  empty: string[];
}

export async function importMeetingAugmentation(
  config: NolendarConfig,
  input: string,
  options: {
    day?: string;
    dryRun?: boolean;
    heading: string;
  },
  deps: {
    notion: NotionClient;
  },
): Promise<ImportMeetingAugmentationResult> {
  if (!deps.notion.listMeetingPagePropertiesForWindow) {
    throw new Error("Meeting augmentation requires a Notion client with page listing support.");
  }

  const heading = options.heading.trim();
  if (!heading) {
    throw new Error("`--heading` must not be empty.");
  }

  const day = resolveDayWindow(options.day ?? "today");
  const pages = filterAugmentableMeetingPages(
    config,
    day,
    await deps.notion.listMeetingPagePropertiesForWindow({
      dataSourceId: config.notion.databaseId,
      datePropertyName: config.mapping.due,
      start: day.start,
      end: day.end,
    }),
  );
  const pageMatches = buildPageTitleMatches(config, pages);
  const sections = parseMeetingAugmentationSections(input, listPageTitles(config, pages));
  const result: ImportMeetingAugmentationResult = {
    day,
    matched: [],
    unmatched: [],
    ambiguous: [],
    empty: [],
  };

  for (const section of sections) {
    if (!section.content.trim()) {
      result.empty.push(section.title);
      continue;
    }

    const matches = pageMatches.get(normalizeMeetingTitle(section.title)) ?? [];
    const resolvedMatches = resolveMatchesBySectionTime(config, section, matches);

    if (resolvedMatches.length === 0) {
      result.unmatched.push(section);
      continue;
    }

    if (resolvedMatches.length > 1) {
      result.ambiguous.push({
        title: section.title,
        pageIds: resolvedMatches.map((page) => page.id),
      });
      continue;
    }

    const page = resolvedMatches[0];
    if (!options.dryRun) {
      if (!deps.notion.appendMarkdownToPage) {
        throw new Error("Meeting augmentation requires a Notion client with append support.");
      }

      await deps.notion.appendMarkdownToPage({
        pageId: page.id,
        heading,
        content: section.content,
      });
    }

    result.matched.push({
      title: section.title,
      pageId: page.id,
      url: page.url,
      appended: !options.dryRun,
    });
  }

  return result;
}

export function parseMeetingAugmentationSections(input: string, knownTitles: string[]): MeetingAugmentationSection[] {
  const titleLookup = buildTitleLookup(knownTitles);
  const sections: MeetingAugmentationSection[] = [];
  let currentTitle: string | undefined;
  let currentLines: string[] = [];

  const flush = () => {
    if (!currentTitle) {
      return;
    }

    sections.push({
      title: currentTitle,
      content: currentLines.join("\n").trim(),
      time: extractSectionTime(currentLines),
    });
    currentLines = [];
  };

  for (const rawLine of input.replace(/\r\n/g, "\n").split("\n")) {
    const matchedTitle = findKnownTitle(rawLine, titleLookup);

    if (matchedTitle) {
      flush();
      currentTitle = matchedTitle;
      currentLines = [];
      continue;
    }

    if (currentTitle) {
      currentLines.push(rawLine);
    }
  }

  flush();
  return sections;
}

function resolveMatchesBySectionTime(
  config: NolendarConfig,
  section: MeetingAugmentationSection,
  matches: NotionMeetingPageProperties[],
): NotionMeetingPageProperties[] {
  const sectionTime = section.time;

  if (matches.length <= 1 || !sectionTime) {
    return matches;
  }

  const timeMatches = matches.filter((page) => pageMatchesSectionTime(page, config.mapping.due, sectionTime));
  return timeMatches.length > 0 ? timeMatches : matches;
}

function pageMatchesSectionTime(
  page: NotionMeetingPageProperties,
  datePropertyName: string,
  sectionTime: MeetingSectionTime,
): boolean {
  const date = page.properties[datePropertyName];

  if (!date || typeof date !== "object" || Array.isArray(date)) {
    return false;
  }

  const start = readDateStart(date);
  if (!start) {
    return false;
  }

  if (isoStringIncludesClock(start, sectionTime.startMinutes)) {
    return true;
  }

  const startCandidates = timeCandidatesFromIso(start);
  return startCandidates.includes(sectionTime.startMinutes);
}

function readDateStart(value: object): string | undefined {
  const record = value as { start?: unknown; date?: unknown };

  if (typeof record.start === "string") {
    return record.start;
  }

  if (record.date && typeof record.date === "object" && !Array.isArray(record.date)) {
    const rawDate = record.date as { start?: unknown };
    return typeof rawDate.start === "string" ? rawDate.start : undefined;
  }

  return undefined;
}

export function normalizeMeetingTitle(title: string): string {
  return title
    .normalize("NFKC")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function buildTitleLookup(knownTitles: string[]): Map<string, string> {
  const titleLookup = new Map<string, string>();

  for (const title of knownTitles) {
    for (const alias of titleAliases(title)) {
      titleLookup.set(alias, title);
    }
  }

  return titleLookup;
}

function titleAliases(title: string): string[] {
  return Array.from(
    new Set([
      normalizeMeetingTitle(title),
      normalizeComparableTitle(title),
      normalizeComparableTitle(cleanCandidateTitle(title)),
    ]),
  ).filter(Boolean);
}

function buildPageTitleMatches(
  config: NolendarConfig,
  pages: NotionMeetingPageProperties[],
): Map<string, NotionMeetingPageProperties[]> {
  const matches = new Map<string, NotionMeetingPageProperties[]>();

  for (const page of pages) {
    const title = page.properties[config.mapping.title];

    if (typeof title !== "string" || !title.trim()) {
      continue;
    }

    const normalized = normalizeMeetingTitle(title);
    matches.set(normalized, [...(matches.get(normalized) ?? []), page]);
  }

  return matches;
}

function filterAugmentableMeetingPages(
  config: NolendarConfig,
  day: DayWindow,
  pages: NotionMeetingPageProperties[],
): NotionMeetingPageProperties[] {
  return pages.filter(
    (page) => pageIsOnDay(page, config.mapping.due, day) && !pageIsArchived(page) && !pageIsCanceled(config, page),
  );
}

function pageIsOnDay(page: NotionMeetingPageProperties, datePropertyName: string, day: DayWindow): boolean {
  const date = page.properties[datePropertyName];

  if (!date || typeof date !== "object" || Array.isArray(date)) {
    return false;
  }

  const start = readDateStart(date);
  if (!start) {
    return false;
  }

  const startTime = Date.parse(start);
  if (Number.isNaN(startTime)) {
    return false;
  }

  return startTime >= Date.parse(day.start) && startTime < Date.parse(day.end);
}

function pageIsArchived(page: NotionMeetingPageProperties): boolean {
  return page.archived === true || page.inTrash === true;
}

function pageIsCanceled(config: NolendarConfig, page: NotionMeetingPageProperties): boolean {
  const canceledMeetings = config.notion.canceledMeetings;

  if (canceledMeetings?.action !== "set_status") {
    return false;
  }

  return readStatusName(page.properties[canceledMeetings.statusProperty]) === canceledMeetings.statusValue;
}

function readStatusName(value: unknown): string | undefined {
  if (typeof value === "string") {
    return value;
  }

  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return undefined;
  }

  const record = value as { status?: unknown };
  if (!record.status || typeof record.status !== "object" || Array.isArray(record.status)) {
    return undefined;
  }

  const status = record.status as { name?: unknown };
  return typeof status.name === "string" ? status.name : undefined;
}

function listPageTitles(config: NolendarConfig, pages: NotionMeetingPageProperties[]): string[] {
  return pages
    .map((page) => page.properties[config.mapping.title])
    .filter((title): title is string => typeof title === "string" && Boolean(title.trim()));
}

function findKnownTitle(line: string, titleLookup: Map<string, string>): string | undefined {
  for (const candidate of extractTitleCandidates(line)) {
    const match =
      titleLookup.get(normalizeMeetingTitle(candidate)) ??
      titleLookup.get(normalizeComparableTitle(candidate)) ??
      titleLookup.get(normalizeComparableTitle(cleanCandidateTitle(candidate)));

    if (match) {
      return match;
    }
  }

  return undefined;
}

function extractTitleCandidates(line: string): string[] {
  const trimmed = line.trim();

  if (!trimmed) {
    return [];
  }

  const withoutHeading = trimmed.replace(/^#{1,6}\s+/, "").trim();
  const withoutNumbering = withoutHeading.replace(/^\d+[.)]\s+/, "").trim();
  const candidates = [
    withoutNumbering,
    stripOuterEmphasis(withoutNumbering),
    extractStandaloneMarkdownLinkTitle(stripOuterEmphasis(withoutNumbering)),
    extractBracketedSearchResultTitle(stripOuterEmphasis(withoutNumbering)),
  ].filter((candidate): candidate is string => Boolean(candidate?.trim()));

  return Array.from(new Set(candidates.map(cleanCandidateTitle)));
}

function stripOuterEmphasis(value: string): string {
  return value
    .replace(/^\*\*(.+)\*\*$/, "$1")
    .replace(/^__(.+)__$/, "$1")
    .trim();
}

function extractStandaloneMarkdownLinkTitle(value: string): string | undefined {
  const match = value.match(/^\[([^\]]+)\]\([^)]*\)$/);
  return match?.[1];
}

function extractBracketedSearchResultTitle(value: string): string | undefined {
  const match = value.match(/^\[?\\?\[?(.+?)(?:\s+\|\s+Meeting)?\\?\]?\]?$/);

  if (!match || match[1] === value) {
    return undefined;
  }

  return match[1];
}

function cleanCandidateTitle(value: string): string {
  return value
    .replace(/\s+\|\s+Meeting$/i, "")
    .replace(/\\([[\]])/g, "$1")
    .trim();
}

function extractSectionTime(lines: string[]): MeetingSectionTime | undefined {
  for (const line of lines) {
    const timeLine = stripMarkdownInline(line);
    const timeMatch = timeLine.match(
      /\btime\s*:\s*(?:today|tomorrow|yesterday|[A-Za-z]+,\s*)?\s*(\d{1,2})(?::(\d{2}))?\s*(AM|PM|A\.M\.|P\.M\.)\s*(?:[-–—]\s*(\d{1,2})(?::(\d{2}))?\s*(AM|PM|A\.M\.|P\.M\.))?/i,
    );

    if (timeMatch) {
      const startPeriod = normalizePeriod(timeMatch[3]);
      const endPeriod = normalizePeriod(timeMatch[6]) ?? startPeriod;
      const startMinutes = parseClockMinutes(timeMatch[1], timeMatch[2], startPeriod);
      const endMinutes = timeMatch[4] ? parseClockMinutes(timeMatch[4], timeMatch[5], endPeriod) : undefined;

      if (startMinutes !== undefined) {
        return {
          startMinutes,
          ...(endMinutes !== undefined ? { endMinutes } : {}),
        };
      }
    }

    const looseMatch = timeLine.match(
      /\b(\d{1,2})(?::(\d{2}))?\s*(AM|PM|A\.M\.|P\.M\.)\s*(?:[-–—]\s*(\d{1,2})(?::(\d{2}))?\s*(AM|PM|A\.M\.|P\.M\.))?/i,
    );

    if (looseMatch) {
      const startPeriod = normalizePeriod(looseMatch[3]);
      const endPeriod = normalizePeriod(looseMatch[6]) ?? startPeriod;
      const startMinutes = parseClockMinutes(looseMatch[1], looseMatch[2], startPeriod);
      const endMinutes = looseMatch[4] ? parseClockMinutes(looseMatch[4], looseMatch[5], endPeriod) : undefined;

      if (startMinutes !== undefined) {
        return {
          startMinutes,
          ...(endMinutes !== undefined ? { endMinutes } : {}),
        };
      }
    }
  }

  return undefined;
}

function stripMarkdownInline(value: string): string {
  return value
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/\*([^*]+)\*/g, "$1")
    .replace(/__([^_]+)__/g, "$1")
    .replace(/_([^_]+)_/g, "$1")
    .trim();
}

function normalizePeriod(value: string | undefined): "AM" | "PM" | undefined {
  const normalized = value?.replace(/\./g, "").toUpperCase();

  if (normalized === "AM" || normalized === "PM") {
    return normalized;
  }

  return undefined;
}

function parseClockMinutes(hourValue: string, minuteValue: string | undefined, period: "AM" | "PM" | undefined): number | undefined {
  const hour = Number.parseInt(hourValue, 10);
  const minute = minuteValue ? Number.parseInt(minuteValue, 10) : 0;

  if (!Number.isInteger(hour) || !Number.isInteger(minute) || minute < 0 || minute > 59) {
    return undefined;
  }

  if (period) {
    if (hour < 1 || hour > 12) {
      return undefined;
    }

    const normalizedHour = hour === 12 ? (period === "AM" ? 0 : 12) : period === "PM" ? hour + 12 : hour;
    return normalizedHour * 60 + minute;
  }

  if (hour < 0 || hour > 23) {
    return undefined;
  }

  return hour * 60 + minute;
}

function timeCandidatesFromIso(value: string): number[] {
  const candidates = new Set<number>();
  const rawTimeMatch = value.match(/T(\d{2}):(\d{2})/);

  if (rawTimeMatch) {
    const rawMinutes = parseClockMinutes(rawTimeMatch[1], rawTimeMatch[2], undefined);
    if (rawMinutes !== undefined) {
      candidates.add(rawMinutes);
    }
  }

  const parsedDate = new Date(value);
  if (!Number.isNaN(parsedDate.valueOf())) {
    candidates.add(parsedDate.getHours() * 60 + parsedDate.getMinutes());
  }

  return Array.from(candidates);
}

function isoStringIncludesClock(value: string, minutes: number): boolean {
  const hour = Math.floor(minutes / 60);
  const minute = minutes % 60;
  const clock = `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`;

  return value.includes(`T${clock}`);
}

function normalizeComparableTitle(title: string): string {
  return cleanCandidateTitle(title)
    .normalize("NFKC")
    .replace(/&amp;/gi, "&")
    .replace(/\s+\|\s+meeting$/i, "")
    .replace(/['']/g, "'")
    .replace(/[""]/g, "\"")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}
