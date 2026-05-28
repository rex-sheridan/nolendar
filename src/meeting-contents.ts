import type { NolendarConfig } from "./domain/config.js";
import type { Meeting } from "./domain/meeting.js";
import type { NotionMeetingPage } from "./domain/notion.js";
import type { Clock } from "./clock.js";
import { systemClock } from "./clock.js";
import type { CalendarWindow, MeetingSource } from "./list.js";
import type { NotionClient } from "./notion/client.js";

export type MeetingContentsSource = "outlook" | "notion";
export type MeetingContentsDetail = "compact" | "full";

export interface DayWindow {
  label: string;
  start: string;
  end: string;
}

export interface MeetingContentsOptions {
  source: MeetingContentsSource;
  day?: string;
  detail: MeetingContentsDetail;
}

export interface MeetingContentsDependencies {
  meetingSource?: MeetingSource;
  notion?: NotionClient;
  clock?: Clock;
}

export async function listMeetingContentsForDay(
  config: NolendarConfig,
  options: MeetingContentsOptions,
  deps: MeetingContentsDependencies,
): Promise<string[]> {
  const window = resolveDayWindow(options.day ?? "today", deps.clock ?? systemClock);

  if (options.source === "outlook") {
    if (!deps.meetingSource) {
      throw new Error("Outlook meeting contents require a meeting source.");
    }

    const meetings = await listOutlookMeetingsForWindow(config, window, deps.meetingSource);
    return formatOutlookMeetingContents(meetings, {
      detail: options.detail,
      window,
    });
  }

  if (!deps.notion?.listMeetingPagesForWindow) {
    throw new Error("Notion meeting contents require a Notion client with page listing support.");
  }

  const pages = await deps.notion.listMeetingPagesForWindow({
    dataSourceId: config.notion.databaseId,
    datePropertyName: config.mapping.due,
    start: window.start,
    end: window.end,
  });

  return formatNotionMeetingContents(pages, {
    config,
    detail: options.detail,
    window,
  });
}

export function resolveDayWindow(day: string, clock: Clock = systemClock): DayWindow {
  const base = clock.now();
  const normalized = day.trim().toLowerCase();
  const start = new Date(base);

  if (normalized === "today") {
    start.setUTCHours(0, 0, 0, 0);
  } else if (normalized === "tomorrow") {
    start.setUTCHours(0, 0, 0, 0);
    start.setUTCDate(start.getUTCDate() + 1);
  } else if (normalized === "yesterday") {
    start.setUTCHours(0, 0, 0, 0);
    start.setUTCDate(start.getUTCDate() - 1);
  } else if (/^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
    start.setTime(Date.parse(`${normalized}T00:00:00.000Z`));
  } else if (/^[+-][1-9]\d*d$/.test(normalized)) {
    start.setUTCHours(0, 0, 0, 0);
    start.setUTCDate(start.getUTCDate() + Number.parseInt(normalized.slice(0, -1), 10));
  } else {
    throw new Error("`--day` must be today, tomorrow, yesterday, +/-Nd, or YYYY-MM-DD.");
  }

  if (Number.isNaN(start.valueOf())) {
    throw new Error("`--day` must be today, tomorrow, yesterday, +/-Nd, or YYYY-MM-DD.");
  }

  const end = new Date(start);
  end.setUTCDate(end.getUTCDate() + 1);

  return {
    label: start.toISOString().slice(0, 10),
    start: start.toISOString(),
    end: end.toISOString(),
  };
}

async function listOutlookMeetingsForWindow(
  config: NolendarConfig,
  window: CalendarWindow,
  meetingSource: MeetingSource,
): Promise<Meeting[]> {
  const meetingsByCalendar = await Promise.all(
    config.calendars.map((calendar) =>
      meetingSource.listMeetings({
        calendar,
        window,
      }),
    ),
  );

  return meetingsByCalendar
    .flat()
    .sort((left, right) => left.start.localeCompare(right.start) || left.title.localeCompare(right.title));
}

function formatOutlookMeetingContents(
  meetings: Meeting[],
  options: {
    detail: MeetingContentsDetail;
    window: DayWindow;
  },
): string[] {
  const lines = formatContentsHeader("outlook", options.window, meetings.length);

  meetings.forEach((meeting, index) => {
    lines.push("", `## Meeting ${index + 1}: ${meeting.title || "(untitled meeting)"}`);
    lines.push(`Title: ${meeting.title || "(untitled meeting)"}`);
    lines.push(`Date: ${meeting.start}${meeting.end ? ` to ${meeting.end}` : ""}`);
    lines.push("Body:");
    lines.push(...indentBody(meeting.details ?? meeting.agenda ?? ""));

    if (options.detail === "full") {
      lines.push("Properties:");
      lines.push(...indentBody(JSON.stringify(meeting, null, 2)));
    }
  });

  return lines;
}

function formatNotionMeetingContents(
  pages: NotionMeetingPage[],
  options: {
    config: NolendarConfig;
    detail: MeetingContentsDetail;
    window: DayWindow;
  },
): string[] {
  const lines = formatContentsHeader("notion", options.window, pages.length);

  pages.forEach((page, index) => {
    const title = stringProperty(page.properties[options.config.mapping.title]) || "(untitled meeting)";
    const date = formatPropertyValue(page.properties[options.config.mapping.due]);

    lines.push("", `## Meeting ${index + 1}: ${title}`);
    lines.push(`Title: ${title}`);
    lines.push(`Date: ${date || "Unknown"}`);
    lines.push("Body:");
    lines.push(...indentBody(page.body));

    if (options.detail === "full") {
      lines.push("Properties:");
      lines.push(...indentBody(JSON.stringify({ id: page.id, url: page.url, properties: page.properties }, null, 2)));
    }
  });

  return lines;
}

function formatContentsHeader(source: MeetingContentsSource, window: DayWindow, count: number): string[] {
  return [
    `# Meeting contents (${source})`,
    `Day: ${window.label}`,
    `Window: ${window.start} to ${window.end}`,
    `Count: ${count}`,
  ];
}

function indentBody(value: string): string[] {
  const body = value.trim() || "(empty)";
  return body.split("\n").map((line) => `  ${line}`);
}

function stringProperty(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value : undefined;
}

function formatPropertyValue(value: unknown): string {
  if (value == null) {
    return "";
  }

  if (typeof value === "string") {
    return value;
  }

  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }

  if (Array.isArray(value)) {
    return value.map(formatPropertyValue).filter(Boolean).join(", ");
  }

  if (typeof value === "object") {
    const record = value as { start?: unknown; end?: unknown };
    if (typeof record.start === "string") {
      return record.end ? `${record.start} to ${String(record.end)}` : record.start;
    }
  }

  return JSON.stringify(value);
}
