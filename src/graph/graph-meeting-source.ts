import type { CalendarConfig } from "../domain/config.js";
import type { GraphEvent } from "../domain/graph.js";
import type { Meeting } from "../domain/meeting.js";
import type { CalendarWindow, MeetingSource } from "../list.js";
import type { AccessTokenProvider } from "./device-code-token-provider.js";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const MAX_THROTTLE_RETRIES = 3;

export class GraphMeetingSource implements MeetingSource {
  constructor(
    private readonly accessTokenProvider: AccessTokenProvider,
    private readonly fetchImpl: typeof fetch = fetch,
  ) {}

  async listMeetings(args: { calendar: CalendarConfig; window: CalendarWindow }): Promise<Meeting[]> {
    const accessToken = await this.accessTokenProvider.getAccessToken();
    const payload = await fetchGraphPayload(this.fetchImpl, buildCalendarViewUrl(args.calendar.id, args.window), accessToken);

    return (payload.value ?? [])
      .filter(isSyncableCalendarViewEvent)
      .map((event) => normalizeGraphEvent(event, args.calendar));
  }

  async listMeetingChanges(args: {
    calendar: CalendarConfig;
    window: CalendarWindow;
    deltaLink?: string;
  }): Promise<{ meetings: Meeting[]; removedEventIds: string[]; deltaLink?: string }> {
    const accessToken = await this.accessTokenProvider.getAccessToken();
    let nextUrl: URL | undefined = args.deltaLink
      ? new URL(args.deltaLink)
      : buildCalendarViewDeltaUrl(args.calendar.id, args.window);
    const meetings: Meeting[] = [];
    const removedEventIds: string[] = [];
    let deltaLink: string | undefined;

    while (nextUrl) {
      const payload = await fetchGraphPayload(this.fetchImpl, nextUrl, accessToken, {
        maxPageSize: 100,
      });
      const events = payload.value ?? [];

      for (const event of events) {
        if (event["@removed"] != null) {
          if (event.id) {
            removedEventIds.push(event.id);
          }
          continue;
        }

        if (!isSyncableCalendarViewEvent(event)) {
          continue;
        }

        const hydratedEvent = await this.hydrateGraphEvent(args.calendar.id, event, accessToken);
        meetings.push(normalizeGraphEvent(hydratedEvent, args.calendar));
      }
      nextUrl = payload["@odata.nextLink"] ? new URL(payload["@odata.nextLink"]) : undefined;
      deltaLink = payload["@odata.deltaLink"] ?? deltaLink;
    }

    return {
      meetings,
      removedEventIds,
      deltaLink,
    };
  }

  private async hydrateGraphEvent(calendarId: string, event: GraphEvent, accessToken: string): Promise<GraphEvent> {
    if (!needsHydration(event)) {
      return event;
    }

    const eventId = required(event.id, "Graph delta event is missing `id`.");
    const payload = await fetchGraphPayload(
      this.fetchImpl,
      buildEventUrl(calendarId, eventId),
      accessToken,
    );
    const hydratedEvent = payload.value?.[0] ?? payload;

    return {
      ...event,
      ...hydratedEvent,
    } as GraphEvent;
  }
}

function buildCalendarViewUrl(calendarId: string, window: CalendarWindow): URL {
  const url = new URL(`${GRAPH_BASE_URL}/me/calendars/${encodeURIComponent(calendarId)}/calendarView`);
  url.searchParams.set("startDateTime", window.start);
  url.searchParams.set("endDateTime", window.end);
  url.searchParams.set("$select", graphEventSelect());
  url.searchParams.set("$top", "100");
  url.searchParams.set("$orderby", "start/dateTime");
  return url;
}

function buildCalendarViewDeltaUrl(calendarId: string, window: CalendarWindow): URL {
  const url = new URL(`${GRAPH_BASE_URL}/me/calendars/${encodeURIComponent(calendarId)}/calendarView/delta`);
  url.searchParams.set("startDateTime", window.start);
  url.searchParams.set("endDateTime", window.end);
  url.searchParams.set("$select", graphEventSelect());
  return url;
}

function buildEventUrl(calendarId: string, eventId: string): URL {
  const url = new URL(`${GRAPH_BASE_URL}/me/calendars/${encodeURIComponent(calendarId)}/events/${encodeURIComponent(eventId)}`);
  url.searchParams.set("$select", graphEventSelect());
  return url;
}

async function fetchGraphPayload(
  fetchImpl: typeof fetch,
  url: URL,
  accessToken: string,
  options: {
    maxPageSize?: number;
  } = {},
): Promise<{ value?: GraphEvent[]; "@odata.nextLink"?: string; "@odata.deltaLink"?: string }> {
  const preferHeaders = ['outlook.timezone="UTC"'];

  if (options.maxPageSize !== undefined) {
    preferHeaders.push(`odata.maxpagesize=${options.maxPageSize}`);
  }

  for (let attempt = 0; attempt <= MAX_THROTTLE_RETRIES; attempt += 1) {
    const response = await fetchImpl(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
        Prefer: preferHeaders.join(", "),
      },
    });

    if (response.ok) {
      return (await response.json()) as { value?: GraphEvent[]; "@odata.nextLink"?: string; "@odata.deltaLink"?: string };
    }

    if (response.status === 429 && attempt < MAX_THROTTLE_RETRIES) {
      const retryAfterMs = getRetryDelayMs(response, attempt);
      await sleep(retryAfterMs);
      continue;
    }

    const body = await response.text();
    throw new Error(`Graph request failed with ${response.status}: ${body}`);
  }

  throw new Error("Graph request failed after retries.");
}

function getRetryDelayMs(response: Response, attempt: number): number {
  const retryAfter = response.headers.get("Retry-After");

  if (retryAfter) {
    const parsedSeconds = Number.parseInt(retryAfter, 10);
    if (Number.isFinite(parsedSeconds) && parsedSeconds >= 0) {
      return parsedSeconds * 1000;
    }
  }

  return (attempt + 1) * 1000;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

function needsHydration(event: GraphEvent): boolean {
  return !event.changeKey || !event.start?.dateTime || !event.end?.dateTime;
}

function isSyncableCalendarViewEvent(event: GraphEvent): boolean {
  return event.type !== "seriesMaster";
}

function graphEventSelect(): string {
  return [
    "id",
    "changeKey",
    "subject",
    "start",
    "end",
    "organizer",
    "attendees",
    "onlineMeeting",
    "onlineMeetingUrl",
    "webLink",
    "bodyPreview",
    "body",
    "responseStatus",
    "type",
    "isCancelled",
    "recurrence",
  ].join(",");
}

export function normalizeGraphEvent(event: GraphEvent, calendar: CalendarConfig): Meeting {
  const id = required(event.id, "Graph event is missing `id`.");
  const changeKey = required(event.changeKey, `Graph event ${id} is missing \`changeKey\`.`);
  const start = normalizeGraphDateTime(event.start, `Graph event ${id} is missing \`start.dateTime\`.`);
  const end = normalizeGraphDateTime(event.end, `Graph event ${id} is missing \`end.dateTime\`.`);
  const details = normalizeBodyText(event.body?.content, event.body?.contentType) ?? event.bodyPreview ?? undefined;
  const fallbackMeetingLink = extractMeetingLink(event.body?.content);

  return {
    id,
    changeKey,
    calendarId: calendar.id,
    calendarName: calendar.name,
    title: event.subject?.trim() || "(untitled meeting)",
    start,
    end,
    organizer: event.organizer?.emailAddress?.name ?? event.organizer?.emailAddress?.address ?? undefined,
    attendees:
      event.attendees?.map((attendee) => ({
        name: attendee.emailAddress?.name ?? undefined,
        email: attendee.emailAddress?.address ?? undefined,
        optional: attendee.type?.toLowerCase() === "optional",
      })) ?? [],
    meetingLink: event.onlineMeeting?.joinUrl ?? event.onlineMeetingUrl ?? fallbackMeetingLink ?? undefined,
    eventLink: event.webLink ?? undefined,
    agenda: event.bodyPreview ?? undefined,
    details,
    responseStatus: event.responseStatus?.response ?? undefined,
    isCancelled: Boolean(event.isCancelled),
    isRecurring: event.type === "seriesMaster" || event.type === "occurrence" || event.recurrence != null,
  };
}

function required(value: string | null | undefined, message: string): string {
  if (!value) {
    throw new Error(message);
  }

  return value;
}

function normalizeGraphDateTime(
  value: { dateTime?: string | null; timeZone?: string | null } | null | undefined,
  missingMessage: string,
): string {
  const dateTime = required(value?.dateTime, missingMessage);
  const normalizedInput = hasExplicitOffset(dateTime) ? dateTime : applyGraphTimezone(dateTime, value?.timeZone);
  const date = new Date(normalizedInput);

  if (Number.isNaN(date.valueOf())) {
    throw new Error(`Invalid Graph date/time: ${dateTime}`);
  }

  return date.toISOString();
}

function hasExplicitOffset(value: string): boolean {
  return /(?:Z|[+-]\d{2}:\d{2})$/i.test(value);
}

function applyGraphTimezone(dateTime: string, timeZone?: string | null): string {
  const normalizedZone = timeZone?.trim().toUpperCase();

  if (!normalizedZone) {
    return dateTime;
  }

  if (normalizedZone === "UTC" || normalizedZone === "ETC/UTC" || normalizedZone === "GMT" || normalizedZone === "ETC/GMT") {
    return `${dateTime}Z`;
  }

  return dateTime;
}

function normalizeBodyText(content?: string | null, contentType?: string | null): string | undefined {
  if (!content || content.trim() === "") {
    return undefined;
  }

  const normalized =
    contentType?.toLowerCase() === "html"
      ? decodeHtmlEntities(
          content
            .replace(/<style[\s\S]*?<\/style>/gi, " ")
            .replace(/<script[\s\S]*?<\/script>/gi, " ")
            .replace(/<\/(p|div|li|tr|br|h1|h2|h3|h4|h5|h6)>/gi, "\n")
            .replace(/<[^>]+>/g, " "),
        )
      : content;

  const compact = normalized
    .split("\n")
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter(Boolean)
    .join("\n");

  return compact || undefined;
}

function extractMeetingLink(content?: string | null): string | undefined {
  if (!content) {
    return undefined;
  }

  const candidates = new Set<string>();

  for (const match of content.matchAll(/href=["']([^"']+)["']/gi)) {
    if (match[1]) {
      candidates.add(decodeHtmlEntities(match[1]));
    }
  }

  for (const match of content.matchAll(/https:\/\/[^\s"'<>]+/gi)) {
    candidates.add(decodeHtmlEntities(match[0]));
  }

  return Array.from(candidates).find((candidate) =>
    /(teams\.microsoft\.com\/l\/meetup-join|teams\.live\.com\/meet|meet\.office\.com)/i.test(candidate),
  );
}

function decodeHtmlEntities(value: string): string {
  return value
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
}
