import type { CalendarConfig } from "../domain/config.js";
import type { GraphEvent } from "../domain/graph.js";
import type { Meeting } from "../domain/meeting.js";
import type { CalendarWindow, MeetingSource } from "../list.js";
import type { AccessTokenProvider } from "./device-code-token-provider.js";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

export class GraphMeetingSource implements MeetingSource {
  constructor(
    private readonly accessTokenProvider: AccessTokenProvider,
    private readonly fetchImpl: typeof fetch = fetch,
  ) {}

  async listMeetings(args: { calendar: CalendarConfig; window: CalendarWindow }): Promise<Meeting[]> {
    const accessToken = await this.accessTokenProvider.getAccessToken();
    const url = new URL(`${GRAPH_BASE_URL}/me/calendars/${encodeURIComponent(args.calendar.id)}/calendarView`);
    url.searchParams.set("startDateTime", args.window.start);
    url.searchParams.set("endDateTime", args.window.end);
    url.searchParams.set(
      "$select",
      [
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
      ].join(","),
    );
    url.searchParams.set("$top", "100");
    url.searchParams.set("$orderby", "start/dateTime");

    const response = await this.fetchImpl(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
      },
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`Graph request failed with ${response.status}: ${body}`);
    }

    const payload = (await response.json()) as { value?: GraphEvent[] };
    const events = payload.value ?? [];

    return events.map((event) => normalizeGraphEvent(event, args.calendar));
  }
}

export function normalizeGraphEvent(event: GraphEvent, calendar: CalendarConfig): Meeting {
  const id = required(event.id, "Graph event is missing `id`.");
  const changeKey = required(event.changeKey, `Graph event ${id} is missing \`changeKey\`.`);
  const start = required(event.start?.dateTime, `Graph event ${id} is missing \`start.dateTime\`.`);
  const end = required(event.end?.dateTime, `Graph event ${id} is missing \`end.dateTime\`.`);
  const details = normalizeBodyText(event.body?.content, event.body?.contentType) ?? event.bodyPreview ?? undefined;
  const fallbackMeetingLink = extractMeetingLink(event.body?.content);

  return {
    id,
    changeKey,
    calendarId: calendar.id,
    calendarName: calendar.name,
    title: event.subject?.trim() || "(untitled meeting)",
    start: toIsoString(start),
    end: toIsoString(end),
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

function toIsoString(value: string): string {
  const date = new Date(value);

  if (Number.isNaN(date.valueOf())) {
    throw new Error(`Invalid Graph date/time: ${value}`);
  }

  return date.toISOString();
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
