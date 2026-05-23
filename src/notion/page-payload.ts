import type { NolendarConfig } from "../domain/config.js";
import type { Meeting } from "../domain/meeting.js";
import type { NotionDataSourceSchema } from "../domain/notion.js";

export function buildMeetingProperties(
  config: NolendarConfig,
  dataSource: NotionDataSourceSchema,
  meeting: Meeting,
): Record<string, unknown> {
  const properties: Record<string, unknown> = {
    [config.mapping.title]: {
      title: [textBlock(meeting.title || "(untitled meeting)")],
    },
    [config.mapping.due]: {
      date: {
        start: meeting.start,
        end: meeting.end,
      },
    },
    [config.mapping.eventId]: {
      rich_text: [textBlock(meeting.id)],
    },
    [config.mapping.changeKey]: {
      rich_text: [textBlock(meeting.changeKey)],
    },
  };

  maybeSetRichText(properties, dataSource, "Organizer", meeting.organizer);
  maybeSetRichText(properties, dataSource, "Attendees", formatAttendees(meeting));
  maybeSetRichText(properties, dataSource, "Calendar", meeting.calendarName ?? meeting.calendarId);
  maybeSetUrl(properties, dataSource, "Meeting Link", meeting.meetingLink);
  maybeSetUrl(properties, dataSource, "Event Link", meeting.eventLink);
  maybeSetConfiguredUrl(properties, dataSource, config.mapping.eventLink, meeting.eventLink);
  maybeSetTags(properties, dataSource, config.mapping.tags, config.notion.defaultTags);

  return properties;
}

export function buildMeetingChildren(meeting: Meeting): unknown[] {
  const children: unknown[] = [];

  if (meeting.agenda) {
    children.push(headingBlock("Agenda"));
    children.push(paragraphBlock(meeting.agenda));
  }

  children.push(headingBlock("Notes"));
  children.push(paragraphBlock(" "));
  children.push(headingBlock("Action items"));
  children.push(paragraphBlock(" "));

  return children;
}

function maybeSetRichText(
  properties: Record<string, unknown>,
  dataSource: NotionDataSourceSchema,
  propertyName: string,
  value?: string,
): void {
  if (!value || dataSource.properties[propertyName]?.type !== "rich_text") {
    return;
  }

  properties[propertyName] = {
    rich_text: [textBlock(value)],
  };
}

function maybeSetUrl(
  properties: Record<string, unknown>,
  dataSource: NotionDataSourceSchema,
  propertyName: string,
  value?: string,
): void {
  if (!value || dataSource.properties[propertyName]?.type !== "url") {
    return;
  }

  properties[propertyName] = {
    url: value,
  };
}

function maybeSetConfiguredUrl(
  properties: Record<string, unknown>,
  dataSource: NotionDataSourceSchema,
  propertyName: string | undefined,
  value?: string,
): void {
  if (!propertyName) {
    return;
  }

  maybeSetUrl(properties, dataSource, propertyName, value);
}

function maybeSetTags(
  properties: Record<string, unknown>,
  dataSource: NotionDataSourceSchema,
  propertyName: string | undefined,
  tags?: string[],
): void {
  if (!propertyName || !tags || tags.length === 0 || dataSource.properties[propertyName]?.type !== "multi_select") {
    return;
  }

  properties[propertyName] = {
    multi_select: tags.map((tag) => ({ name: tag })),
  };
}

function formatAttendees(meeting: Meeting): string | undefined {
  if (meeting.attendees.length === 0) {
    return undefined;
  }

  return meeting.attendees
    .map((attendee) => attendee.name ?? attendee.email ?? "Unknown attendee")
    .join(", ");
}

function textBlock(content: string): { type: "text"; text: { content: string } } {
  return {
    type: "text",
    text: {
      content,
    },
  };
}

function headingBlock(content: string): unknown {
  return {
    object: "block",
    type: "heading_2",
    heading_2: {
      rich_text: [textBlock(content)],
    },
  };
}

function paragraphBlock(content: string): unknown {
  return {
    object: "block",
    type: "paragraph",
    paragraph: {
      rich_text: [textBlock(content)],
    },
  };
}
