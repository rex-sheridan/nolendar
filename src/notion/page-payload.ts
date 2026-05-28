import type { NolendarConfig } from "../domain/config.js";
import type { Meeting } from "../domain/meeting.js";
import type { NotionDataSourceSchema } from "../domain/notion.js";

const NOTION_TEXT_CONTENT_LIMIT = 2000;

export function buildMeetingProperties(
  config: NolendarConfig,
  dataSource: NotionDataSourceSchema,
  meeting: Meeting,
  options: {
    assigneeUserId?: string;
    participantPageIds?: string[];
  } = {},
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
  maybeSetAssignee(properties, dataSource, config.mapping.assignee, options.assigneeUserId);
  maybeSetParticipants(properties, dataSource, config.mapping.participants, options.participantPageIds);

  return properties;
}

export function buildMeetingChildren(config: NolendarConfig, meeting: Meeting): unknown[] {
  const children: unknown[] = [];
  const sections = config.notion.pageContent?.sections ?? [
    "meeting_link",
    "calendar_event",
    "meeting_details",
    "notes",
    "action_items",
  ];

  for (const section of sections) {
    switch (section) {
      case "meeting_link":
        if (meeting.meetingLink) {
          children.push(headingBlock("Meeting Link"));
          children.push(linkParagraphBlock("Open meeting", meeting.meetingLink));
        }
        break;
      case "calendar_event":
        if (meeting.eventLink) {
          children.push(headingBlock("Calendar Event"));
          children.push(linkParagraphBlock("Open event in Outlook", meeting.eventLink));
        }
        break;
      case "meeting_details":
        if (meeting.details || meeting.agenda) {
          children.push(headingBlock("Meeting Details"));
          children.push(...paragraphBlocks(meeting.details ?? meeting.agenda ?? ""));
        }
        break;
      case "notes":
        children.push(headingBlock("Notes"));
        children.push(...paragraphBlocks(" "));
        break;
      case "action_items":
        children.push(headingBlock("Action items"));
        children.push(...paragraphBlocks(" "));
        break;
    }
  }

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

function maybeSetAssignee(
  properties: Record<string, unknown>,
  dataSource: NotionDataSourceSchema,
  propertyName: string | undefined,
  userId?: string,
): void {
  if (!propertyName || !userId || dataSource.properties[propertyName]?.type !== "people") {
    return;
  }

  properties[propertyName] = {
    people: [{ id: userId }],
  };
}

function maybeSetParticipants(
  properties: Record<string, unknown>,
  dataSource: NotionDataSourceSchema,
  propertyName: string | undefined,
  participantPageIds?: string[],
): void {
  if (
    !propertyName ||
    !participantPageIds ||
    participantPageIds.length === 0 ||
    dataSource.properties[propertyName]?.type !== "relation"
  ) {
    return;
  }

  properties[propertyName] = {
    relation: participantPageIds.map((id) => ({ id })),
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

function paragraphBlocks(content: string): unknown[] {
  return splitTextContent(content).map(paragraphBlock);
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

function splitTextContent(content: string): string[] {
  if (content.length <= NOTION_TEXT_CONTENT_LIMIT) {
    return [content];
  }

  const chunks: string[] = [];
  let remaining = content;

  while (remaining.length > NOTION_TEXT_CONTENT_LIMIT) {
    const splitAt = findSplitIndex(remaining);
    chunks.push(remaining.slice(0, splitAt).trimEnd());
    remaining = remaining.slice(splitAt).trimStart();
  }

  if (remaining.length > 0) {
    chunks.push(remaining);
  }

  return chunks;
}

function findSplitIndex(value: string): number {
  const hardLimit = Math.min(value.length, NOTION_TEXT_CONTENT_LIMIT);
  const preferredBreak = Math.max(
    value.lastIndexOf("\n", hardLimit - 1),
    value.lastIndexOf(". ", hardLimit - 1),
    value.lastIndexOf(" ", hardLimit - 1),
  );

  return preferredBreak > 0 ? Math.min(preferredBreak + 1, hardLimit) : hardLimit;
}

function linkParagraphBlock(label: string, url: string): unknown {
  return {
    object: "block",
    type: "paragraph",
    paragraph: {
      rich_text: [
        {
          type: "text",
          text: {
            content: label,
            link: {
              url,
            },
          },
        },
      ],
    },
  };
}
