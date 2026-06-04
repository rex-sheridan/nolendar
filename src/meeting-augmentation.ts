import type { NolendarConfig } from "./domain/config.js";
import type { NotionMeetingPageProperties } from "./domain/notion.js";
import type { DayWindow } from "./meeting-contents.js";
import { resolveDayWindow } from "./meeting-contents.js";
import type { NotionClient } from "./notion/client.js";

export interface MeetingAugmentationSection {
  title: string;
  content: string;
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
  const pages = await deps.notion.listMeetingPagePropertiesForWindow({
    dataSourceId: config.notion.databaseId,
    datePropertyName: config.mapping.due,
    start: day.start,
    end: day.end,
  });
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

    if (matches.length === 0) {
      result.unmatched.push(section);
      continue;
    }

    if (matches.length > 1) {
      result.ambiguous.push({
        title: section.title,
        pageIds: matches.map((page) => page.id),
      });
      continue;
    }

    const page = matches[0];
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
  const titleLookup = new Map(knownTitles.map((title) => [normalizeMeetingTitle(title), title]));
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
    });
    currentLines = [];
  };

  for (const rawLine of input.replace(/\r\n/g, "\n").split("\n")) {
    const normalizedLineTitle = normalizeMeetingTitle(stripTitleMarkup(rawLine));

    if (titleLookup.has(normalizedLineTitle)) {
      flush();
      currentTitle = titleLookup.get(normalizedLineTitle);
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

export function normalizeMeetingTitle(title: string): string {
  return title
    .normalize("NFKC")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
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

function listPageTitles(config: NolendarConfig, pages: NotionMeetingPageProperties[]): string[] {
  return pages
    .map((page) => page.properties[config.mapping.title])
    .filter((title): title is string => typeof title === "string" && Boolean(title.trim()));
}

function stripTitleMarkup(line: string): string {
  return line
    .trim()
    .replace(/^#{1,6}\s+/, "")
    .replace(/^\*\*(.+)\*\*$/, "$1")
    .replace(/^__(.+)__$/, "$1")
    .trim();
}
