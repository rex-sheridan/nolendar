import { describe, expect, it, vi } from "vitest";

import type { NolendarConfig } from "../src/domain/config.js";
import { importMeetingAugmentation, parseMeetingAugmentationSections } from "../src/meeting-augmentation.js";

const CONFIG: NolendarConfig = {
  microsoft: {
    tenant: "common",
    authMode: "device_code",
  },
  notion: {
    databaseId: "meetings",
  },
  calendars: [{ id: "primary" }],
  filters: {
    ignoreDeclined: true,
    requireAttendees: false,
    ignorePersonal: false,
    ignoreOptionalAttendance: false,
  },
  mapping: {
    title: "Name",
    due: "Due",
    eventId: "Outlook Event ID",
    changeKey: "Outlook ChangeKey",
  },
  sync: {
    lookahead: "today",
    statePath: "/tmp/.nolendar/state.json",
  },
};

describe("parseMeetingAugmentationSections", () => {
  it("splits agent output by known meeting titles and preserves markdown content", () => {
    expect(
      parseMeetingAugmentationSections(
        [
          "Preamble the agent may include.",
          "# Planning",
          "- Decide launch scope",
          "## Key decision",
          "Ship date",
          "**Design Review**",
          "Focus on open comments.",
        ].join("\n"),
        ["Planning", "Design Review"],
      ),
    ).toEqual([
      {
        title: "Planning",
        content: "- Decide launch scope\n## Key decision\nShip date",
      },
      {
        title: "Design Review",
        content: "Focus on open comments.",
      },
    ]);
  });

  it("recognizes numbered markdown-link headings as meeting boundaries", () => {
    expect(
      parseMeetingAugmentationSections(
        [
          "# Executive Summary",
          "Prep for [Cloud Vulnerability Management](https://example.com): this is not the section yet.",
          "",
          "## 1) [Cloud Vulnerability Management](https://example.com)",
          "* **Time:** today, 1:30 PM",
          "### Suggested questions",
          "- Which risks matter?",
          "",
          "## 2) [Dennis and Rex](https://example.com)",
          "- Discuss load.",
        ].join("\n"),
        ["Cloud Vulnerability Management", "Dennis and Rex"],
      ),
    ).toEqual([
      {
        title: "Cloud Vulnerability Management",
        content: "* **Time:** today, 1:30 PM\n### Suggested questions\n- Which risks matter?",
        time: {
          startMinutes: 810,
        },
      },
      {
        title: "Dennis and Rex",
        content: "- Discuss load.",
        time: undefined,
      },
    ]);
  });

  it("recognizes bracketed meeting result labels as meeting boundaries", () => {
    expect(
      parseMeetingAugmentationSections(
        [
          "[Clinton and Rex | Meeting]",
          "Content from a search result style title.",
          "\\[SRCI Standup | Meeting\\]",
          "Escaped bracket title.",
        ].join("\n"),
        ["Clinton and Rex", "SRCI Standup"],
      ),
    ).toEqual([
      {
        title: "Clinton and Rex",
        content: "Content from a search result style title.",
      },
      {
        title: "SRCI Standup",
        content: "Escaped bracket title.",
      },
    ]);
  });

  it("matches title boundaries across punctuation and suffix differences", () => {
    expect(
      parseMeetingAugmentationSections(
        [
          "## 1) [Daily Standup Infrastructure / Dev Ops / Cloud | Meeting](https://example.com)",
          "Cloud standup notes.",
          "## 2) [Planning | Meeting](https://example.com)",
          "Planning notes.",
        ].join("\n"),
        ["Daily Standup - Infrastructure / Dev Ops / Cloud", "Planning"],
      ),
    ).toEqual([
      {
        title: "Daily Standup - Infrastructure / Dev Ops / Cloud",
        content: "Cloud standup notes.",
      },
      {
        title: "Planning",
        content: "Planning notes.",
      },
    ]);
  });
});

describe("importMeetingAugmentation", () => {
  it("matches sections to Notion pages by title within the selected day", async () => {
    const appendMarkdownToPage = vi.fn(async () => undefined);
    const listMeetingPagePropertiesForWindow = vi.fn(async () => [
      {
        id: "page-1",
        url: "https://notion.so/page-1",
        properties: {
          Name: "Planning",
        },
      },
      {
        id: "page-2",
        properties: {
          Name: "Design Review",
        },
      },
    ]);

    const result = await importMeetingAugmentation(
      CONFIG,
      "Planning\nDecide launch scope.",
      {
        day: "2026-05-22",
        heading: "Decisions",
      },
      {
        notion: {
          listMeetingPagePropertiesForWindow,
          appendMarkdownToPage,
        } as never,
      },
    );

    expect(listMeetingPagePropertiesForWindow).toHaveBeenCalledWith({
      dataSourceId: "meetings",
      datePropertyName: "Due",
      start: "2026-05-22T00:00:00.000Z",
      end: "2026-05-23T00:00:00.000Z",
    });
    expect(appendMarkdownToPage).toHaveBeenCalledWith({
      pageId: "page-1",
      heading: "Decisions",
      content: "Decide launch scope.",
    });
    expect(result.matched).toEqual([
      {
        title: "Planning",
        pageId: "page-1",
        url: "https://notion.so/page-1",
        appended: true,
      },
    ]);
    expect(result.unmatched).toEqual([]);
  });

  it("uses section time as a tie-breaker when title matches multiple pages", async () => {
    const appendMarkdownToPage = vi.fn(async () => undefined);

    const result = await importMeetingAugmentation(
      CONFIG,
      [
        "## [Cloud Vulnerability Management](https://example.com)",
        "* **Time:** today, **1:30 PM–2:00 PM**",
        "- Which vulnerabilities are highest risk?",
      ].join("\n"),
      {
        day: "2026-06-04",
        heading: "Preparation",
      },
      {
        notion: {
          listMeetingPagePropertiesForWindow: vi.fn(async () => [
            {
              id: "morning-page",
              properties: {
                Name: "Cloud Vulnerability Management",
                Due: {
                  start: "2026-06-04T09:30:00.000-04:00",
                  end: "2026-06-04T10:00:00.000-04:00",
                },
              },
            },
            {
              id: "afternoon-page",
              url: "https://notion.so/afternoon-page",
              properties: {
                Name: "Cloud Vulnerability Management",
                Due: {
                  start: "2026-06-04T13:30:00.000-04:00",
                  end: "2026-06-04T14:00:00.000-04:00",
                },
              },
            },
          ]),
          appendMarkdownToPage,
        } as never,
      },
    );

    expect(appendMarkdownToPage).toHaveBeenCalledWith({
      pageId: "afternoon-page",
      heading: "Preparation",
      content: "* **Time:** today, **1:30 PM–2:00 PM**\n- Which vulnerabilities are highest risk?",
    });
    expect(result.matched).toEqual([
      {
        title: "Cloud Vulnerability Management",
        pageId: "afternoon-page",
        url: "https://notion.so/afternoon-page",
        appended: true,
      },
    ]);
    expect(result.ambiguous).toEqual([]);
  });

  it("reports ambiguous titles and leaves dry-run matches unmodified", async () => {
    const appendMarkdownToPage = vi.fn(async () => undefined);

    const result = await importMeetingAugmentation(
      CONFIG,
      "Planning\nDecide launch scope.",
      {
        day: "2026-05-22",
        dryRun: true,
        heading: "Risks",
      },
      {
        notion: {
          listMeetingPagePropertiesForWindow: vi.fn(async () => [
            { id: "page-1", properties: { Name: "Planning" } },
            { id: "page-2", properties: { Name: "Planning" } },
          ]),
          appendMarkdownToPage,
        } as never,
      },
    );

    expect(appendMarkdownToPage).not.toHaveBeenCalled();
    expect(result.matched).toEqual([]);
    expect(result.ambiguous).toEqual([
      {
        title: "Planning",
        pageIds: ["page-1", "page-2"],
      },
    ]);
  });

  it("keeps a title ambiguous when the section has no usable time tie-breaker", async () => {
    const result = await importMeetingAugmentation(
      CONFIG,
      "Planning\nDecide launch scope.",
      {
        day: "2026-05-22",
        dryRun: true,
        heading: "Risks",
      },
      {
        notion: {
          listMeetingPagePropertiesForWindow: vi.fn(async () => [
            {
              id: "page-1",
              properties: {
                Name: "Planning",
                Due: { start: "2026-05-22T09:00:00.000-04:00" },
              },
            },
            {
              id: "page-2",
              properties: {
                Name: "Planning",
                Due: { start: "2026-05-22T13:00:00.000-04:00" },
              },
            },
          ]),
        } as never,
      },
    );

    expect(result.ambiguous).toEqual([
      {
        title: "Planning",
        pageIds: ["page-1", "page-2"],
      },
    ]);
  });

  it("rejects empty headings", async () => {
    await expect(
      importMeetingAugmentation(
        CONFIG,
        "Planning\nContent",
        {
          day: "2026-05-22",
          heading: " ",
        },
        {
          notion: {
            listMeetingPagePropertiesForWindow: vi.fn(async () => []),
          } as never,
        },
      ),
    ).rejects.toThrow("`--heading` must not be empty.");
  });
});
