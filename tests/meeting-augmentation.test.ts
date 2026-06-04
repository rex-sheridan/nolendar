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
