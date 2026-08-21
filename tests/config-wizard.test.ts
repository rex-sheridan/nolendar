import { mkdtemp, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { parse } from "yaml";

import { afterEach, describe, expect, it, vi } from "vitest";

import { runConfigWizard, type ConfigWizardPrompt } from "../src/config-wizard.js";
import type { MicrosoftConfig } from "../src/domain/config.js";

const tempDirs: string[] = [];

describe("runConfigWizard", () => {
  afterEach(async () => {
    await Promise.all(tempDirs.splice(0).map((tempDir) => rm(tempDir, { recursive: true, force: true })));
  });

  it("writes a complete config from discovered calendars and Notion data sources", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "nolendar-wizard-"));
    tempDirs.push(tempDir);
    const configPath = path.join(tempDir, "nolendar.yml");
    const ensureProperties = vi.fn(async () => undefined);

    const writtenPath = await runConfigWizard({
      configPath,
      prompt: new ScriptedPrompt({
        selects: [
          "common",
          "device_code",
          "meeting-id",
          "people-id",
          "Name",
          "Email Address",
          "Name",
          "Due",
          "__create__",
          "__create__",
          "__skip__",
          "__skip__",
          "__skip__",
          "__create__",
        ],
        multiselects: [["calendar-1"]],
        confirms: [true, true, false],
        asks: ["Outlook Event ID", "Outlook ChangeKey", "Participants"],
      }),
      stdout: { log: vi.fn() },
      listCalendars: async (_microsoft: MicrosoftConfig) => [
        {
          id: "calendar-1",
          name: "Work",
          isDefaultCalendar: true,
        },
      ],
      notion: {
        listDataSources: async () => [
          { id: "meeting-id", title: "Meetings" },
          { id: "people-id", title: "People" },
        ],
        retrieveDataSource: async (dataSourceId: string) => {
          if (dataSourceId === "people-id") {
            return {
              id: "people-id",
              title: "People",
              properties: {
                Name: { id: "name", name: "Name", type: "title" },
                "Email Address": { id: "email", name: "Email Address", type: "email" },
              },
            };
          }

          return {
            id: "meeting-id",
            title: "Meetings",
            properties: {
              Name: { id: "name", name: "Name", type: "title" },
              Due: { id: "due", name: "Due", type: "date" },
            },
          };
        },
        ensureProperties,
      } as never,
    });

    const written = parse(await readFile(configPath, "utf8"));

    expect(writtenPath).toBe(configPath);
    expect(written.calendars).toEqual([{ id: "calendar-1", name: "Work" }]);
    expect(written.notion.databaseId).toBe("meeting-id");
    expect(written.notion.peopleDataSource.databaseId).toBe("people-id");
    expect(written.mapping).toMatchObject({
      title: "Name",
      due: "Due",
      eventId: "Outlook Event ID",
      changeKey: "Outlook ChangeKey",
      participants: "Participants",
    });
    expect(ensureProperties).toHaveBeenCalledWith("meeting-id", [
      { name: "Outlook Event ID", type: "rich_text" },
      { name: "Outlook ChangeKey", type: "rich_text" },
      { name: "Participants", type: "relation", relationDataSourceId: "people-id" },
    ]);
  });

  it("enumerates data source templates when configuring a specific meeting page template", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "nolendar-wizard-"));
    tempDirs.push(tempDir);
    const configPath = path.join(tempDir, "nolendar.yml");
    const listDataSourceTemplates = vi.fn(async () => [
      { id: "template-1", name: "Plain meeting", isDefault: true },
      { id: "template-2", name: "Leadership review", isDefault: false },
    ]);

    await runConfigWizard({
      configPath,
      prompt: new ScriptedPrompt({
        selects: [
          "common",
          "device_code",
          "meeting-id",
          "Name",
          "Due",
          "Outlook Event ID",
          "Outlook ChangeKey",
          "__skip__",
          "__skip__",
          "__skip__",
          "notion",
          "template_id",
          "template-2",
          "none",
          "archive",
        ],
        multiselects: [
          ["calendar-1"],
          ["meeting_link", "calendar_event", "meeting_details", "notes", "action_items"],
        ],
        confirms: [false, true, false],
        asks: ["", "meeting", "", ""],
      }),
      stdout: { log: vi.fn() },
      listCalendars: async () => [
        {
          id: "calendar-1",
          name: "Work",
          isDefaultCalendar: true,
        },
      ],
      notion: {
        listDataSources: async () => [{ id: "meeting-id", title: "Meetings" }],
        listDataSourceTemplates,
        retrieveDataSource: async () => ({
          id: "meeting-id",
          title: "Meetings",
          properties: {
            Name: { id: "name", name: "Name", type: "title" },
            Due: { id: "due", name: "Due", type: "date" },
            "Outlook Event ID": { id: "event-id", name: "Outlook Event ID", type: "rich_text" },
            "Outlook ChangeKey": { id: "change-key", name: "Outlook ChangeKey", type: "rich_text" },
          },
        }),
        ensureProperties: vi.fn(async () => undefined),
      } as never,
    });

    const written = parse(await readFile(configPath, "utf8"));

    expect(listDataSourceTemplates).toHaveBeenCalledWith("meeting-id");
    expect(written.notion.dataSourceTemplate).toEqual({
      type: "template_id",
      templateId: "template-2",
      timezone: Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC",
    });
  });
});

class ScriptedPrompt implements ConfigWizardPrompt {
  constructor(
    private readonly script: {
      asks?: string[];
      confirms?: boolean[];
      selects?: string[];
      multiselects?: string[][];
    },
  ) {}

  async ask(_question: string, defaultValue?: string): Promise<string> {
    const answer = this.shift(this.script.asks, "ask");

    return answer || defaultValue || "";
  }

  async confirm(): Promise<boolean> {
    return this.shift(this.script.confirms, "confirm");
  }

  async select<T extends string>(): Promise<T> {
    return this.shift(this.script.selects, "select") as T;
  }

  async multiselect<T extends string>(): Promise<T[]> {
    return this.shift(this.script.multiselects, "multiselect") as T[];
  }

  private shift<T>(values: T[] | undefined, label: string): T {
    const next = values?.shift();
    if (next === undefined) {
      throw new Error(`No scripted ${label} response left.`);
    }

    return next;
  }
}
