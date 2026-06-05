import { Readable } from "node:stream";

import { afterEach, describe, expect, it, vi } from "vitest";

import { createCli, runCli } from "../src/cli.js";
import type { NolendarConfig } from "../src/domain/config.js";

afterEach(() => {
  vi.unstubAllEnvs();
});

describe("cli", () => {
  it("registers the expected commands", () => {
    const cli = createCli({
      stdout: { log: vi.fn() },
      stderr: { error: vi.fn() },
    });

    const commandNames = cli.commands.map((command) => command.name());

    expect(commandNames).toEqual([
      "list-calendars",
      "list",
      "meetings",
      "augment",
      "validate-config",
      "validate-notion",
      "print-notion-schema",
      "sync",
      "finalize-templates",
      "init",
      "wizard",
    ]);
  });

  it("returns a non-zero exit code for invalid calendar discovery tenant values", async () => {
    const stdout = { log: vi.fn() };
    const stderr = { error: vi.fn() };

    const exitCode = await runCli(["node", "nolendar", "list-calendars", "--tenant", "invalid"], {
      stdout,
      stderr,
    });

    expect(exitCode).toBe(1);
    expect(stderr.error).toHaveBeenCalledWith("`--tenant` must be one of: common, organizations, consumers.");
  });

  it("uses nolendar.yml as Microsoft auth config for calendar discovery when present", async () => {
    const stdout = { log: vi.fn() };
    const stderr = { error: vi.fn() };
    const loadConfigMock = vi.fn(async () => ({
      microsoft: { tenant: "organizations", authMode: "auth_code" },
    }));

    const exitCode = await runCli(["node", "nolendar", "list-calendars", "--tenant", "invalid"], {
      stdout,
      stderr,
      loadConfig: loadConfigMock as never,
    });

    expect(exitCode).toBe(1);
    expect(loadConfigMock).toHaveBeenCalledWith("nolendar.yml");
    expect(stderr.error).toHaveBeenCalledWith("`--tenant` must be one of: common, organizations, consumers.");
  });

  it("returns a non-zero exit code for invalid calendar discovery auth modes", async () => {
    const stdout = { log: vi.fn() };
    const stderr = { error: vi.fn() };

    const exitCode = await runCli(["node", "nolendar", "list-calendars", "--auth-mode", "password"], {
      stdout,
      stderr,
    });

    expect(exitCode).toBe(1);
    expect(stderr.error).toHaveBeenCalledWith(
      "`--auth-mode` must be one of: device_code, interactive_browser, auth_code.",
    );
  });

  it("returns a non-zero exit code for invalid lookahead values", async () => {
    const stdout = { log: vi.fn() };
    const stderr = { error: vi.fn() };

    const exitCode = await runCli(
      ["node", "nolendar", "list", "--config", "tests/fixtures/valid-config.yml", "--lookahead", "30x"],
      {
        stdout,
        stderr,
      },
    );

    expect(exitCode).toBe(1);
    expect(stderr.error).toHaveBeenCalledWith(
      "`--lookahead` must be `today` or a relative range like `12h`, `5d`, `2w`, or `3m`.",
    );
  });

  it("accepts verbose output for meeting augmentation", async () => {
    const stdout = { log: vi.fn() };
    const stderr = { error: vi.fn() };
    const config: NolendarConfig = {
      microsoft: { tenant: "common", authMode: "device_code" },
      notion: {
        databaseId: "meetings-id",
        augmentation: {
          delimiter: "%%MEETING%%",
        },
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
    const appendMarkdownToPage = vi.fn(async () => undefined);

    const exitCode = await runCli(
      ["node", "nolendar", "augment", "--heading", "Follow-ups", "--day", "2026-05-22", "--verbose"],
      {
        stdout,
        stderr,
        stdin: Readable.from(["%%MEETING%%\nPlanning\nSend summary."]),
        loadConfig: vi.fn(async () => config),
        buildNotionClient: () =>
          ({
            listMeetingPagePropertiesForWindow: vi.fn(async () => [
              {
                id: "page-1",
                url: "https://notion.so/page-1",
                properties: {
                  Name: "Planning",
                  Due: {
                    start: "2026-05-22T13:00:00.000Z",
                  },
                },
              },
            ]),
            appendMarkdownToPage,
          }) as never,
      },
    );

    expect(exitCode).toBe(0);
    expect(appendMarkdownToPage).toHaveBeenCalledWith({
      pageId: "page-1",
      heading: "Follow-ups",
      content: "Send summary.",
    });
    expect(stdout.log).toHaveBeenCalledWith(
      'Meeting augmentation summary: day=2026-05-22, heading="Follow-ups", matched=1, unmatched=0, ambiguous=0, empty=0, dryRun=false',
    );
    expect(stdout.log).toHaveBeenCalledWith("  matched: Planning -> page-1 (https://notion.so/page-1) appended");
    expect(stderr.error).not.toHaveBeenCalled();
  });

  it("returns a non-zero exit code for invalid finalize delay values", async () => {
    const stdout = { log: vi.fn() };
    const stderr = { error: vi.fn() };

    const exitCode = await runCli(
      ["node", "nolendar", "sync", "--config", "tests/fixtures/valid-config.yml", "--finalize-delay-ms", "-1"],
      {
        stdout,
        stderr,
      },
    );

    expect(exitCode).toBe(1);
    expect(stderr.error).toHaveBeenCalledWith("`--finalize-delay-ms` must be a non-negative integer.");
  });

  it("returns a non-zero exit code when auth configuration is incomplete", async () => {
    const stdout = { log: vi.fn() };
    const stderr = { error: vi.fn() };
    vi.stubEnv("MICROSOFT_CLIENT_ID", undefined);
    vi.stubEnv("MICROSOFT_CLIENT_SECRET", undefined);
    vi.stubEnv("MICROSOFT_ACCESS_TOKEN", undefined);

    const exitCode = await runCli(
      ["node", "nolendar", "list", "--config", "tests/fixtures/auth-code-config.yml"],
      {
        stdout,
        stderr,
      },
    );

    expect(exitCode).toBe(1);
    expect(stderr.error).toHaveBeenCalled();
  });

  it("prints the detected meeting and People data source schemas", async () => {
    const stdout = { log: vi.fn() };
    const stderr = { error: vi.fn() };
    const config: NolendarConfig = {
      microsoft: { tenant: "common", authMode: "device_code" },
      notion: {
        databaseId: "meetings-id",
        peopleDataSource: {
          databaseId: "people-id",
          nameProperty: "Name",
          emailProperty: "Email Address",
        },
      },
      calendars: [{ id: "primary" }],
      filters: {
        ignoreDeclined: true,
        requireAttendees: false,
        ignorePersonal: false,
        ignoreOptionalAttendance: false,
      },
      mapping: {
        title: "Task name",
        due: "Due",
        eventId: "Outlook Event ID",
        changeKey: "Outlook ChangeKey",
        participants: "Participants",
      },
      sync: {
        lookahead: "today",
        statePath: "/tmp/.nolendar/state.json",
      },
    };
    const loadConfigMock = vi.fn(async () => config);
    const retrieveDataSource = vi
      .fn()
      .mockResolvedValueOnce({
        id: "meetings-id",
        title: "Meetings",
        properties: {
          "Task name": { id: "title", name: "Task name", type: "title" },
          Participants: { id: "participants", name: "Participants", type: "relation" },
          Due: { id: "due", name: "Due", type: "date" },
        },
      })
      .mockResolvedValueOnce({
        id: "people-id",
        title: "People",
        properties: {
          Name: { id: "name", name: "Name", type: "title" },
          "Email Address": { id: "email", name: "Email Address", type: "email" },
        },
      });

    const exitCode = await runCli(["node", "nolendar", "print-notion-schema"], {
      stdout,
      stderr,
      loadConfig: loadConfigMock,
      buildNotionClient: () =>
        ({
          retrieveDataSource,
        }) as never,
    });

    expect(exitCode).toBe(0);
    expect(retrieveDataSource).toHaveBeenCalledWith("meetings-id");
    expect(retrieveDataSource).toHaveBeenCalledWith("people-id");
    expect(stdout.log).toHaveBeenNthCalledWith(1, "Meeting data source: Meetings (meetings-id)");
    expect(stdout.log).toHaveBeenNthCalledWith(2, "  - Due: date");
    expect(stdout.log).toHaveBeenNthCalledWith(3, "  - Participants: relation");
    expect(stdout.log).toHaveBeenNthCalledWith(4, "  - Task name: title");
    expect(stdout.log).toHaveBeenNthCalledWith(5, "People data source: People (people-id)");
    expect(stdout.log).toHaveBeenNthCalledWith(6, "  - Email Address: email");
    expect(stdout.log).toHaveBeenNthCalledWith(7, "  - Name: title");
    expect(stderr.error).not.toHaveBeenCalled();
  });
});
