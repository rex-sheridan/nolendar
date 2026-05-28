import { describe, expect, it, vi } from "vitest";

import { createCli, runCli } from "../src/cli.js";
import type { NolendarConfig } from "../src/domain/config.js";

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
      "validate-config",
      "validate-notion",
      "print-notion-schema",
      "sync",
      "finalize-templates",
      "init",
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
