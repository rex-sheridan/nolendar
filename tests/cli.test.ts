import { describe, expect, it, vi } from "vitest";

import { createCli, runCli } from "../src/cli.js";

describe("cli", () => {
  it("registers the expected milestone 1 commands", () => {
    const cli = createCli({
      stdout: { log: vi.fn() },
      stderr: { error: vi.fn() },
    });

    const commandNames = cli.commands.map((command) => command.name());

    expect(commandNames).toEqual(["list", "validate-config"]);
  });

  it("returns a non-zero exit code for invalid lookahead values", async () => {
    const stdout = { log: vi.fn() };
    const stderr = { error: vi.fn() };

    const exitCode = await runCli(
      ["node", "nolendar", "list", "--config", "tests/fixtures/valid-config.yml", "--lookahead", "30d"],
      {
        stdout,
        stderr,
      },
    );

    expect(exitCode).toBe(1);
    expect(stderr.error).toHaveBeenCalledWith("`--lookahead` must be one of: today, 24h, 7d.");
  });

  it("prints configured calendars for the list command", async () => {
    const stdout = { log: vi.fn() };
    const stderr = { error: vi.fn() };

    const exitCode = await runCli(["node", "nolendar", "list", "--config", "tests/fixtures/valid-config.yml"], {
      stdout,
      stderr,
    });

    expect(exitCode).toBe(0);
    expect(stdout.log).toHaveBeenCalledWith("Meetings for today from 2 configured calendar(s):");
    expect(stdout.log).toHaveBeenCalledWith("- Primary");
    expect(stdout.log).toHaveBeenCalledWith("- Team");
  });
});
