import { describe, expect, it, vi } from "vitest";

import { createCli, runCli } from "../src/cli.js";

describe("cli", () => {
  it("registers the expected commands", () => {
    const cli = createCli({
      stdout: { log: vi.fn() },
      stderr: { error: vi.fn() },
    });

    const commandNames = cli.commands.map((command) => command.name());

    expect(commandNames).toEqual(["list", "validate-config", "validate-notion", "sync", "init"]);
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
});
