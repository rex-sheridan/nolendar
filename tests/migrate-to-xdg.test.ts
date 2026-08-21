import { execFile } from "node:child_process";
import { mkdir, mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";

import { afterEach, describe, expect, it } from "vitest";

const execFileAsync = promisify(execFile);
const tempDirectories: string[] = [];
const migrationScript = path.resolve("bin/migrate-to-xdg");

describe("migrate-to-xdg", () => {
  afterEach(async () => {
    await Promise.all(tempDirectories.splice(0).map((directory) => rm(directory, { recursive: true, force: true })));
  });

  it("copies legacy files to XDG locations without changing the source files", async () => {
    const fixture = await createFixture();

    const { stdout } = await execFileAsync(migrationScript, ["--legacy-dir", fixture.legacyDirectory], {
      env: fixture.env,
    });

    const configPath = path.join(fixture.xdgConfigHome, "nolendar", "config.yml");
    const envPath = path.join(fixture.xdgConfigHome, "nolendar", "env");
    const statePath = path.join(fixture.xdgStateHome, "nolendar", "state.json");
    const tokenCachePath = path.join(fixture.xdgDataHome, "nolendar", "msal-cache.json");

    expect(await readFile(configPath, "utf8")).toBe("sync:\n  lookahead: today\nnotion:\n  databaseId: db-1\n");
    expect(await readFile(envPath, "utf8")).toBe("NOTION_TOKEN=secret\n");
    expect(await readFile(statePath, "utf8")).toBe('{"calendars":{}}\n');
    expect(await readFile(tokenCachePath, "utf8")).toBe('{"token":"cached"}\n');
    expect((await stat(envPath)).mode & 0o777).toBe(0o600);
    expect((await stat(path.dirname(configPath))).mode & 0o777).toBe(0o700);
    expect(await readFile(path.join(fixture.legacyDirectory, "nolendar.yml"), "utf8")).toContain(
      "statePath: .nolendar/state.json",
    );
    expect(stdout).toContain("Migration complete");
  });

  it("is repeatable when the migrated files already match", async () => {
    const fixture = await createFixture();

    await execFileAsync(migrationScript, ["--legacy-dir", fixture.legacyDirectory], { env: fixture.env });
    const { stdout } = await execFileAsync(migrationScript, ["--legacy-dir", fixture.legacyDirectory], {
      env: fixture.env,
    });

    expect(stdout).toContain("already up to date");
    expect(stdout).toContain("Migration complete");
  });

  it("refuses to overwrite a conflicting destination", async () => {
    const fixture = await createFixture();
    const configDirectory = path.join(fixture.xdgConfigHome, "nolendar");
    await mkdir(configDirectory, { recursive: true });
    await writeFile(path.join(configDirectory, "config.yml"), "existing: true\n", "utf8");

    await expect(
      execFileAsync(migrationScript, ["--legacy-dir", fixture.legacyDirectory], { env: fixture.env }),
    ).rejects.toMatchObject({
      stderr: expect.stringContaining("Refusing to overwrite"),
    });
  });

  it("reports actions without writing during a dry run", async () => {
    const fixture = await createFixture();

    const { stdout } = await execFileAsync(migrationScript, ["--dry-run", "--legacy-dir", fixture.legacyDirectory], {
      env: fixture.env,
    });

    expect(stdout).toContain("Would migrate");
    await expect(stat(path.join(fixture.xdgConfigHome, "nolendar"))).rejects.toMatchObject({ code: "ENOENT" });
  });
});

async function createFixture(): Promise<{
  env: NodeJS.ProcessEnv;
  legacyDirectory: string;
  xdgConfigHome: string;
  xdgDataHome: string;
  xdgStateHome: string;
}> {
  const root = await mkdtemp(path.join(os.tmpdir(), "nolendar-migrate-"));
  tempDirectories.push(root);
  const home = path.join(root, "home");
  const legacyDirectory = path.join(root, "checkout");
  const xdgConfigHome = path.join(root, "xdg-config");
  const xdgDataHome = path.join(root, "xdg-data");
  const xdgStateHome = path.join(root, "xdg-state");

  await mkdir(path.join(legacyDirectory, ".nolendar"), { recursive: true });
  await mkdir(path.join(home, ".nolendar"), { recursive: true });
  await writeFile(
    path.join(legacyDirectory, "nolendar.yml"),
    "sync:\n  lookahead: today\n  statePath: .nolendar/state.json\nnotion:\n  databaseId: db-1\n",
    "utf8",
  );
  await writeFile(path.join(legacyDirectory, ".env"), "NOTION_TOKEN=secret\n", "utf8");
  await writeFile(path.join(legacyDirectory, ".nolendar", "state.json"), '{"calendars":{}}\n', "utf8");
  await writeFile(path.join(home, ".nolendar", "msal-cache.json"), '{"token":"cached"}\n', "utf8");

  return {
    env: {
      HOME: home,
      PATH: process.env.PATH,
      XDG_CONFIG_HOME: xdgConfigHome,
      XDG_DATA_HOME: xdgDataHome,
      XDG_STATE_HOME: xdgStateHome,
    },
    legacyDirectory,
    xdgConfigHome,
    xdgDataHome,
    xdgStateHome,
  };
}
