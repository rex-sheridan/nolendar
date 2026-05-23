import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { afterEach, describe, expect, it } from "vitest";

import { InitConfigError, writeDefaultConfig } from "../src/init-config.js";

const tempDirs: string[] = [];

describe("writeDefaultConfig", () => {
  afterEach(async () => {
    await Promise.all(
      tempDirs.splice(0).map(async (tempDir) => {
        await import("node:fs/promises").then(({ rm }) => rm(tempDir, { recursive: true, force: true }));
      }),
    );
  });

  it("writes a starter nolendar config", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "nolendar-init-"));
    tempDirs.push(tempDir);
    const configPath = path.join(tempDir, "nolendar.yml");

    const writtenPath = await writeDefaultConfig(configPath);
    const contents = await readFile(configPath, "utf8");

    expect(writtenPath).toBe(configPath);
    expect(contents).toContain("microsoft:");
    expect(contents).toContain("databaseId: your_notion_data_source_id");
    expect(contents).toContain("id: primary");
  });

  it("fails if the config already exists without force", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "nolendar-init-"));
    tempDirs.push(tempDir);
    const configPath = path.join(tempDir, "nolendar.yml");
    await writeFile(configPath, "existing: true\n", "utf8");

    await expect(writeDefaultConfig(configPath)).rejects.toThrowError(
      new InitConfigError(`Config file already exists at ${configPath}. Use --force to overwrite it.`),
    );
  });

  it("overwrites an existing config when force is true", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "nolendar-init-"));
    tempDirs.push(tempDir);
    const configPath = path.join(tempDir, "nolendar.yml");
    await writeFile(configPath, "existing: true\n", "utf8");

    await writeDefaultConfig(configPath, { force: true });
    const contents = await readFile(configPath, "utf8");

    expect(contents).toContain("microsoft:");
    expect(contents).not.toContain("existing: true");
  });
});
