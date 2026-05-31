import { mkdtemp, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";

import { EnvFileError, loadLocalEnvFile, parseEnvFile } from "../src/env.js";

describe("parseEnvFile", () => {
  it("parses common env file assignment forms", () => {
    expect(
      Object.fromEntries(
        parseEnvFile(`
# local credentials
MICROSOFT_CLIENT_ID=client-id
export NOTION_TOKEN="secret token"
SINGLE_QUOTED='literal value'
INLINE_COMMENT=value # comment
EMPTY=
MULTILINE="line one\\nline two"
`),
      ),
    ).toEqual({
      MICROSOFT_CLIENT_ID: "client-id",
      NOTION_TOKEN: "secret token",
      SINGLE_QUOTED: "literal value",
      INLINE_COMMENT: "value",
      EMPTY: "",
      MULTILINE: "line one\nline two",
    });
  });

  it("rejects invalid assignment lines", () => {
    expect(() => parseEnvFile("not an assignment")).toThrowError(
      new EnvFileError(".env:1 must be a KEY=value assignment."),
    );
  });
});

describe("loadLocalEnvFile", () => {
  it("loads values without overriding existing environment", async () => {
    const directory = await mkdtemp(path.join(os.tmpdir(), "nolendar-env-"));
    const envPath = path.join(directory, ".env");
    const env: NodeJS.ProcessEnv = {
      NOTION_TOKEN: "from-shell",
    };

    await writeFile(envPath, "NOTION_TOKEN=from-file\nMICROSOFT_CLIENT_ID=client-id\n", "utf8");

    await loadLocalEnvFile(envPath, env);

    expect(env).toEqual({
      NOTION_TOKEN: "from-shell",
      MICROSOFT_CLIENT_ID: "client-id",
    });
  });

  it("ignores a missing env file", async () => {
    const env: NodeJS.ProcessEnv = {};

    await loadLocalEnvFile("/tmp/nolendar-missing-env-file", env);

    expect(env).toEqual({});
  });
});
