import { mkdtemp, readFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { loadSyncState, saveSyncState } from "../src/storage/sync-state.js";

describe("sync-state storage", () => {
  it("returns an empty state when the file does not exist", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "nolendar-state-"));
    const statePath = path.join(tempDir, "missing.json");

    await expect(loadSyncState(statePath)).resolves.toEqual({
      version: 1,
      calendars: {},
    });
  });

  it("persists and reloads calendar delta state", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "nolendar-state-"));
    const statePath = path.join(tempDir, ".nolendar", "state.json");
    const expected = {
      version: 1 as const,
      calendars: {
        primary: {
          lookahead: "today" as const,
          window: {
            start: "2026-05-23T00:00:00.000Z",
            end: "2026-05-24T00:00:00.000Z",
          },
          deltaLink: "https://graph.microsoft.com/v1.0/me/calendars/primary/calendarView/delta?$deltatoken=abc",
          updatedAt: "2026-05-23T12:00:00.000Z",
        },
      },
    };

    await saveSyncState(statePath, expected);

    await expect(loadSyncState(statePath)).resolves.toEqual(expected);
    await expect(readFile(statePath, "utf8")).resolves.toContain("\"deltaLink\"");
  });
});
