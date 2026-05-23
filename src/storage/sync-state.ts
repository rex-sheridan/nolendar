import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import type { LookaheadWindow } from "../domain/config.js";
import type { CalendarWindow } from "../list.js";

export interface CalendarSyncState {
  lookahead: LookaheadWindow;
  window: CalendarWindow;
  deltaLink: string;
  updatedAt: string;
}

export interface SyncState {
  version: 1;
  calendars: Record<string, CalendarSyncState>;
}

const EMPTY_STATE: SyncState = {
  version: 1,
  calendars: {},
};

export async function loadSyncState(statePath: string): Promise<SyncState> {
  try {
    const raw = await readFile(statePath, "utf8");
    const parsed = JSON.parse(raw) as Partial<SyncState>;

    return {
      version: 1,
      calendars: parsed.calendars ?? {},
    };
  } catch (error) {
    if (isNotFoundError(error)) {
      return EMPTY_STATE;
    }

    throw error;
  }
}

export async function saveSyncState(statePath: string, state: SyncState): Promise<void> {
  await mkdir(path.dirname(statePath), { recursive: true });
  await writeFile(statePath, JSON.stringify(state, null, 2), "utf8");
}

function isNotFoundError(error: unknown): error is NodeJS.ErrnoException {
  return typeof error === "object" && error !== null && "code" in error && error.code === "ENOENT";
}
