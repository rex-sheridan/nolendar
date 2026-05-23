import type { Clock } from "./clock.js";
import { systemClock } from "./clock.js";
import type { LookaheadWindow, NolendarConfig } from "./domain/config.js";
import type { Meeting } from "./domain/meeting.js";
import { resolveWindow, type CalendarWindow } from "./list.js";
import type { NotionClient } from "./notion/client.js";
import { type SyncOptions, type SyncResult, syncMeetingsToNotion } from "./sync.js";
import {
  loadSyncState as defaultLoadSyncState,
  saveSyncState as defaultSaveSyncState,
  type SyncState,
} from "./storage/sync-state.js";

export interface DeltaMeetingSource {
  listMeetingChanges(args: {
    calendar: NolendarConfig["calendars"][number];
    window: CalendarWindow;
    deltaLink?: string;
  }): Promise<{
    meetings: Meeting[];
    removedEventIds: string[];
    deltaLink?: string;
  }>;
}

export interface DeltaSyncDependencies {
  meetingSource: DeltaMeetingSource;
  loadState?: typeof defaultLoadSyncState;
  saveState?: typeof defaultSaveSyncState;
  clock?: Clock;
  syncOptions?: SyncOptions;
}

export async function syncCalendarChangesToNotion(
  config: NolendarConfig,
  notion: NotionClient,
  deps: DeltaSyncDependencies,
): Promise<SyncResult> {
  const clock = deps.clock ?? systemClock;
  const loadState = deps.loadState ?? defaultLoadSyncState;
  const saveState = deps.saveState ?? defaultSaveSyncState;
  const syncOptions = deps.syncOptions ?? {};
  const lookahead = config.sync.lookahead;
  const window = resolveWindow(lookahead, clock);
  const existingState = await loadState(config.sync.statePath);
  const nextCalendars = { ...existingState.calendars };
  const meetingsByCalendar: Meeting[][] = [];
  let archived = 0;

  for (const calendar of config.calendars) {
    const saved = existingState.calendars[calendar.id];
    const deltaLink = canReuseDelta(saved, lookahead, window) ? saved.deltaLink : undefined;
    const result = await deps.meetingSource.listMeetingChanges({
      calendar,
      window,
      deltaLink,
    });

    for (const removedEventId of result.removedEventIds) {
      const existing = await notion.findPageByEventId({
        dataSourceId: config.notion.databaseId,
        eventIdPropertyName: config.mapping.eventId,
        changeKeyPropertyName: config.mapping.changeKey,
        eventId: removedEventId,
      });

      if (!existing) {
        continue;
      }

      archived += 1;
      if (!syncOptions.dryRun) {
        await notion.archivePage(existing.id);
      }
    }

    if (result.deltaLink) {
      nextCalendars[calendar.id] = {
        lookahead,
        window,
        deltaLink: result.deltaLink,
        updatedAt: clock.now().toISOString(),
      };
    }

    meetingsByCalendar.push(result.meetings);
  }

  const meetings = meetingsByCalendar
    .flat()
    .sort((left, right) => left.start.localeCompare(right.start) || left.title.localeCompare(right.title));
  const syncResult = await syncMeetingsToNotion(config, meetings, notion, syncOptions);

  if (!syncResult.dryRun) {
    await saveState(config.sync.statePath, {
      version: 1,
      calendars: nextCalendars,
    });
  }

  return {
    ...syncResult,
    archived: syncResult.archived + archived,
  };
}

function canReuseDelta(
  saved: SyncState["calendars"][string] | undefined,
  lookahead: LookaheadWindow,
  window: CalendarWindow,
): saved is SyncState["calendars"][string] {
  return (
    saved !== undefined &&
    saved.lookahead === lookahead &&
    saved.window.start === window.start &&
    saved.window.end === window.end
  );
}
