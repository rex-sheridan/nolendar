import type { Clock } from "./clock.js";
import { systemClock } from "./clock.js";
import type { LookaheadWindow, NolendarConfig } from "./domain/config.js";
import type { Meeting } from "./domain/meeting.js";
import { resolveWindow, type CalendarWindow } from "./list.js";
import type { NotionClient } from "./notion/client.js";
import {
  applyCanceledMeetingAction,
  getCanceledMeetingAction,
  type SyncOptions,
  type SyncResult,
  syncMeetingsToNotion,
} from "./sync.js";
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
  const currentWindowMeetingsByCalendar: Meeting[][] = [];
  let archived = 0;
  let updated = 0;
  const canReconcileMissingMeetings = canListMeetingPagesForWindow(notion);

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
        syncOptions.onDecision?.(`sync decision: removedEventId=${removedEventId} pageId=- decision=skip_removed_missing_page`);
        continue;
      }

      if (getCanceledMeetingAction(config).action === "archive") {
        archived += 1;
        syncOptions.onDecision?.(`sync decision: removedEventId=${removedEventId} pageId=${existing.id} decision=archive_removed`);
      } else {
        updated += 1;
        syncOptions.onDecision?.(`sync decision: removedEventId=${removedEventId} pageId=${existing.id} decision=set_status_removed`);
      }

      if (!syncOptions.dryRun) {
        await applyCanceledMeetingAction(config, notion, existing.id);
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
    currentWindowMeetingsByCalendar.push(
      deltaLink && canReconcileMissingMeetings
        ? (await deps.meetingSource.listMeetingChanges({
            calendar,
            window,
          })).meetings
        : result.meetings,
    );
  }

  const meetings = meetingsByCalendar
    .flat()
    .sort((left, right) => left.start.localeCompare(right.start) || left.title.localeCompare(right.title));
  const currentWindowMeetings = currentWindowMeetingsByCalendar
    .flat()
    .sort((left, right) => left.start.localeCompare(right.start) || left.title.localeCompare(right.title));
  const syncResult = await syncMeetingsToNotion(config, meetings, notion, syncOptions);
  const reconcileResult = await reconcileMissingNotionMeetings(config, currentWindowMeetings, notion, window, syncOptions);

  if (!syncResult.dryRun) {
    await saveState(config.sync.statePath, {
      version: 1,
      calendars: nextCalendars,
    });
  }

  return {
    ...syncResult,
    updated: syncResult.updated + updated + reconcileResult.updated,
    archived: syncResult.archived + archived + reconcileResult.archived,
  };
}

function canListMeetingPagesForWindow(notion: NotionClient): boolean {
  return Boolean(notion.listMeetingPagePropertiesForWindow || notion.listMeetingPagesForWindow);
}

async function reconcileMissingNotionMeetings(
  config: NolendarConfig,
  meetings: Meeting[],
  notion: NotionClient,
  window: CalendarWindow,
  syncOptions: SyncOptions,
): Promise<{ archived: number; updated: number }> {
  if (!notion.listMeetingPagePropertiesForWindow && !notion.listMeetingPagesForWindow) {
    return { archived: 0, updated: 0 };
  }

  const currentEventIds = new Set(meetings.map((meeting) => meeting.id));
  const queryArgs = {
    dataSourceId: config.notion.databaseId,
    datePropertyName: config.mapping.due,
    start: window.start,
    end: window.end,
  };
  const pages = notion.listMeetingPagePropertiesForWindow
    ? await notion.listMeetingPagePropertiesForWindow(queryArgs)
    : await notion.listMeetingPagesForWindow?.(queryArgs);
  let archived = 0;
  let updated = 0;

  for (const page of pages ?? []) {
    const eventId = readStringProperty(page.properties[config.mapping.eventId]);

    if (!eventId || currentEventIds.has(eventId)) {
      continue;
    }

    if (!isPageDateWithinWindow(page.properties[config.mapping.due], window)) {
      syncOptions.onDecision?.(
        `sync decision: notionEventId=${eventId} pageId=${page.id} decision=skip_missing_outside_window`,
      );
      continue;
    }

    const action = getCanceledMeetingAction(config);

    if (action.action === "set_status") {
      const currentStatus = readStringProperty(page.properties[action.statusProperty]);

      if (currentStatus === action.statusValue) {
        syncOptions.onDecision?.(
          `sync decision: notionEventId=${eventId} pageId=${page.id} decision=skip_missing_already_status property=${action.statusProperty} value=${action.statusValue}`,
        );
        continue;
      }

      updated += 1;
      syncOptions.onDecision?.(
        `sync decision: notionEventId=${eventId} pageId=${page.id} decision=set_status_missing_from_outlook property=${action.statusProperty} value=${action.statusValue}`,
      );
    } else {
      archived += 1;
      syncOptions.onDecision?.(`sync decision: notionEventId=${eventId} pageId=${page.id} decision=archive_missing_from_outlook`);
    }

    if (!syncOptions.dryRun) {
      await applyCanceledMeetingAction(config, notion, page.id);
    }
  }

  return { archived, updated };
}

function readStringProperty(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() !== "" ? value : undefined;
}

function isPageDateWithinWindow(value: unknown, window: CalendarWindow): boolean {
  const date = readDatePropertyStart(value);

  if (!date) {
    return false;
  }

  return date >= window.start && date < window.end;
}

function readDatePropertyStart(value: unknown): string | undefined {
  if (typeof value === "string") {
    return normalizeDateString(value);
  }

  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return undefined;
  }

  const start = (value as { start?: unknown }).start;

  return typeof start === "string" ? normalizeDateString(start) : undefined;
}

function normalizeDateString(value: string): string | undefined {
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const [year, month, day] = value.split("-").map((part) => Number.parseInt(part, 10));
    return new Date(year, month - 1, day).toISOString();
  }

  const parsed = new Date(value);

  return Number.isNaN(parsed.getTime()) ? undefined : parsed.toISOString();
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
