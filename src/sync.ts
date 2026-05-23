import type { NolendarConfig } from "./domain/config.js";
import type { Meeting } from "./domain/meeting.js";
import type { NotionClient } from "./notion/client.js";
import { shouldSyncMeeting } from "./filters.js";
import { validateOrEnsureNotionSchema } from "./notion/validation.js";

export interface SyncOptions {
  dryRun?: boolean;
  ensureProperties?: boolean;
  forceUpdate?: boolean;
}

export interface SyncResult {
  created: number;
  updated: number;
  skipped: number;
  filtered: number;
  archived: number;
  dryRun: boolean;
}

export async function syncMeetingsToNotion(
  config: NolendarConfig,
  meetings: Meeting[],
  notion: NotionClient,
  options: SyncOptions = {},
): Promise<SyncResult> {
  await validateOrEnsureNotionSchema(config, notion, {
    ensureProperties: options.ensureProperties,
  });

  const dataSource = await notion.retrieveDataSource(config.notion.databaseId);
  const result: SyncResult = {
    created: 0,
    updated: 0,
    skipped: 0,
    filtered: 0,
    archived: 0,
    dryRun: options.dryRun ?? false,
  };

  for (const meeting of meetings) {
    if (!shouldSyncMeeting(meeting, config.filters)) {
      result.filtered += 1;
      continue;
    }

    const existing = await notion.findPageByEventId({
      dataSourceId: dataSource.id,
      eventIdPropertyName: config.mapping.eventId,
      changeKeyPropertyName: config.mapping.changeKey,
      eventId: meeting.id,
    });

    if (!existing) {
      result.created += 1;
      if (!options.dryRun) {
        await notion.createMeetingPage({
          config,
          dataSource,
          meeting,
        });
      }
      continue;
    }

    if (existing.changeKey === meeting.changeKey && !options.forceUpdate) {
      result.skipped += 1;
      continue;
    }

    result.updated += 1;
    if (!options.dryRun) {
      await notion.updateMeetingPage({
        pageId: existing.id,
        config,
        dataSource,
        meeting,
      });
    }
  }

  return result;
}
