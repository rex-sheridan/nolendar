import type { NolendarConfig } from "./domain/config.js";
import type { Meeting } from "./domain/meeting.js";
import { shouldSyncMeeting } from "./filters.js";
import type { NotionClient } from "./notion/client.js";
import { validateOrEnsureNotionSchema } from "./notion/validation.js";

export interface FinalizeTemplatePagesOptions {
  ensureProperties?: boolean;
}

export interface FinalizeTemplatePagesResult {
  finalized: number;
  markedExisting: number;
  missingPage: number;
  skippedCancelled: number;
  filtered: number;
}

export async function finalizeTemplatedMeetingPages(
  config: NolendarConfig,
  meetings: Meeting[],
  notion: NotionClient,
  options: FinalizeTemplatePagesOptions = {},
): Promise<FinalizeTemplatePagesResult> {
  if (!config.notion.dataSourceTemplate) {
    throw new Error("`finalize-templates` requires `notion.dataSourceTemplate` to be configured.");
  }

  await validateOrEnsureNotionSchema(config, notion, {
    ensureProperties: options.ensureProperties,
  });

  const dataSource = await notion.retrieveDataSource(config.notion.databaseId);
  const result: FinalizeTemplatePagesResult = {
    finalized: 0,
    markedExisting: 0,
    missingPage: 0,
    skippedCancelled: 0,
    filtered: 0,
  };

  for (const meeting of meetings) {
    if (!shouldSyncMeeting(meeting, config.filters)) {
      result.filtered += 1;
      continue;
    }

    if (meeting.isCancelled) {
      result.skippedCancelled += 1;
      continue;
    }

    const existing = await notion.findPageByEventId({
      dataSourceId: dataSource.id,
      eventIdPropertyName: config.mapping.eventId,
      changeKeyPropertyName: config.mapping.changeKey,
      eventId: meeting.id,
    });

    if (!existing) {
      result.missingPage += 1;
      continue;
    }

    const finalizeResult = await notion.finalizeMeetingPageContent({
      pageId: existing.id,
      config,
      dataSource,
      meeting,
    });

    if (finalizeResult === "marked_existing") {
      result.markedExisting += 1;
      continue;
    }

    result.finalized += 1;
  }

  return result;
}
