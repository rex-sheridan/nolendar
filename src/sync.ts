import type { NolendarConfig } from "./domain/config.js";
import type { Meeting } from "./domain/meeting.js";
import type { NotionClient } from "./notion/client.js";
import { getMeetingFilterReason } from "./filters.js";
import { validateOrEnsureNotionSchema } from "./notion/validation.js";

export interface SyncOptions {
  dryRun?: boolean;
  ensureProperties?: boolean;
  forceUpdate?: boolean;
  onDecision?: SyncDecisionReporter;
}

export type SyncDecisionReporter = (message: string) => void;

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
    const existing = await notion.findPageByEventId({
      dataSourceId: dataSource.id,
      eventIdPropertyName: config.mapping.eventId,
      changeKeyPropertyName: config.mapping.changeKey,
      eventId: meeting.id,
    });

    options.onDecision?.(
      formatMeetingDecisionPrefix(meeting, existing?.id) +
        ` changeKey=${meeting.changeKey} notionChangeKey=${existing?.changeKey ?? "-"} cancelled=${meeting.isCancelled}`,
    );

    if (meeting.isCancelled) {
      if (!existing) {
        result.skipped += 1;
        options.onDecision?.(`${formatMeetingDecisionPrefix(meeting)} decision=skip_cancelled_missing_page`);
        continue;
      }

      const action = getCanceledMeetingAction(config);

      if (action.action === "archive") {
        result.archived += 1;
        options.onDecision?.(`${formatMeetingDecisionPrefix(meeting, existing.id)} decision=archive_cancelled`);
      } else {
        result.updated += 1;
        options.onDecision?.(
          `${formatMeetingDecisionPrefix(meeting, existing.id)} decision=set_status_cancelled property=${action.statusProperty} value=${action.statusValue}`,
        );
      }

      if (!options.dryRun) {
        await applyCanceledMeetingAction(config, notion, existing.id);
      }
      continue;
    }

    const filterReason = getMeetingFilterReason(meeting, config.filters);

    if (filterReason) {
      result.filtered += 1;
      options.onDecision?.(`${formatMeetingDecisionPrefix(meeting, existing?.id)} decision=filtered reason=${filterReason}`);
      continue;
    }

    if (!existing) {
      result.created += 1;
      options.onDecision?.(`${formatMeetingDecisionPrefix(meeting)} decision=create`);
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
      options.onDecision?.(`${formatMeetingDecisionPrefix(meeting, existing.id)} decision=skip_change_key_match`);
      continue;
    }

    result.updated += 1;
    options.onDecision?.(`${formatMeetingDecisionPrefix(meeting, existing.id)} decision=update`);
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

export async function applyCanceledMeetingAction(
  config: NolendarConfig,
  notion: NotionClient,
  pageId: string,
): Promise<"archived" | "updated"> {
  const action = getCanceledMeetingAction(config);

  if (action.action === "archive") {
    await notion.archivePage(pageId);
    return "archived";
  }

  if (!notion.setPageStatus) {
    throw new Error("The configured Notion client does not support setting page status properties.");
  }

  await notion.setPageStatus({
    pageId,
    propertyName: action.statusProperty,
    statusName: action.statusValue,
  });
  return "updated";
}

export function getCanceledMeetingAction(
  config: NolendarConfig,
): NonNullable<NolendarConfig["notion"]["canceledMeetings"]> {
  return config.notion.canceledMeetings ?? { action: "archive" };
}

function formatMeetingDecisionPrefix(meeting: Meeting, pageId?: string): string {
  return `sync decision: title=${JSON.stringify(meeting.title)} eventId=${meeting.id} pageId=${pageId ?? "-"}`;
}
