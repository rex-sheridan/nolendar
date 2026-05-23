import { Command } from "commander";

import { loadConfig } from "./config.js";
import type { NolendarConfig } from "./domain/config.js";
import { resolveGraphAuthConfig } from "./graph/auth.js";
import { AuthorizationCodeTokenProvider } from "./graph/authorization-code-token-provider.js";
import { DeviceCodeTokenProvider } from "./graph/device-code-token-provider.js";
import { GraphMeetingSource } from "./graph/graph-meeting-source.js";
import type { LookaheadWindow } from "./domain/config.js";
import { listMeetings } from "./list.js";
import { formatMeeting } from "./meeting-format.js";
import { writeDefaultConfig } from "./init-config.js";
import { resolveNotionAuthToken } from "./notion/auth.js";
import { ApiNotionClient } from "./notion/api-notion-client.js";
import { validateNotionSchema } from "./notion/schema.js";
import { validateOrEnsureNotionSchema } from "./notion/validation.js";
import { syncMeetingsToNotion } from "./sync.js";

export interface CliDependencies {
  stdout: Pick<Console, "log">;
  stderr: Pick<Console, "error">;
}

export function createCli(deps: CliDependencies = defaultDeps()): Command {
  const program = new Command();

  program.name("nolendar").description("Sync Outlook meetings into Notion.");

  program
    .command("list")
    .description("Print upcoming meetings for the requested time period.")
    .option("-c, --config <path>", "Path to YAML config file", "nolendar.yml")
    .option("--lookahead <window>", "One of: today, 24h, 7d")
    .action(async (options: { config: string; lookahead?: string }) => {
      const config = await loadConfig(options.config);
      const lookahead = resolveLookahead(config.sync.lookahead, options.lookahead);
      const meetingSource = buildGraphMeetingSource(config, deps);
      const result = await listMeetings(config, lookahead, { meetingSource });

      deps.stdout.log(
        `Meetings for ${result.lookahead} from ${config.calendars.length} configured calendar(s) between ${result.window.start} and ${result.window.end}:`,
      );

      if (result.meetings.length === 0) {
        deps.stdout.log("No meetings found.");
        return;
      }

      for (const meeting of result.meetings) {
        for (const line of formatMeeting(meeting)) {
          deps.stdout.log(line);
        }
      }
    });

  program
    .command("validate-config")
    .description("Validate YAML config and print the normalized result.")
    .option("-c, --config <path>", "Path to YAML config file", "nolendar.yml")
    .action(async (options: { config: string }) => {
      const config = await loadConfig(options.config);
      deps.stdout.log(JSON.stringify(config, null, 2));
    });

  program
    .command("validate-notion")
    .description("Validate Notion access and required data source properties.")
    .option("-c, --config <path>", "Path to YAML config file", "nolendar.yml")
    .option("--ensure-properties", "Create missing required properties if possible", false)
    .action(async (options: { config: string; ensureProperties: boolean }) => {
      const config = await loadConfig(options.config);
      const notion = buildNotionClient(deps);

      await validateOrEnsureNotionSchema(config, notion, {
        ensureProperties: options.ensureProperties,
      });

      const schema = await notion.retrieveDataSource(config.notion.databaseId);
      const result = validateNotionSchema(config, schema);

      deps.stdout.log(`Notion data source ${config.notion.databaseId} is valid.`);
      if (result.missing.length === 0 && result.mismatched.length === 0) {
        deps.stdout.log("Required properties are present.");
      }
    });

  program
    .command("sync")
    .description("Create or update Notion meeting pages from Outlook meetings.")
    .option("-c, --config <path>", "Path to YAML config file", "nolendar.yml")
    .option("--lookahead <window>", "One of: today, 24h, 7d")
    .option("--dry-run", "Preview sync actions without changing Notion", false)
    .option("--ensure-properties", "Create missing required properties if possible", false)
    .action(async (options: { config: string; lookahead?: string; dryRun: boolean; ensureProperties: boolean }) => {
      const config = await loadConfig(options.config);
      const lookahead = resolveLookahead(config.sync.lookahead, options.lookahead);
      const meetingSource = buildGraphMeetingSource(config, deps);
      const notion = buildNotionClient(deps);
      const meetingsResult = await listMeetings(config, lookahead, { meetingSource });
      const syncResult = await syncMeetingsToNotion(config, meetingsResult.meetings, notion, {
        dryRun: options.dryRun,
        ensureProperties: options.ensureProperties,
      });

      deps.stdout.log(
        `Sync summary: created=${syncResult.created}, updated=${syncResult.updated}, skipped=${syncResult.skipped}, filtered=${syncResult.filtered}, dryRun=${syncResult.dryRun}`,
      );
    });

  program
    .command("init")
    .description("Generate a starter nolendar.yml config file.")
    .option("-c, --config <path>", "Path to YAML config file", "nolendar.yml")
    .option("--force", "Overwrite an existing config file", false)
    .action(async (options: { config: string; force: boolean }) => {
      const writtenPath = await writeDefaultConfig(options.config, {
        force: options.force,
      });

      deps.stdout.log(`Wrote starter config to ${writtenPath}`);
    });

  program.exitOverride();

  return program;
}

export async function runCli(argv = process.argv, deps?: CliDependencies): Promise<number> {
  const cli = createCli(deps);

  try {
    await cli.parseAsync(argv);
    return 0;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const stderr = deps?.stderr ?? console;
    stderr.error(message);
    return 1;
  }
}

function resolveLookahead(configLookahead: LookaheadWindow, cliLookahead?: string): LookaheadWindow {
  if (cliLookahead === undefined) {
    return configLookahead;
  }

  if (cliLookahead === "today" || cliLookahead === "24h" || cliLookahead === "7d") {
    return cliLookahead;
  }

  throw new Error("`--lookahead` must be one of: today, 24h, 7d.");
}

function defaultDeps(): CliDependencies {
  return {
    stdout: console,
    stderr: console,
  };
}

function buildGraphMeetingSource(config: NolendarConfig, deps: CliDependencies): GraphMeetingSource {
  const authConfig = resolveGraphAuthConfig(config);

  if (authConfig.mode === "auth_code") {
    return new GraphMeetingSource(new AuthorizationCodeTokenProvider(authConfig, deps.stdout));
  }

  return new GraphMeetingSource(new DeviceCodeTokenProvider(authConfig, deps.stdout));
}

function buildNotionClient(_: CliDependencies): ApiNotionClient {
  return new ApiNotionClient(resolveNotionAuthToken());
}
