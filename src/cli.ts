import { Command } from "commander";

import { createConsoleTimingReporter, type ApiTimingReporter } from "./api-timing.js";
import { loadConfig } from "./config.js";
import { ReadlineConfigWizardPrompt, runConfigWizard, type ConfigWizardPrompt } from "./config-wizard.js";
import type { MicrosoftAuthMode, MicrosoftConfig, NolendarConfig } from "./domain/config.js";
import { resolveGraphAuthConfig } from "./graph/auth.js";
import { AuthorizationCodeTokenProvider } from "./graph/authorization-code-token-provider.js";
import { DeviceCodeTokenProvider } from "./graph/device-code-token-provider.js";
import { GraphMeetingSource, type AvailableCalendar } from "./graph/graph-meeting-source.js";
import { InteractiveBrowserTokenProvider } from "./graph/interactive-browser-token-provider.js";
import { StaticAccessTokenProvider } from "./graph/static-access-token-provider.js";
import type { LookaheadWindow } from "./domain/config.js";
import { listMeetings } from "./list.js";
import { isValidLookaheadWindow } from "./lookahead.js";
import { formatMeeting } from "./meeting-format.js";
import { writeDefaultConfig } from "./init-config.js";
import { finalizeTemplatedMeetingPages } from "./finalize-templates.js";
import { resolveNotionAuthToken } from "./notion/auth.js";
import { ApiNotionClient } from "./notion/api-notion-client.js";
import type { NotionClient } from "./notion/client.js";
import { validateNotionSchema } from "./notion/schema.js";
import { validateOrEnsureNotionSchema } from "./notion/validation.js";
import { syncCalendarChangesToNotion } from "./delta-sync.js";
import { listMeetingContentsForDay, type MeetingContentsDetail, type MeetingContentsSource } from "./meeting-contents.js";

export interface CliDependencies {
  stdout: Pick<Console, "log">;
  stderr: Pick<Console, "error">;
  loadConfig?: typeof loadConfig;
  timingSink?: Pick<Console, "log">;
  buildNotionClient?: (options?: { timingReporter?: ApiTimingReporter }) => NotionClient;
  configWizardPrompt?: ConfigWizardPrompt;
  stdin?: NodeJS.ReadableStream;
  output?: NodeJS.WritableStream;
}

export function createCli(deps: CliDependencies = defaultDeps()): Command {
  const program = new Command();
  const loadConfigFn = deps.loadConfig ?? loadConfig;
  const buildNotionClientFn = deps.buildNotionClient ?? ((options?: { timingReporter?: ApiTimingReporter }) => buildNotionClient(deps, options));

  program.name("nolendar").description("Sync Outlook meetings into Notion.");

  program
    .command("list-calendars")
    .description("Print available Outlook calendars and their IDs.")
    .option("-c, --config <path>", "Path to YAML config file for Microsoft auth settings")
    .option("--tenant <tenant>", "Microsoft tenant: common, organizations, or consumers")
    .option("--auth-mode <mode>", "Microsoft auth mode: device_code, interactive_browser, or auth_code")
    .option("--timings", "Print API call timings", false)
    .action(async (options: { config?: string; tenant?: string; authMode?: string; timings: boolean }) => {
      const microsoft = await resolveMicrosoftConfig(options, loadConfigFn);
      const timingReporter = options.timings ? createTimingReporter(deps) : undefined;
      const graph = buildGraphMeetingSourceFromMicrosoftConfig(microsoft, deps, timingReporter);
      const calendars = await graph.listCalendars();

      if (calendars.length === 0) {
        deps.stdout.log("No calendars found.");
        return;
      }

      deps.stdout.log("Available calendars:");
      for (const line of formatCalendarOutput(calendars)) {
        deps.stdout.log(line);
      }
    });

  program
    .command("list")
    .description("Print upcoming meetings for the requested time period.")
    .option("-c, --config <path>", "Path to YAML config file", "nolendar.yml")
    .option("--lookahead <window>", "One of: today, 24h, 7d")
    .option("--timings", "Print API call timings", false)
    .action(async (options: { config: string; lookahead?: string; timings: boolean }) => {
      const config = await loadConfigFn(options.config);
      const lookahead = resolveLookahead(config.sync.lookahead, options.lookahead);
      const timingReporter = options.timings ? createTimingReporter(deps) : undefined;
      const meetingSource = buildGraphMeetingSource(config, deps, timingReporter);
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
    .command("meetings")
    .description("Print meeting contents for a single day from Outlook or Notion.")
    .option("-c, --config <path>", "Path to YAML config file", "nolendar.yml")
    .option("--source <source>", "Meeting source: outlook or notion", "outlook")
    .option("--day <day>", "Day to inspect: today, tomorrow, yesterday, +/-Nd, or YYYY-MM-DD", "today")
    .option("--full-properties", "Print all available normalized properties in addition to title/date/body", false)
    .option("--timings", "Print API call timings", false)
    .action(
      async (options: {
        config: string;
        source: string;
        day: string;
        fullProperties: boolean;
        timings: boolean;
      }) => {
        const config = await loadConfigFn(options.config);
        const source = resolveMeetingContentsSource(options.source);
        const detail: MeetingContentsDetail = options.fullProperties ? "full" : "compact";
        const timingReporter = options.timings ? createTimingReporter(deps) : undefined;
        const output = await listMeetingContentsForDay(
          config,
          {
            source,
            day: options.day,
            detail,
          },
          source === "outlook"
            ? {
                meetingSource: buildGraphMeetingSource(config, deps, timingReporter),
              }
            : {
                notion: buildNotionClientFn({
                  timingReporter,
                }),
              },
        );

        for (const line of output) {
          deps.stdout.log(line);
        }
      },
    );

  program
    .command("validate-config")
    .description("Validate YAML config and print the normalized result.")
    .option("-c, --config <path>", "Path to YAML config file", "nolendar.yml")
    .action(async (options: { config: string }) => {
      const config = await loadConfigFn(options.config);
      deps.stdout.log(JSON.stringify(config, null, 2));
    });

  program
    .command("validate-notion")
    .description("Validate Notion access and required data source properties.")
    .option("-c, --config <path>", "Path to YAML config file", "nolendar.yml")
    .option("--ensure-properties", "Create missing required properties if possible", false)
    .option("--timings", "Print API call timings", false)
    .action(async (options: { config: string; ensureProperties: boolean; timings: boolean }) => {
      const config = await loadConfigFn(options.config);
      const notion = buildNotionClientFn({
        timingReporter: options.timings ? createTimingReporter(deps) : undefined,
      });

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
    .command("print-notion-schema")
    .description("Print the detected Notion schema for the configured meeting and People data sources.")
    .option("-c, --config <path>", "Path to YAML config file", "nolendar.yml")
    .option("--timings", "Print API call timings", false)
    .action(async (options: { config: string; timings: boolean }) => {
      const config = await loadConfigFn(options.config);
      const notion = buildNotionClientFn({
        timingReporter: options.timings ? createTimingReporter(deps) : undefined,
      });

      const meetingSchema = await notion.retrieveDataSource(config.notion.databaseId);
      for (const line of formatSchemaOutput("Meeting data source", meetingSchema.id, meetingSchema.title, meetingSchema.properties)) {
        deps.stdout.log(line);
      }

      if (config.notion.peopleDataSource) {
        const peopleSchema = await notion.retrieveDataSource(config.notion.peopleDataSource.databaseId);
        for (const line of formatSchemaOutput("People data source", peopleSchema.id, peopleSchema.title, peopleSchema.properties)) {
          deps.stdout.log(line);
        }
      }
    });

  program
    .command("sync")
    .description("Create or update Notion meeting pages from Outlook meetings.")
    .option("-c, --config <path>", "Path to YAML config file", "nolendar.yml")
    .option("--lookahead <window>", "One of: today, 24h, 7d")
    .option("--dry-run", "Preview sync actions without changing Notion", false)
    .option("--ensure-properties", "Create missing required properties if possible", false)
    .option("--force-update", "Update matching Notion pages even when the Outlook changeKey is unchanged", false)
    .option("--finalize-delay-ms <ms>", "Delay in milliseconds before the native-template finalize pass", "3000")
    .option("--timings", "Print API call timings", false)
    .option("--verbose", "Print per-meeting sync decisions", false)
    .action(
      async (options: {
        config: string;
        lookahead?: string;
        dryRun: boolean;
        ensureProperties: boolean;
        forceUpdate: boolean;
        finalizeDelayMs: string;
        timings: boolean;
        verbose: boolean;
      }) => {
      const startedAt = Date.now();
      const config = await loadConfigFn(options.config);
      const lookahead = resolveLookahead(config.sync.lookahead, options.lookahead);
      const finalizeDelayMs = parseDelayMs(options.finalizeDelayMs);
      const timingReporter = options.timings ? createTimingReporter(deps) : undefined;
      const meetingSource = buildGraphMeetingSource(config, deps, timingReporter);
      const notion = buildNotionClientFn({
        timingReporter,
      });
      const syncResult = await syncCalendarChangesToNotion(
        {
          ...config,
          sync: {
            ...config.sync,
            lookahead,
          },
        },
        notion,
        {
          meetingSource,
          syncOptions: {
            dryRun: options.dryRun,
            ensureProperties: options.ensureProperties,
            forceUpdate: options.forceUpdate,
            onDecision: options.verbose ? (message) => deps.stdout.log(message) : undefined,
          },
        },
      );

      if (config.notion.dataSourceTemplate && !options.dryRun) {
        await sleep(finalizeDelayMs);
        const meetingsResult = await listMeetings(config, lookahead, { meetingSource });
        await finalizeTemplatedMeetingPages(config, meetingsResult.meetings, notion, {
          ensureProperties: options.ensureProperties,
        });
      }

      deps.stdout.log(
        `Sync summary: created=${syncResult.created}, updated=${syncResult.updated}, archived=${syncResult.archived}, skipped=${syncResult.skipped}, filtered=${syncResult.filtered}, dryRun=${syncResult.dryRun}, elapsedMs=${Date.now() - startedAt}`,
      );
      },
    );

  program
    .command("finalize-templates")
    .description("Append Nolendar meeting content to pages created from native Notion data source templates.")
    .option("-c, --config <path>", "Path to YAML config file", "nolendar.yml")
    .option("--lookahead <window>", "One of: today, 24h, 7d")
    .option("--ensure-properties", "Create missing required properties if possible", false)
    .option("--timings", "Print API call timings", false)
    .action(
      async (options: {
        config: string;
        lookahead?: string;
        ensureProperties: boolean;
        timings: boolean;
      }) => {
        const config = await loadConfigFn(options.config);
        const lookahead = resolveLookahead(config.sync.lookahead, options.lookahead);
        const timingReporter = options.timings ? createTimingReporter(deps) : undefined;
        const meetingSource = buildGraphMeetingSource(config, deps, timingReporter);
        const notion = buildNotionClientFn({
          timingReporter,
        });
        const meetingsResult = await listMeetings(config, lookahead, { meetingSource });
        await finalizeTemplatedMeetingPages(config, meetingsResult.meetings, notion, {
          ensureProperties: options.ensureProperties,
        });
      },
    );

  program
    .command("init")
    .description("Generate a starter nolendar.yml config file.")
    .option("-c, --config <path>", "Path to YAML config file", "nolendar.yml")
    .option("--force", "Overwrite an existing config file", false)
    .option("--wizard", "Guide setup interactively", false)
    .action(async (options: { config: string; force: boolean; wizard: boolean }) => {
      if (options.wizard) {
        const writtenPath = await runConfigWizard({
          configPath: options.config,
          force: options.force,
          prompt: deps.configWizardPrompt ?? new ReadlineConfigWizardPrompt(deps.stdin ?? process.stdin, deps.output ?? process.stdout, deps.stdout),
          stdout: deps.stdout,
          notion: buildNotionClientFn(),
          listCalendars: async (microsoft) =>
            buildGraphMeetingSourceFromMicrosoftConfig(microsoft, deps).listCalendars(),
        });

        deps.stdout.log(`Wrote config to ${writtenPath}`);
        return;
      }

      const writtenPath = await writeDefaultConfig(options.config, {
        force: options.force,
      });

      deps.stdout.log(`Wrote starter config to ${writtenPath}`);
    });

  program
    .command("wizard")
    .description("Guide setup of a nolendar.yml config file.")
    .option("-c, --config <path>", "Path to YAML config file", "nolendar.yml")
    .option("--force", "Overwrite an existing config file", false)
    .action(async (options: { config: string; force: boolean }) => {
      const writtenPath = await runConfigWizard({
        configPath: options.config,
        force: options.force,
        prompt: deps.configWizardPrompt ?? new ReadlineConfigWizardPrompt(deps.stdin ?? process.stdin, deps.output ?? process.stdout, deps.stdout),
        stdout: deps.stdout,
        notion: buildNotionClientFn(),
        listCalendars: async (microsoft) =>
          buildGraphMeetingSourceFromMicrosoftConfig(microsoft, deps).listCalendars(),
      });

      deps.stdout.log(`Wrote config to ${writtenPath}`);
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

  if (isValidLookaheadWindow(cliLookahead)) {
    return cliLookahead;
  }

  throw new Error("`--lookahead` must be `today` or a relative range like `12h`, `5d`, `2w`, or `3m`.");
}

function resolveMeetingContentsSource(value: string): MeetingContentsSource {
  if (value === "outlook" || value === "notion") {
    return value;
  }

  throw new Error("`--source` must be one of: outlook, notion.");
}

function defaultDeps(): CliDependencies {
  return {
    stdout: console,
    stderr: console,
    timingSink: console,
  };
}

function formatSchemaOutput(
  label: string,
  dataSourceId: string,
  title: string | undefined,
  properties: Record<string, { type: string }>,
): string[] {
  const lines = [`${label}: ${title ? `${title} ` : ""}(${dataSourceId})`];
  const entries = Object.entries(properties).sort(([left], [right]) => left.localeCompare(right));

  if (entries.length === 0) {
    lines.push("  No properties found.");
    return lines;
  }

  for (const [name, property] of entries) {
    lines.push(`  - ${name}: ${property.type}`);
  }

  return lines;
}

function createTimingReporter(deps: CliDependencies): ApiTimingReporter {
  return createConsoleTimingReporter(deps.timingSink ?? deps.stdout);
}

function parseDelayMs(value: string): number {
  const parsed = Number.parseInt(value, 10);

  if (!Number.isFinite(parsed) || parsed < 0) {
    throw new Error("`--finalize-delay-ms` must be a non-negative integer.");
  }

  return parsed;
}

async function resolveMicrosoftConfig(
  options: { config?: string; tenant?: string; authMode?: string },
  loadConfigFn: typeof loadConfig,
): Promise<MicrosoftConfig> {
  const baseConfig = options.config ? (await loadConfigFn(options.config)).microsoft : defaultMicrosoftConfig();

  return {
    tenant: resolveMicrosoftTenant(options.tenant, baseConfig.tenant),
    authMode: resolveMicrosoftAuthMode(options.authMode, baseConfig.authMode),
  };
}

function defaultMicrosoftConfig(): MicrosoftConfig {
  return {
    tenant: "common",
    authMode: "device_code",
  };
}

function resolveMicrosoftTenant(value: string | undefined, fallback: MicrosoftConfig["tenant"]): MicrosoftConfig["tenant"] {
  if (value === undefined) {
    return fallback;
  }

  if (value === "common" || value === "organizations" || value === "consumers") {
    return value;
  }

  throw new Error("`--tenant` must be one of: common, organizations, consumers.");
}

function resolveMicrosoftAuthMode(value: string | undefined, fallback: MicrosoftAuthMode): MicrosoftAuthMode {
  if (value === undefined) {
    return fallback;
  }

  if (value === "device_code" || value === "interactive_browser" || value === "auth_code") {
    return value;
  }

  throw new Error("`--auth-mode` must be one of: device_code, interactive_browser, auth_code.");
}

function formatCalendarOutput(calendars: AvailableCalendar[]): string[] {
  return calendars.flatMap((calendar) => {
    const tags = calendar.isDefaultCalendar ? " [default]" : "";
    const owner = formatCalendarOwner(calendar);
    const lines = [`  - ${calendar.name}${tags}`, `    id: ${calendar.id}`];

    if (owner) {
      lines.push(`    owner: ${owner}`);
    }

    return lines;
  });
}

function formatCalendarOwner(calendar: AvailableCalendar): string | undefined {
  if (calendar.ownerName && calendar.ownerAddress) {
    return `${calendar.ownerName} <${calendar.ownerAddress}>`;
  }

  return calendar.ownerName ?? calendar.ownerAddress;
}

async function sleep(ms: number): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, ms));
}

function buildGraphMeetingSource(
  config: NolendarConfig,
  deps: CliDependencies,
  timingReporter?: ApiTimingReporter,
): GraphMeetingSource {
  return buildGraphMeetingSourceFromMicrosoftConfig(config.microsoft, deps, timingReporter);
}

function buildGraphMeetingSourceFromMicrosoftConfig(
  microsoft: MicrosoftConfig,
  deps: CliDependencies,
  timingReporter?: ApiTimingReporter,
): GraphMeetingSource {
  const authConfig = resolveGraphAuthConfig({ microsoft });

  if (authConfig.mode === "auth_code") {
    return new GraphMeetingSource(new AuthorizationCodeTokenProvider(authConfig, deps.stdout), undefined, timingReporter);
  }

  if (authConfig.mode === "interactive_browser") {
    return new GraphMeetingSource(new InteractiveBrowserTokenProvider(authConfig), undefined, timingReporter);
  }

  if (authConfig.mode === "static_access_token") {
    return new GraphMeetingSource(new StaticAccessTokenProvider(authConfig.accessToken), undefined, timingReporter);
  }

  return new GraphMeetingSource(new DeviceCodeTokenProvider(authConfig, deps.stdout), undefined, timingReporter);
}

function buildNotionClient(_: CliDependencies, options?: { timingReporter?: ApiTimingReporter }): ApiNotionClient {
  return new ApiNotionClient(resolveNotionAuthToken(), undefined, options?.timingReporter);
}
