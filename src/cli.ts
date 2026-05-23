import { Command } from "commander";

import { loadConfig } from "./config.js";
import { resolveGraphAuthConfig } from "./graph/auth.js";
import { DeviceCodeTokenProvider } from "./graph/device-code-token-provider.js";
import { GraphMeetingSource } from "./graph/graph-meeting-source.js";
import type { LookaheadWindow } from "./domain/config.js";
import { listMeetings } from "./list.js";
import { formatMeeting } from "./meeting-format.js";

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
      const authConfig = resolveGraphAuthConfig(config);
      const meetingSource = new GraphMeetingSource(new DeviceCodeTokenProvider(authConfig, deps.stdout));
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
