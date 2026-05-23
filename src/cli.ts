import { Command } from "commander";

import { loadConfig } from "./config.js";
import type { LookaheadWindow } from "./domain/config.js";

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

      deps.stdout.log(`Meetings for ${lookahead} from ${config.calendars.length} configured calendar(s):`);
      for (const calendar of config.calendars) {
        deps.stdout.log(`- ${calendar.name ?? calendar.id}`);
      }
      deps.stdout.log("Graph meeting retrieval is not implemented yet.");
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
