import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";

const DEFAULT_CONFIG_TEMPLATE = `microsoft:
  tenant: common
  authMode: device_code

notion:
  databaseId: your_notion_data_source_id

calendars:
  - id: primary
    name: Primary

filters:
  ignoreDeclined: true
  minDurationMinutes: 15
  requireAttendees: false
  ignorePersonal: false
  ignoreOptionalAttendance: false

mapping:
  title: Name
  due: Due
  eventId: Outlook Event ID
  changeKey: Outlook ChangeKey

sync:
  lookahead: today
  statePath: ./.nolendar/state.json
`;

export class InitConfigError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "InitConfigError";
  }
}

export interface InitConfigOptions {
  force?: boolean;
}

export async function writeDefaultConfig(configPath: string, options: InitConfigOptions = {}): Promise<string> {
  const absolutePath = path.resolve(configPath);
  await mkdir(path.dirname(absolutePath), { recursive: true });

  try {
    await writeFile(absolutePath, DEFAULT_CONFIG_TEMPLATE, {
      encoding: "utf8",
      flag: options.force ? "w" : "wx",
    });
  } catch (error) {
    if (isAlreadyExistsError(error)) {
      throw new InitConfigError(`Config file already exists at ${absolutePath}. Use --force to overwrite it.`);
    }

    throw error;
  }

  return absolutePath;
}

function isAlreadyExistsError(error: unknown): error is NodeJS.ErrnoException {
  return typeof error === "object" && error !== null && "code" in error && error.code === "EEXIST";
}
