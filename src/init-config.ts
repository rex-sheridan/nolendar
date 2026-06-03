import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";

const DEFAULT_CONFIG_TEMPLATE = `microsoft:
  tenant: common
  authMode: device_code

notion:
  databaseId: your_notion_data_source_id
  # Choose at most one template mode:
  # templatePageId: your_template_page_id
  # dataSourceTemplate:
  #   type: default
  #   # templateId: your_data_source_template_id
  #   # timezone: America/New_York
  # Native data source templates use a deferred content workflow:
  # sync will finalize generated meeting content automatically after a short delay
  # finalize-templates is still available as a manual fallback
  peopleDataSource:
    databaseId: your_people_data_source_id
    nameProperty: Name
    emailProperty: Email Address
    maxAttendeesPerMeeting: 10
  canceledMeetings:
    action: set_status
    statusProperty: Status
    statusValue: Canceled
  completedMeetings:
    statusProperty: Status
    doneStatusValue: Done
    canceledStatusValue: Canceled
    lookback: 1d
  defaultTags:
    - meeting
  defaultAssigneeEmail: you@example.com
  pageIcon:
    type: emoji
    emoji: "📝"
  pageContent:
    sections:
      - meeting_link
      - calendar_event
      - meeting_details
      - notes
      - action_items

calendars:
  - id: primary
    name: Primary

filters:
  ignoreDeclined: true
  minDurationMinutes: 15
  requireAttendees: false
  ignorePersonal: false
  ignoreOptionalAttendance: false
  ignoreNames:
    - Focus time
  ignorePatterns:
    - "^OOO"

mapping:
  title: Name
  due: Due
  eventId: Outlook Event ID
  changeKey: Outlook ChangeKey
  eventLink: Source URL
  tags: Tags
  assignee: Assignee
  participants: Participants

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
