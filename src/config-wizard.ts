import { access, mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { createInterface } from "node:readline/promises";
import { stringify } from "yaml";

import type {
  CalendarConfig,
  FiltersConfig,
  MappingConfig,
  MicrosoftAuthMode,
  MicrosoftConfig,
  NolendarConfig,
  NotionPageContentSection,
} from "./domain/config.js";
import type {
  NotionDataSourceSchema,
  NotionDataSourceSummary,
  NotionDataSourceTemplateSummary,
  RequiredNotionProperty,
} from "./domain/notion.js";
import type { AvailableCalendar } from "./graph/graph-meeting-source.js";
import type { NotionClient } from "./notion/client.js";
import { buildRequiredNotionProperties, buildRequiredPeopleDataSourceProperties } from "./notion/schema.js";

export interface ConfigWizardPrompt {
  ask(question: string, defaultValue?: string): Promise<string>;
  confirm(question: string, defaultValue?: boolean): Promise<boolean>;
  select<T extends string>(question: string, choices: Array<{ label: string; value: T }>, defaultValue?: T): Promise<T>;
  multiselect<T extends string>(
    question: string,
    choices: Array<{ label: string; value: T }>,
    defaultValues?: T[],
  ): Promise<T[]>;
  close?(): void;
}

export interface ConfigWizardOptions {
  configPath: string;
  force?: boolean;
  prompt: ConfigWizardPrompt;
  stdout: Pick<Console, "log">;
  notion: NotionClient;
  listCalendars(microsoft: MicrosoftConfig): Promise<AvailableCalendar[]>;
}

type MeetingMappingKey = keyof MappingConfig;

const MEETING_MAPPING_STEPS: Array<{
  key: MeetingMappingKey;
  label: string;
  source: string;
  type: RequiredNotionProperty["type"];
  defaultName: string;
  required: boolean;
}> = [
  { key: "title", label: "Meeting title", source: "Outlook subject", type: "title", defaultName: "Name", required: true },
  { key: "due", label: "Meeting date", source: "Outlook start/end", type: "date", defaultName: "Due", required: true },
  {
    key: "eventId",
    label: "Outlook event ID",
    source: "Outlook event id",
    type: "rich_text",
    defaultName: "Outlook Event ID",
    required: true,
  },
  {
    key: "changeKey",
    label: "Outlook change key",
    source: "Outlook changeKey",
    type: "rich_text",
    defaultName: "Outlook ChangeKey",
    required: true,
  },
  { key: "eventLink", label: "Source URL", source: "Outlook web link", type: "url", defaultName: "Source URL", required: false },
  { key: "tags", label: "Default tags", source: "notion.defaultTags", type: "multi_select", defaultName: "Tags", required: false },
  { key: "assignee", label: "Default assignee", source: "Notion user", type: "people", defaultName: "Assignee", required: false },
  {
    key: "participants",
    label: "Participants",
    source: "Outlook attendees",
    type: "relation",
    defaultName: "Participants",
    required: false,
  },
];

const PAGE_CONTENT_SECTIONS: NotionPageContentSection[] = [
  "meeting_link",
  "calendar_event",
  "meeting_details",
  "notes",
  "action_items",
];

export class ReadlineConfigWizardPrompt implements ConfigWizardPrompt {
  private readonly rl: ReturnType<typeof createInterface>;

  constructor(input: NodeJS.ReadableStream, output: NodeJS.WritableStream, private readonly stdout: Pick<Console, "log">) {
    this.rl = createInterface({ input, output });
  }

  async ask(question: string, defaultValue?: string): Promise<string> {
    const suffix = defaultValue === undefined ? "" : ` [${defaultValue}]`;
    const answer = (await this.rl.question(`${question}${suffix}: `)).trim();

    return answer || defaultValue || "";
  }

  async confirm(question: string, defaultValue = false): Promise<boolean> {
    const suffix = defaultValue ? " [Y/n]" : " [y/N]";
    const answer = (await this.rl.question(`${question}${suffix}: `)).trim().toLowerCase();

    if (!answer) {
      return defaultValue;
    }

    return answer === "y" || answer === "yes";
  }

  async select<T extends string>(question: string, choices: Array<{ label: string; value: T }>, defaultValue?: T): Promise<T> {
    this.stdout.log(question);
    choices.forEach((choice, index) => {
      const marker = choice.value === defaultValue ? " (default)" : "";
      this.stdout.log(`  ${index + 1}. ${choice.label}${marker}`);
    });

    while (true) {
      const answer = await this.ask("Choose one", defaultValue ? String(choices.findIndex((choice) => choice.value === defaultValue) + 1) : undefined);
      const index = Number.parseInt(answer, 10) - 1;

      if (choices[index]) {
        return choices[index].value;
      }

      const directMatch = choices.find((choice) => choice.value === answer);
      if (directMatch) {
        return directMatch.value;
      }

      this.stdout.log("Please enter a listed number.");
    }
  }

  async multiselect<T extends string>(
    question: string,
    choices: Array<{ label: string; value: T }>,
    defaultValues: T[] = [],
  ): Promise<T[]> {
    this.stdout.log(question);
    choices.forEach((choice, index) => {
      const marker = defaultValues.includes(choice.value) ? " (default)" : "";
      this.stdout.log(`  ${index + 1}. ${choice.label}${marker}`);
    });

    while (true) {
      const defaultAnswer = defaultValues
        .map((value) => choices.findIndex((choice) => choice.value === value) + 1)
        .filter((index) => index > 0)
        .join(",");
      const answer = await this.ask("Choose one or more, comma-separated", defaultAnswer || undefined);
      const indexes = answer
        .split(",")
        .map((entry) => Number.parseInt(entry.trim(), 10) - 1)
        .filter((index) => Number.isInteger(index));
      const selected = indexes.map((index) => choices[index]).filter(Boolean);

      if (selected.length > 0 && selected.length === indexes.length) {
        return selected.map((choice) => choice.value);
      }

      this.stdout.log("Please enter one or more listed numbers.");
    }
  }

  close(): void {
    this.rl.close();
  }
}

export async function runConfigWizard(options: ConfigWizardOptions): Promise<string> {
  const absolutePath = path.resolve(options.configPath);
  await prepareConfigPath(absolutePath, options.force);

  const config = defaultConfig();
  const save = async () => {
    await writeConfig(absolutePath, config);
  };

  try {
    options.stdout.log("Nolendar configuration wizard");
    options.stdout.log(`Writing progress to ${absolutePath}`);

    await configureMicrosoft(config, options.prompt, save);
    await configureCalendars(config, options, save);
    const meetingSchema = await configureNotionDataSources(config, options, save);
    await configureMappings(config, meetingSchema, options, save);
    await ensureConfiguredProperties(config, options);

    while (await options.prompt.confirm("Configure additional sections?", false)) {
      const section = await options.prompt.select(
        "Additional sections",
        [
          { label: "Notion templates, content, defaults, and canceled meetings", value: "notion" },
          { label: "Filters", value: "filters" },
          { label: "Sync window and state path", value: "sync" },
          { label: "Microsoft authentication", value: "microsoft" },
          { label: "Calendars", value: "calendars" },
          { label: "Mappings", value: "mappings" },
          { label: "Done", value: "done" },
        ],
        "done",
      );

      if (section === "done") {
        break;
      }

      if (section === "notion") {
        await configureNotionOptional(config, options, save);
      } else if (section === "filters") {
        await configureFilters(config, options.prompt, save);
      } else if (section === "sync") {
        await configureSync(config, options.prompt, save);
      } else if (section === "microsoft") {
        await configureMicrosoft(config, options.prompt, save);
      } else if (section === "calendars") {
        await configureCalendars(config, options, save);
      } else if (section === "mappings") {
        const schema = await options.notion.retrieveDataSource(config.notion.databaseId);
        await configureMappings(config, schema, options, save);
        await ensureConfiguredProperties(config, options);
      }
    }

    await save();
    return absolutePath;
  } finally {
    options.prompt.close?.();
  }
}

async function configureMicrosoft(config: NolendarConfig, prompt: ConfigWizardPrompt, save: () => Promise<void>): Promise<void> {
  config.microsoft.tenant = await prompt.select(
    "Microsoft tenant",
    [
      { label: "Any Microsoft account or tenant", value: "common" },
      { label: "Work or school accounts only", value: "organizations" },
      { label: "Personal Microsoft accounts only", value: "consumers" },
    ],
    config.microsoft.tenant,
  );
  await save();

  config.microsoft.authMode = await prompt.select<MicrosoftAuthMode>(
    "Microsoft auth mode",
    [
      { label: "Device code", value: "device_code" },
      { label: "Interactive browser", value: "interactive_browser" },
      { label: "Authorization code", value: "auth_code" },
    ],
    config.microsoft.authMode,
  );
  await save();
}

async function configureCalendars(
  config: NolendarConfig,
  options: ConfigWizardOptions,
  save: () => Promise<void>,
): Promise<void> {
  options.stdout.log("Loading Outlook calendars...");
  const calendars = await options.listCalendars(config.microsoft);

  if (calendars.length === 0) {
    throw new Error("No Outlook calendars were found for the selected Microsoft account.");
  }

  const defaultCalendarIds = calendars.filter((calendar) => calendar.isDefaultCalendar).map((calendar) => calendar.id);
  const selectedIds = await options.prompt.multiselect(
    "Outlook calendars to sync",
    calendars.map((calendar) => ({
      label: calendarLabel(calendar),
      value: calendar.id,
    })),
    defaultCalendarIds.length > 0 ? defaultCalendarIds : [calendars[0].id],
  );

  config.calendars = selectedIds.map((id) => {
    const calendar = calendars.find((entry) => entry.id === id);
    return {
      id,
      name: calendar?.name,
    };
  });
  await save();
}

async function configureNotionDataSources(
  config: NolendarConfig,
  options: ConfigWizardOptions,
  save: () => Promise<void>,
): Promise<NotionDataSourceSchema> {
  const dataSources = await listDataSources(options.notion);
  config.notion.databaseId = await chooseDataSourceId(options.prompt, "Main Notion meeting data source", dataSources);
  await save();

  const meetingSchema = await options.notion.retrieveDataSource(config.notion.databaseId);

  if (await options.prompt.confirm("Configure a People data source for attendee relations?", true)) {
    config.notion.peopleDataSource = {
      databaseId: await chooseDataSourceId(options.prompt, "People Notion data source", dataSources, config.notion.databaseId),
      nameProperty: "Name",
      emailProperty: "Email Address",
      maxAttendeesPerMeeting: 10,
    };
    await save();

    const peopleSchema = await options.notion.retrieveDataSource(config.notion.peopleDataSource.databaseId);
    config.notion.peopleDataSource.nameProperty = requireSelectedProperty(await chooseProperty(
      options.prompt,
      peopleSchema,
      "People name property",
      "title",
      "Name",
      true,
    ));
    await save();
    config.notion.peopleDataSource.emailProperty = requireSelectedProperty(await chooseProperty(
      options.prompt,
      peopleSchema,
      "People email property",
      "email",
      "Email Address",
      true,
    ));
    await save();
  } else {
    delete config.notion.peopleDataSource;
    delete config.mapping.participants;
    await save();
  }

  return meetingSchema;
}

async function configureMappings(
  config: NolendarConfig,
  meetingSchema: NotionDataSourceSchema,
  options: ConfigWizardOptions,
  save: () => Promise<void>,
): Promise<void> {
  for (const step of MEETING_MAPPING_STEPS) {
    if (step.key === "participants" && !config.notion.peopleDataSource) {
      delete config.mapping.participants;
      await save();
      continue;
    }

    const question = `${step.label} property (${step.source})`;
    const selected = await chooseProperty(options.prompt, meetingSchema, question, step.type, step.defaultName, step.required);

    if (selected) {
      config.mapping[step.key] = selected;
    } else {
      delete config.mapping[step.key];
    }

    await save();
  }
}

async function configureNotionOptional(
  config: NolendarConfig,
  options: ConfigWizardOptions,
  save: () => Promise<void>,
): Promise<void> {
  const prompt = options.prompt;
  const previousDataSourceTemplate = config.notion.dataSourceTemplate;
  const previousTemplatePageId = config.notion.templatePageId;
  const templateMode = await prompt.select(
    "Meeting page template",
    [
      { label: "No template", value: "none" },
      { label: "Copy blocks from a template page", value: "page" },
      { label: "Use the default data source template", value: "default_data_source" },
      { label: "Use a specific data source template", value: "template_id" },
    ],
    config.notion.templatePageId ? "page" : config.notion.dataSourceTemplate?.type === "template_id" ? "template_id" : config.notion.dataSourceTemplate ? "default_data_source" : "none",
  );
  delete config.notion.templatePageId;
  delete config.notion.dataSourceTemplate;

  if (templateMode === "page") {
    config.notion.templatePageId = await prompt.ask("Template page ID", previousTemplatePageId);
  } else if (templateMode === "default_data_source") {
    config.notion.dataSourceTemplate = {
      type: "default",
      timezone: await prompt.ask("Template timezone", previousDataSourceTemplate?.timezone ?? currentTimeZone()),
    };
  } else if (templateMode === "template_id") {
    config.notion.dataSourceTemplate = {
      type: "template_id",
      templateId: await chooseDataSourceTemplateId(
        options,
        config.notion.databaseId,
        previousDataSourceTemplate?.templateId,
      ),
      timezone: await prompt.ask("Template timezone", previousDataSourceTemplate?.timezone ?? currentTimeZone()),
    };
  }
  await save();

  config.notion.defaultTags = splitList(await prompt.ask("Default tags", config.notion.defaultTags?.join(", ") ?? "meeting"));
  await save();
  config.notion.defaultAssigneeEmail = await optionalAnswer(prompt, "Default assignee email", config.notion.defaultAssigneeEmail);
  await save();

  const iconMode = await prompt.select(
    "Page icon",
    [
      { label: "No icon", value: "none" },
      { label: "Emoji", value: "emoji" },
      { label: "Native Notion icon", value: "icon" },
    ],
    config.notion.pageIcon?.type ?? "none",
  );

  if (iconMode === "none") {
    delete config.notion.pageIcon;
  } else if (iconMode === "emoji") {
    config.notion.pageIcon = { type: "emoji", emoji: await prompt.ask("Emoji", "📝") };
  } else {
    config.notion.pageIcon = {
      type: "icon",
      name: await prompt.ask("Notion icon name", config.notion.pageIcon?.type === "icon" ? config.notion.pageIcon.name : "calendar"),
      color: await optionalAnswer(prompt, "Notion icon color", config.notion.pageIcon?.type === "icon" ? config.notion.pageIcon.color : undefined),
    };
  }
  await save();

  config.notion.pageContent = {
    sections: await prompt.multiselect(
      "Generated page content sections",
      PAGE_CONTENT_SECTIONS.map((section) => ({ label: section, value: section })),
      config.notion.pageContent?.sections ?? PAGE_CONTENT_SECTIONS,
    ),
    insertAfterHeading: await optionalAnswer(
      prompt,
      "Insert generated content after template heading",
      config.notion.pageContent?.insertAfterHeading,
    ),
  };
  await save();

  const cancelAction = await prompt.select(
    "Canceled Outlook meetings",
    [
      { label: "Archive the Notion page", value: "archive" },
      { label: "Set a Notion status", value: "set_status" },
    ],
    config.notion.canceledMeetings?.action ?? "archive",
  );

  if (cancelAction === "archive") {
    config.notion.canceledMeetings = { action: "archive" };
  } else {
    config.notion.canceledMeetings = {
      action: "set_status",
      statusProperty: await prompt.ask(
        "Canceled status property",
        config.notion.canceledMeetings?.action === "set_status" ? config.notion.canceledMeetings.statusProperty : "Status",
      ),
      statusValue: await prompt.ask(
        "Canceled status value",
        config.notion.canceledMeetings?.action === "set_status" ? config.notion.canceledMeetings.statusValue : "Canceled",
      ),
    };
  }
  await save();
}

async function configureFilters(config: NolendarConfig, prompt: ConfigWizardPrompt, save: () => Promise<void>): Promise<void> {
  const filters: FiltersConfig = {
    ignoreDeclined: await prompt.confirm("Ignore declined meetings?", config.filters.ignoreDeclined),
    requireAttendees: await prompt.confirm("Require attendees?", config.filters.requireAttendees),
    ignorePersonal: await prompt.confirm("Ignore personal/private meetings?", config.filters.ignorePersonal),
    ignoreOptionalAttendance: await prompt.confirm("Ignore meetings where you are optional?", config.filters.ignoreOptionalAttendance),
    minDurationMinutes: parseOptionalNumber(await optionalAnswer(prompt, "Minimum duration minutes", config.filters.minDurationMinutes?.toString())),
    ignoreNames: splitList(await prompt.ask("Ignored exact meeting names", config.filters.ignoreNames?.join(", ") ?? "")),
    ignorePatterns: splitList(await prompt.ask("Ignored meeting name regex patterns", config.filters.ignorePatterns?.join(", ") ?? "")),
  };

  config.filters = filters;
  await save();
}

async function configureSync(config: NolendarConfig, prompt: ConfigWizardPrompt, save: () => Promise<void>): Promise<void> {
  config.sync.lookahead = await prompt.ask("Sync lookahead (today, 24h, 7d, 2w, 1m)", config.sync.lookahead) as NolendarConfig["sync"]["lookahead"];
  await save();
  config.sync.statePath = await prompt.ask("Sync state path", config.sync.statePath);
  await save();
}

async function ensureConfiguredProperties(config: NolendarConfig, options: ConfigWizardOptions): Promise<void> {
  const missingMeetingProperties = missingProperties(buildRequiredNotionProperties(config), await options.notion.retrieveDataSource(config.notion.databaseId));
  const missingPeopleProperties = config.notion.peopleDataSource
    ? missingProperties(
        buildRequiredPeopleDataSourceProperties(config),
        await options.notion.retrieveDataSource(config.notion.peopleDataSource.databaseId),
      )
    : [];

  if (missingMeetingProperties.length > 0 && await options.prompt.confirm("Create missing properties in the meeting data source?", true)) {
    await options.notion.ensureProperties(config.notion.databaseId, missingMeetingProperties);
  }

  if (config.notion.peopleDataSource && missingPeopleProperties.length > 0 && await options.prompt.confirm("Create missing properties in the People data source?", true)) {
    await options.notion.ensureProperties(config.notion.peopleDataSource.databaseId, missingPeopleProperties);
  }
}

function missingProperties(required: RequiredNotionProperty[], schema: NotionDataSourceSchema): RequiredNotionProperty[] {
  return required.filter((property) => !schema.properties[property.name]);
}

async function listDataSources(notion: NotionClient): Promise<NotionDataSourceSummary[]> {
  if (notion.listDataSources) {
    return notion.listDataSources();
  }

  return [];
}

async function listDataSourceTemplates(
  notion: NotionClient,
  dataSourceId: string,
): Promise<NotionDataSourceTemplateSummary[]> {
  if (notion.listDataSourceTemplates) {
    return notion.listDataSourceTemplates(dataSourceId);
  }

  return [];
}

async function chooseDataSourceTemplateId(
  options: ConfigWizardOptions,
  dataSourceId: string,
  previousTemplateId?: string,
): Promise<string> {
  const templates = await listDataSourceTemplates(options.notion, dataSourceId);
  const choices = templates.map((template) => ({
    label: `${template.name}${template.isDefault ? " [default]" : ""} (${template.id})`,
    value: template.id,
  }));

  if (choices.length === 0) {
    return options.prompt.ask("Data source template ID", previousTemplateId);
  }

  return options.prompt.select(
    "Data source template",
    [...choices, { label: "Enter an ID manually", value: "__manual__" }],
    previousTemplateId ?? choices.find((choice) => templates.find((template) => template.id === choice.value)?.isDefault)?.value,
  ).then((selected) => {
    if (selected === "__manual__") {
      return options.prompt.ask("Data source template ID", previousTemplateId);
    }

    return selected;
  });
}

async function chooseDataSourceId(
  prompt: ConfigWizardPrompt,
  question: string,
  dataSources: NotionDataSourceSummary[],
  previousId?: string,
): Promise<string> {
  const choices = dataSources
    .filter((dataSource) => dataSource.id !== previousId)
    .map((dataSource) => ({
      label: dataSource.title ? `${dataSource.title} (${dataSource.id})` : dataSource.id,
      value: dataSource.id,
    }));

  if (choices.length === 0) {
    return prompt.ask(`${question} ID`);
  }

  return prompt.select(question, [...choices, { label: "Enter an ID manually", value: "__manual__" }]).then((selected) => {
    if (selected === "__manual__") {
      return prompt.ask(`${question} ID`);
    }

    return selected;
  });
}

async function chooseProperty(
  prompt: ConfigWizardPrompt,
  schema: NotionDataSourceSchema,
  question: string,
  expectedType: RequiredNotionProperty["type"],
  defaultName: string,
  required: boolean,
): Promise<string | undefined> {
  const compatible = Object.values(schema.properties)
    .filter((property) => property.type === expectedType)
    .sort((left, right) => left.name.localeCompare(right.name));
  const choices = compatible.map((property) => ({ label: `${property.name} (${property.type})`, value: property.name }));
  const selected = await prompt.select(
    question,
    [
      ...choices,
      { label: `Create "${defaultName}" (${expectedType})`, value: "__create__" },
      ...(required ? [] : [{ label: "Skip this mapping", value: "__skip__" }]),
    ],
    compatible.find((property) => property.name === defaultName)?.name ?? "__create__",
  );

  if (selected === "__skip__") {
    return undefined;
  }

  if (selected === "__create__") {
    return prompt.ask("New property name", defaultName);
  }

  return selected;
}

function defaultConfig(): NolendarConfig {
  return {
    microsoft: {
      tenant: "common",
      authMode: "device_code",
    },
    notion: {
      databaseId: "",
      peopleDataSource: {
        databaseId: "",
        nameProperty: "Name",
        emailProperty: "Email Address",
        maxAttendeesPerMeeting: 10,
      },
      canceledMeetings: {
        action: "archive",
      },
      defaultTags: ["meeting"],
      pageContent: {
        sections: PAGE_CONTENT_SECTIONS,
      },
    },
    calendars: [],
    filters: {
      ignoreDeclined: true,
      requireAttendees: false,
      ignorePersonal: false,
      ignoreOptionalAttendance: false,
    },
    mapping: {
      title: "Name",
      due: "Due",
      eventId: "Outlook Event ID",
      changeKey: "Outlook ChangeKey",
    },
    sync: {
      lookahead: "today",
      statePath: ".nolendar/state.json",
    },
  };
}

async function prepareConfigPath(configPath: string, force = false): Promise<void> {
  try {
    await access(configPath);
    if (!force) {
      throw new Error(`Config file already exists at ${configPath}. Use --force to overwrite it.`);
    }
  } catch (error) {
    if (!isNotFoundError(error)) {
      throw error;
    }
  }

  await mkdir(path.dirname(configPath), { recursive: true });
}

async function writeConfig(configPath: string, config: NolendarConfig): Promise<void> {
  await writeFile(configPath, stringify(config), "utf8");
}

function calendarLabel(calendar: AvailableCalendar): string {
  const tags = calendar.isDefaultCalendar ? " [default]" : "";
  const owner = calendar.ownerName || calendar.ownerAddress ? ` - ${calendar.ownerName ?? calendar.ownerAddress}` : "";

  return `${calendar.name}${tags}${owner}`;
}

async function optionalAnswer(prompt: ConfigWizardPrompt, question: string, defaultValue?: string): Promise<string | undefined> {
  const answer = await prompt.ask(question, defaultValue ?? "");

  return answer.trim() || undefined;
}

function splitList(value: string): string[] {
  return value
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function parseOptionalNumber(value: string | undefined): number | undefined {
  if (!value) {
    return undefined;
  }

  const parsed = Number.parseInt(value, 10);
  if (!Number.isFinite(parsed)) {
    return undefined;
  }

  return parsed;
}

function requireSelectedProperty(value: string | undefined): string {
  if (!value) {
    throw new Error("A property selection is required.");
  }

  return value;
}

function currentTimeZone(): string {
  return Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
}

function isNotFoundError(error: unknown): error is NodeJS.ErrnoException {
  return typeof error === "object" && error !== null && "code" in error && error.code === "ENOENT";
}
