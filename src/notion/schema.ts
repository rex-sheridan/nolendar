import type { NolendarConfig } from "../domain/config.js";
import type { NotionDataSourceSchema, RequiredNotionProperty } from "../domain/notion.js";

export interface NotionSchemaMismatch {
  name: string;
  expectedType: string;
  actualType: string;
}

export interface NotionSchemaValidationResult {
  valid: boolean;
  missing: RequiredNotionProperty[];
  mismatched: NotionSchemaMismatch[];
}

export function buildRequiredNotionProperties(config: NolendarConfig): RequiredNotionProperty[] {
  const required: RequiredNotionProperty[] = [
    {
      name: config.mapping.title,
      type: "title",
    },
    {
      name: config.mapping.due,
      type: "date",
    },
    {
      name: config.mapping.eventId,
      type: "rich_text",
    },
    {
      name: config.mapping.changeKey,
      type: "rich_text",
    },
  ];

  if (config.mapping.eventLink) {
    required.push({
      name: config.mapping.eventLink,
      type: "url",
    });
  }

  if (config.mapping.tags && config.notion.defaultTags && config.notion.defaultTags.length > 0) {
    required.push({
      name: config.mapping.tags,
      type: "multi_select",
    });
  }

  if (config.mapping.assignee) {
    required.push({
      name: config.mapping.assignee,
      type: "people",
    });
  }

  if (config.mapping.participants) {
    required.push({
      name: config.mapping.participants,
      type: "relation",
      relationDataSourceId: config.notion.peopleDataSource?.databaseId,
    });
  }

  const canceledMeetings = config.notion.canceledMeetings ?? { action: "archive" };

  if (canceledMeetings.action === "set_status") {
    pushRequiredProperty(required, {
      name: canceledMeetings.statusProperty,
      type: "status",
    });
  }

  if (config.notion.completedMeetings) {
    pushRequiredProperty(required, {
      name: config.notion.completedMeetings.statusProperty,
      type: "status",
    });
  }

  return required;
}

function pushRequiredProperty(required: RequiredNotionProperty[], property: RequiredNotionProperty): void {
  if (!required.some((existing) => existing.name === property.name)) {
    required.push(property);
  }
}

export function buildRequiredPeopleDataSourceProperties(config: NolendarConfig): RequiredNotionProperty[] {
  if (!config.notion.peopleDataSource) {
    return [];
  }

  return [
    {
      name: config.notion.peopleDataSource.nameProperty,
      type: "title",
    },
    {
      name: config.notion.peopleDataSource.emailProperty,
      type: "email",
    },
  ];
}

export function validateNotionSchema(
  config: NolendarConfig,
  schema: NotionDataSourceSchema,
): NotionSchemaValidationResult {
  return validateSchemaRequirements(buildRequiredNotionProperties(config), schema);
}

export function validatePeopleDataSourceSchema(
  config: NolendarConfig,
  schema: NotionDataSourceSchema,
): NotionSchemaValidationResult {
  return validateSchemaRequirements(buildRequiredPeopleDataSourceProperties(config), schema);
}

function validateSchemaRequirements(
  required: RequiredNotionProperty[],
  schema: NotionDataSourceSchema,
): NotionSchemaValidationResult {
  const missing: RequiredNotionProperty[] = [];
  const mismatched: NotionSchemaMismatch[] = [];

  for (const property of required) {
    const actual = schema.properties[property.name];

    if (!actual) {
      missing.push(property);
      continue;
    }

    if (actual.type !== property.type) {
      mismatched.push({
        name: property.name,
        expectedType: property.type,
        actualType: actual.type,
      });
    }
  }

  return {
    valid: missing.length === 0 && mismatched.length === 0,
    missing,
    mismatched,
  };
}

export function formatNotionSchemaIssues(result: NotionSchemaValidationResult): string[] {
  return [
    ...result.missing.map((property) => `Missing property: ${property.name} (${property.type})`),
    ...result.mismatched.map(
      (property) =>
        `Property type mismatch: ${property.name} expected ${property.expectedType} but found ${property.actualType}`,
    ),
  ];
}
