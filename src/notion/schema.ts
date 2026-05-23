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
  return [
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
}

export function validateNotionSchema(
  config: NolendarConfig,
  schema: NotionDataSourceSchema,
): NotionSchemaValidationResult {
  const required = buildRequiredNotionProperties(config);
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
