import type { NolendarConfig } from "../domain/config.js";
import type { NotionClient } from "./client.js";
import { formatNotionSchemaIssues, validateNotionSchema } from "./schema.js";

export class NotionSchemaError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "NotionSchemaError";
  }
}

export async function validateOrEnsureNotionSchema(
  config: NolendarConfig,
  notion: NotionClient,
  options: { ensureProperties?: boolean } = {},
): Promise<void> {
  let schema = await notion.retrieveDataSource(config.notion.databaseId);
  let result = validateNotionSchema(config, schema);

  if (!result.valid && options.ensureProperties && result.missing.length > 0 && result.mismatched.length === 0) {
    await notion.ensureProperties(config.notion.databaseId, result.missing);
    schema = await notion.retrieveDataSource(config.notion.databaseId);
    result = validateNotionSchema(config, schema);
  }

  if (!result.valid) {
    throw new NotionSchemaError(formatNotionSchemaIssues(result).join("\n"));
  }
}
