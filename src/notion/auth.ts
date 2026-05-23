export class NotionAuthError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "NotionAuthError";
  }
}

export function resolveNotionAuthToken(env: NodeJS.ProcessEnv = process.env): string {
  const token = env.NOTION_TOKEN ?? env.NOTION_API_KEY;

  if (!token) {
    throw new NotionAuthError("Missing `NOTION_TOKEN` or `NOTION_API_KEY` for Notion API access.");
  }

  return token;
}
