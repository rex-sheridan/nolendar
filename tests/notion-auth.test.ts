import { describe, expect, it } from "vitest";

import { NotionAuthError, resolveNotionAuthToken } from "../src/notion/auth.js";

describe("resolveNotionAuthToken", () => {
  it("prefers NOTION_TOKEN", () => {
    expect(
      resolveNotionAuthToken({
        NOTION_TOKEN: "secret_123",
        NOTION_API_KEY: "secret_456",
      }),
    ).toBe("secret_123");
  });

  it("falls back to NOTION_API_KEY", () => {
    expect(
      resolveNotionAuthToken({
        NOTION_API_KEY: "secret_456",
      }),
    ).toBe("secret_456");
  });

  it("throws when no Notion token is configured", () => {
    expect(() => resolveNotionAuthToken({})).toThrowError(
      new NotionAuthError("Missing `NOTION_TOKEN` or `NOTION_API_KEY` for Notion API access."),
    );
  });
});
