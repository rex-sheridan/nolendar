import { describe, expect, it, vi } from "vitest";

import { ApiNotionClient } from "../src/notion/api-notion-client.js";

describe("ApiNotionClient.getDefaultAssigneeUserId", () => {
  it("resolves a people property assignee from defaultAssigneeEmail", async () => {
    const client = new ApiNotionClient("token", {
      dataSources: {
        retrieve: vi.fn(),
        update: vi.fn(),
        query: vi.fn(),
      },
      pages: {
        create: vi.fn(),
        update: vi.fn(),
      },
      users: {
        me: vi.fn(async () => ({
          id: "bot-user",
          type: "bot",
        })),
        list: vi.fn(async () => ({
          results: [
            {
              id: "user-123",
              type: "person",
              person: {
                email: "me@example.com",
              },
            },
          ],
          has_more: false,
          next_cursor: null,
        })),
      },
    });

    await expect(client.getDefaultAssigneeUserId("me@example.com")).resolves.toBe("user-123");
  });

  it("falls back to the authenticated user when email lookup does not match", async () => {
    const client = new ApiNotionClient("token", {
      dataSources: {
        retrieve: vi.fn(),
        update: vi.fn(),
        query: vi.fn(),
      },
      pages: {
        create: vi.fn(),
        update: vi.fn(),
      },
      users: {
        me: vi.fn(async () => ({
          id: "person-456",
          type: "person",
        })),
        list: vi.fn(async () => ({
          results: [],
          has_more: false,
          next_cursor: null,
        })),
      },
    });

    await expect(client.getDefaultAssigneeUserId("missing@example.com")).resolves.toBe("person-456");
  });
});
