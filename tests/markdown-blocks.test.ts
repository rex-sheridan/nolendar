import { describe, expect, it } from "vitest";

import { buildMarkdownBlocks } from "../src/notion/markdown-blocks.js";

describe("buildMarkdownBlocks", () => {
  it("converts common Markdown blocks and inline annotations into Notion blocks", () => {
    expect(
      buildMarkdownBlocks(
        [
          "## Key decisions",
          "- Confirm **scope**",
          "1. Ask *owner*",
          "- [ ] Send `brief`",
          "> Risk: [dependency](https://example.com)",
        ].join("\n"),
      ),
    ).toEqual([
      {
        object: "block",
        type: "heading_2",
        heading_2: {
          rich_text: [{ type: "text", text: { content: "Key decisions" } }],
        },
      },
      {
        object: "block",
        type: "bulleted_list_item",
        bulleted_list_item: {
          rich_text: [
            { type: "text", text: { content: "Confirm " } },
            { type: "text", text: { content: "scope" }, annotations: { bold: true } },
          ],
        },
      },
      {
        object: "block",
        type: "numbered_list_item",
        numbered_list_item: {
          rich_text: [
            { type: "text", text: { content: "Ask " } },
            { type: "text", text: { content: "owner" }, annotations: { italic: true } },
          ],
        },
      },
      {
        object: "block",
        type: "to_do",
        to_do: {
          rich_text: [
            { type: "text", text: { content: "Send " } },
            { type: "text", text: { content: "brief" }, annotations: { code: true } },
          ],
          checked: false,
        },
      },
      {
        object: "block",
        type: "quote",
        quote: {
          rich_text: [
            { type: "text", text: { content: "Risk: " } },
            {
              type: "text",
              text: {
                content: "dependency",
                link: {
                  url: "https://example.com",
                },
              },
            },
          ],
        },
      },
    ]);
  });
});
