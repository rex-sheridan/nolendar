const NOTION_TEXT_CONTENT_LIMIT = 2000;

interface RichText {
  type: "text";
  text: {
    content: string;
    link?: {
      url: string;
    };
  };
  annotations?: {
    bold?: boolean;
    italic?: boolean;
    strikethrough?: boolean;
    code?: boolean;
  };
}

export function buildMarkdownBlocks(markdown: string): unknown[] {
  const blocks: unknown[] = [];
  const lines = markdown.replace(/\r\n/g, "\n").split("\n");
  let paragraph: string[] = [];
  let codeFence: { language?: string; lines: string[] } | undefined;

  const flushParagraph = () => {
    const content = paragraph.join("\n").trim();
    paragraph = [];

    if (!content) {
      return;
    }

    blocks.push(...splitTextContent(content).map((chunk) => paragraphBlock(parseRichText(chunk))));
  };

  for (const line of lines) {
    const fenceMatch = line.match(/^```([A-Za-z0-9_-]+)?\s*$/);
    if (fenceMatch) {
      if (codeFence) {
        flushParagraph();
        blocks.push(codeBlock(codeFence.lines.join("\n"), codeFence.language));
        codeFence = undefined;
      } else {
        flushParagraph();
        codeFence = {
          language: fenceMatch[1],
          lines: [],
        };
      }
      continue;
    }

    if (codeFence) {
      codeFence.lines.push(line);
      continue;
    }

    if (!line.trim()) {
      flushParagraph();
      continue;
    }

    const headingMatch = line.match(/^(#{1,3})\s+(.+)$/);
    if (headingMatch) {
      flushParagraph();
      blocks.push(headingBlock(headingMatch[1].length, parseRichText(headingMatch[2].trim())));
      continue;
    }

    if (/^---+$/.test(line.trim())) {
      flushParagraph();
      blocks.push({ object: "block", type: "divider", divider: {} });
      continue;
    }

    const todoMatch = line.match(/^[-*]\s+\[([ xX])\]\s+(.+)$/);
    if (todoMatch) {
      flushParagraph();
      blocks.push(toDoBlock(parseRichText(todoMatch[2].trim()), todoMatch[1].toLowerCase() === "x"));
      continue;
    }

    const bulletMatch = line.match(/^[-*]\s+(.+)$/);
    if (bulletMatch) {
      flushParagraph();
      blocks.push(listItemBlock("bulleted_list_item", parseRichText(bulletMatch[1].trim())));
      continue;
    }

    const numberedMatch = line.match(/^\d+[.)]\s+(.+)$/);
    if (numberedMatch) {
      flushParagraph();
      blocks.push(listItemBlock("numbered_list_item", parseRichText(numberedMatch[1].trim())));
      continue;
    }

    const quoteMatch = line.match(/^>\s?(.+)$/);
    if (quoteMatch) {
      flushParagraph();
      blocks.push(quoteBlock(parseRichText(quoteMatch[1].trim())));
      continue;
    }

    paragraph.push(line);
  }

  flushParagraph();

  if (codeFence) {
    blocks.push(codeBlock(codeFence.lines.join("\n"), codeFence.language));
  }

  return blocks;
}

function headingBlock(level: number, richText: RichText[]): unknown {
  const type = level === 1 ? "heading_1" : level === 2 ? "heading_2" : "heading_3";

  return {
    object: "block",
    type,
    [type]: {
      rich_text: richText,
    },
  };
}

function paragraphBlock(richText: RichText[]): unknown {
  return {
    object: "block",
    type: "paragraph",
    paragraph: {
      rich_text: richText,
    },
  };
}

function listItemBlock(type: "bulleted_list_item" | "numbered_list_item", richText: RichText[]): unknown {
  return {
    object: "block",
    type,
    [type]: {
      rich_text: richText,
    },
  };
}

function toDoBlock(richText: RichText[], checked: boolean): unknown {
  return {
    object: "block",
    type: "to_do",
    to_do: {
      rich_text: richText,
      checked,
    },
  };
}

function quoteBlock(richText: RichText[]): unknown {
  return {
    object: "block",
    type: "quote",
    quote: {
      rich_text: richText,
    },
  };
}

function codeBlock(content: string, language?: string): unknown {
  return {
    object: "block",
    type: "code",
    code: {
      rich_text: splitTextContent(content || " ").map((chunk) => textBlock(chunk)),
      language: language ?? "plain text",
    },
  };
}

function parseRichText(markdown: string): RichText[] {
  const tokens: RichText[] = [];
  const tokenPattern = /(\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)|\*\*([^*]+)\*\*|\*([^*]+)\*|~~([^~]+)~~|`([^`]+)`)/g;
  let cursor = 0;
  let match: RegExpExecArray | null;

  while ((match = tokenPattern.exec(markdown)) !== null) {
    if (match.index > cursor) {
      tokens.push(textBlock(markdown.slice(cursor, match.index)));
    }

    if (match[2] && match[3]) {
      tokens.push(textBlock(match[2], { link: match[3] }));
    } else if (match[4]) {
      tokens.push(textBlock(match[4], { bold: true }));
    } else if (match[5]) {
      tokens.push(textBlock(match[5], { italic: true }));
    } else if (match[6]) {
      tokens.push(textBlock(match[6], { strikethrough: true }));
    } else if (match[7]) {
      tokens.push(textBlock(match[7], { code: true }));
    }

    cursor = match.index + match[0].length;
  }

  if (cursor < markdown.length) {
    tokens.push(textBlock(markdown.slice(cursor)));
  }

  return tokens.length > 0 ? tokens : [textBlock(" ")];
}

function textBlock(
  content: string,
  options: {
    link?: string;
    bold?: boolean;
    italic?: boolean;
    strikethrough?: boolean;
    code?: boolean;
  } = {},
): RichText {
  const annotations: RichText["annotations"] = {};

  if (options.bold) {
    annotations.bold = true;
  }

  if (options.italic) {
    annotations.italic = true;
  }

  if (options.strikethrough) {
    annotations.strikethrough = true;
  }

  if (options.code) {
    annotations.code = true;
  }

  return {
    type: "text",
    text: {
      content,
      ...(options.link ? { link: { url: options.link } } : {}),
    },
    ...(Object.keys(annotations).length > 0 ? { annotations } : {}),
  };
}

function splitTextContent(content: string): string[] {
  if (content.length <= NOTION_TEXT_CONTENT_LIMIT) {
    return [content];
  }

  const chunks: string[] = [];
  let remaining = content;

  while (remaining.length > NOTION_TEXT_CONTENT_LIMIT) {
    const splitAt = findSplitIndex(remaining);
    chunks.push(remaining.slice(0, splitAt).trimEnd());
    remaining = remaining.slice(splitAt).trimStart();
  }

  if (remaining.length > 0) {
    chunks.push(remaining);
  }

  return chunks;
}

function findSplitIndex(value: string): number {
  const hardLimit = Math.min(value.length, NOTION_TEXT_CONTENT_LIMIT);
  const preferredBreak = Math.max(
    value.lastIndexOf("\n", hardLimit - 1),
    value.lastIndexOf(". ", hardLimit - 1),
    value.lastIndexOf(" ", hardLimit - 1),
  );

  return preferredBreak > 0 ? Math.min(preferredBreak + 1, hardLimit) : hardLimit;
}
