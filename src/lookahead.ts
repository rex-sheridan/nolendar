import type { LookaheadWindow } from "./domain/config.js";

const RELATIVE_LOOKAHEAD_PATTERN = /^([1-9]\d*)([hdwm])$/;

export interface RelativeLookahead {
  quantity: number;
  unit: "h" | "d" | "w" | "m";
}

export function isValidLookaheadWindow(value: string): value is LookaheadWindow {
  return value === "today" || RELATIVE_LOOKAHEAD_PATTERN.test(value);
}

export function parseRelativeLookahead(value: LookaheadWindow): RelativeLookahead | null {
  if (value === "today") {
    return null;
  }

  const match = value.match(RELATIVE_LOOKAHEAD_PATTERN);

  if (!match) {
    throw new Error(`Invalid lookahead window: ${value}`);
  }

  return {
    quantity: Number.parseInt(match[1] ?? "", 10),
    unit: match[2] as RelativeLookahead["unit"],
  };
}
