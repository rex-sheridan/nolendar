import { compactLogIds, type LogCompactionOptions } from "./log-compaction.js";

export interface ApiTimingEntry {
  service: "graph" | "notion";
  operation: string;
  detail?: string;
  durationMs: number;
  status?: string;
  count?: number;
}

export interface ApiTimingReporter {
  record(entry: ApiTimingEntry): void;
}

export function createConsoleTimingReporter(
  sink: Pick<Console, "log">,
  options: LogCompactionOptions = {},
): ApiTimingReporter {
  return {
    record(entry) {
      const status = entry.status ? ` ${entry.status}` : "";
      const operation = compactLogIds(entry.operation, options);
      const detail = entry.detail ? ` ${compactLogIds(entry.detail, options)}` : "";
      const count = entry.count === undefined ? "" : ` count=${entry.count}`;
      sink.log(`[timings] ${entry.service} ${operation}${detail}${status}${count} ${entry.durationMs}ms`);
    },
  };
}
