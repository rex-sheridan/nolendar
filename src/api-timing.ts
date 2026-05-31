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
): ApiTimingReporter {
  return {
    record(entry) {
      const status = entry.status ? ` ${entry.status}` : "";
      const detail = entry.detail ? ` ${entry.detail}` : "";
      const count = entry.count === undefined ? "" : ` count=${entry.count}`;
      sink.log(`[timings] ${entry.service} ${entry.operation}${detail}${status}${count} ${entry.durationMs}ms`);
    },
  };
}
