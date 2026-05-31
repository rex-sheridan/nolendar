import { describe, expect, it, vi } from "vitest";

import { createConsoleTimingReporter } from "../src/api-timing.js";

describe("createConsoleTimingReporter", () => {
  it("prints record counts when provided", () => {
    const sink = {
      log: vi.fn(),
    };
    const reporter = createConsoleTimingReporter(sink);

    reporter.record({
      service: "graph",
      operation: "GET /v1.0/me/events",
      status: "200",
      count: 3,
      durationMs: 42,
    });

    expect(sink.log).toHaveBeenCalledWith("[timings] graph GET /v1.0/me/events 200 count=3 42ms");
  });

  it("compacts IDs in timing operation and detail fields when requested", () => {
    const sink = {
      log: vi.fn(),
    };
    const reporter = createConsoleTimingReporter(sink, { compactIds: true });

    reporter.record({
      service: "notion",
      operation: "dataSources.query",
      detail: "data_source_id=31086680-d5d6-81df-a454-000b46830e24 filter=Due:date",
      status: "ok",
      count: 4,
      durationMs: 99,
    });

    expect(sink.log).toHaveBeenCalledWith(
      "[timings] notion dataSources.query data_source_id=31086680...000b46830e24 filter=Due:date ok count=4 99ms",
    );
  });
});
