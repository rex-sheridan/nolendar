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
});
