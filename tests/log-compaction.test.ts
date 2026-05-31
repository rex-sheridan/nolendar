import { describe, expect, it } from "vitest";

import { compactLogIds } from "../src/log-compaction.js";

describe("compactLogIds", () => {
  it("leaves log lines unchanged unless enabled", () => {
    expect(compactLogIds("data_source_id=31086680-d5d6-81df-a454-000b46830e24")).toBe(
      "data_source_id=31086680-d5d6-81df-a454-000b46830e24",
    );
  });

  it("compacts UUID-like IDs to the first and last significant segments", () => {
    expect(compactLogIds("data_source_id=31086680-d5d6-81df-a454-000b46830e24", { compactIds: true })).toBe(
      "data_source_id=31086680...000b46830e24",
    );
  });

  it("compacts long opaque IDs in URLs and key-value fields", () => {
    const input = [
      "GET /v1.0/me/calendars/AQMkADg5NmM4ZTY1LTQ5ZDYtNGI3ZS1iZjRh/calendarView/delta",
      "eventId=AAMkADg5NmM4ZTY1LTQ5ZDYtNGI3ZS1iZjRhLWYyNzg2N2I5YTc4NQ==",
    ].join(" ");

    expect(compactLogIds(input, { compactIds: true })).toBe(
      "GET /v1.0/me/calendars/AQMkADg5Nm...I3ZS1iZjRh/calendarView/delta eventId=AAMkADg5Nm...I5YTc4NQ==",
    );
  });

  it("does not compact encoded select lists", () => {
    const select = "%24select=id%2CchangeKey%2Csubject%2Cstart%2Cend%2Corganizer";

    expect(compactLogIds(`GET /calendarView?${select}`, { compactIds: true })).toBe(`GET /calendarView?${select}`);
  });
});
