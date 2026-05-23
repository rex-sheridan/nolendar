import type { CalendarConfig, LookaheadWindow, NolendarConfig } from "./domain/config.js";
import type { Meeting } from "./domain/meeting.js";
import type { Clock } from "./clock.js";
import { systemClock } from "./clock.js";

export interface CalendarWindow {
  start: string;
  end: string;
}

export interface MeetingSource {
  listMeetings(args: { calendar: CalendarConfig; window: CalendarWindow }): Promise<Meeting[]>;
}

export interface ListMeetingsDependencies {
  meetingSource: MeetingSource;
  clock?: Clock;
}

export interface ListMeetingsResult {
  lookahead: LookaheadWindow;
  window: CalendarWindow;
  meetings: Meeting[];
}

export async function listMeetings(
  config: NolendarConfig,
  lookahead: LookaheadWindow,
  deps: ListMeetingsDependencies,
): Promise<ListMeetingsResult> {
  const clock = deps.clock ?? systemClock;
  const window = resolveWindow(lookahead, clock);
  const meetingsByCalendar = await Promise.all(
    config.calendars.map((calendar) =>
      deps.meetingSource.listMeetings({
        calendar,
        window,
      }),
    ),
  );

  const meetings = meetingsByCalendar
    .flat()
    .sort((left, right) => left.start.localeCompare(right.start) || left.title.localeCompare(right.title));

  return {
    lookahead,
    window,
    meetings,
  };
}

export function resolveWindow(lookahead: LookaheadWindow, clock: Clock = systemClock): CalendarWindow {
  const now = clock.now();
  const start = new Date(now);
  let end: Date;

  switch (lookahead) {
    case "today": {
      start.setUTCHours(0, 0, 0, 0);
      end = new Date(start);
      end.setUTCDate(end.getUTCDate() + 1);
      break;
    }
    case "24h": {
      end = new Date(start.getTime() + 24 * 60 * 60 * 1000);
      break;
    }
    case "7d": {
      end = new Date(start.getTime() + 7 * 24 * 60 * 60 * 1000);
      break;
    }
  }

  return {
    start: start.toISOString(),
    end: end.toISOString(),
  };
}
