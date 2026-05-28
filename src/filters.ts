import type { FiltersConfig } from "./domain/config.js";
import type { Meeting } from "./domain/meeting.js";

export function shouldSyncMeeting(meeting: Meeting, filters: FiltersConfig): boolean {
  return getMeetingFilterReason(meeting, filters) === undefined;
}

export function getMeetingFilterReason(meeting: Meeting, filters: FiltersConfig): string | undefined {
  if (filters.ignoreNames?.includes(meeting.title)) {
    return "ignoreNames";
  }

  const ignorePatterns = filters.ignorePatterns ?? filters.ignoreNamePatterns;

  if (ignorePatterns?.some((pattern) => new RegExp(pattern).test(meeting.title))) {
    return "ignorePatterns";
  }

  if (filters.ignoreDeclined && meeting.responseStatus === "declined") {
    return "ignoreDeclined";
  }

  if (filters.ignorePersonal && isPersonalMeeting(meeting)) {
    return "ignorePersonal";
  }

  if (filters.ignoreOptionalAttendance && meeting.isOptionalForOwner) {
    return "ignoreOptionalAttendance";
  }

  if (filters.requireAttendees && meeting.attendees.length === 0) {
    return "requireAttendees";
  }

  if (filters.minDurationMinutes !== undefined && meetingDurationMinutes(meeting) < filters.minDurationMinutes) {
    return "minDurationMinutes";
  }

  return undefined;
}

function meetingDurationMinutes(meeting: Meeting): number {
  const start = new Date(meeting.start);
  const end = new Date(meeting.end);

  return Math.max(0, (end.getTime() - start.getTime()) / (60 * 1000));
}

function isPersonalMeeting(meeting: Meeting): boolean {
  const sensitivity = meeting.sensitivity?.toLowerCase();
  return sensitivity === "personal" || sensitivity === "private";
}
