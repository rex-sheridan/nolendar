import type { FiltersConfig } from "./domain/config.js";
import type { Meeting } from "./domain/meeting.js";

export function shouldSyncMeeting(meeting: Meeting, filters: FiltersConfig): boolean {
  if (filters.ignoreNames?.includes(meeting.title)) {
    return false;
  }

  const ignorePatterns = filters.ignorePatterns ?? filters.ignoreNamePatterns;

  if (ignorePatterns?.some((pattern) => new RegExp(pattern).test(meeting.title))) {
    return false;
  }

  if (filters.ignoreDeclined && meeting.responseStatus === "declined") {
    return false;
  }

  if (filters.ignorePersonal && isPersonalMeeting(meeting)) {
    return false;
  }

  if (filters.ignoreOptionalAttendance && meeting.isOptionalForOwner) {
    return false;
  }

  if (filters.requireAttendees && meeting.attendees.length === 0) {
    return false;
  }

  if (filters.minDurationMinutes !== undefined && meetingDurationMinutes(meeting) < filters.minDurationMinutes) {
    return false;
  }

  return true;
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
