import type { Meeting } from "./domain/meeting.js";

export function formatMeeting(meeting: Meeting): string[] {
  const title = meeting.title || "(untitled meeting)";
  const organizer = meeting.organizer ?? "Unknown organizer";
  const attendees = meeting.attendees.length > 0 ? meeting.attendees.map(formatAttendee).join(", ") : "None";
  const markers = [
    meeting.isRecurring ? "recurring" : undefined,
    meeting.isCancelled ? "cancelled" : undefined,
    meeting.responseStatus,
  ].filter(Boolean);

  return [
    `${meeting.start} | ${title}`,
    `  Calendar: ${meeting.calendarName ?? meeting.calendarId}`,
    `  Organizer: ${organizer}`,
    `  Attendees: ${attendees}`,
    `  Meeting Link: ${meeting.meetingLink ?? "None"}`,
    `  Event Link: ${meeting.eventLink ?? "None"}`,
    `  Flags: ${markers.length > 0 ? markers.join(", ") : "none"}`,
  ];
}

function formatAttendee(attendee: Meeting["attendees"][number]): string {
  const identity = attendee.name ?? attendee.email ?? "Unknown attendee";
  return attendee.optional ? `${identity} (optional)` : identity;
}
