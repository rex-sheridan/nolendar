export interface MeetingAttendee {
  name?: string;
  email?: string;
  optional: boolean;
}

export interface Meeting {
  id: string;
  changeKey: string;
  calendarId: string;
  calendarName?: string;
  title: string;
  start: string;
  end: string;
  organizer?: string;
  attendees: MeetingAttendee[];
  meetingLink?: string;
  eventLink?: string;
  agenda?: string;
  details?: string;
  responseStatus?: string;
  isCancelled: boolean;
  isRecurring: boolean;
}
