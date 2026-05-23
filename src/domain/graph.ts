export interface GraphAttendee {
  emailAddress?: {
    name?: string | null;
    address?: string | null;
  } | null;
  type?: string | null;
}

export interface GraphDateTimeTimeZone {
  dateTime?: string | null;
  timeZone?: string | null;
}

export interface GraphEvent {
  id?: string | null;
  changeKey?: string | null;
  subject?: string | null;
  start?: GraphDateTimeTimeZone | null;
  end?: GraphDateTimeTimeZone | null;
  organizer?: {
    emailAddress?: {
      name?: string | null;
      address?: string | null;
    } | null;
  } | null;
  attendees?: GraphAttendee[] | null;
  onlineMeeting?: {
    joinUrl?: string | null;
  } | null;
  onlineMeetingUrl?: string | null;
  webLink?: string | null;
  bodyPreview?: string | null;
  body?: {
    contentType?: string | null;
    content?: string | null;
  } | null;
  responseStatus?: {
    response?: string | null;
  } | null;
  type?: string | null;
  isCancelled?: boolean | null;
  recurrence?: object | null;
}
