import type { CalendarSourceType } from './ICalendarSettings';

export interface ICalendarAttendee {
  id?: string;
  name?: string;
  email?: string;
  role?: string;
  imageUrl?: string;
}

export interface ICalendarOrganizer {
  name?: string;
  email?: string;
}

export type CalendarEventDescriptionFormat = 'plainText' | 'html';
export type CalendarAvailabilityStatus = 'free' | 'tentative' | 'busy' | 'oof' | 'workingElsewhere' | 'unknown';
export type CalendarResponseStatus = 'none' | 'organizer' | 'tentativelyAccepted' | 'accepted' | 'declined' | 'notResponded';

/** Canonical, read-only event model used by source adapters and renderers. */
export interface ICalendarEvent {
  id: string;
  sourceId: string;
  title: string;
  start: string;
  end: string;
  sourceType?: CalendarSourceType;
  sourceDisplayName?: string;
  sharePointSiteName?: string;
  showAs?: CalendarAvailabilityStatus;
  responseStatus?: CalendarResponseStatus;
  isFullDay?: boolean;
  attendees?: ICalendarAttendee[];
  organizer?: ICalendarOrganizer;
  category?: string;
  description?: string;
  /** Omitted values are treated as plain text. */
  descriptionFormat?: CalendarEventDescriptionFormat;
  location?: string;
  importance?: string;
  isOrganizer?: boolean;
  sensitivity?: string;
  type?: string;
  isOnlineMeeting?: boolean;
  joinUrl?: string;
  webLink?: string;
  color?: string;
  colorHex?: string;
  showSourceLogo?: boolean;
  sourceIconName?: string;
  percentComplete?: number;
  isDraft?: boolean;
  imageUrl?: string;
}
