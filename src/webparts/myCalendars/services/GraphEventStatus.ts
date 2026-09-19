import type { CalendarAvailabilityStatus, CalendarResponseStatus } from '../models/ICalendarEvent';

const availabilityStatuses = new Set<CalendarAvailabilityStatus>(['free', 'tentative', 'busy', 'oof', 'workingElsewhere', 'unknown']);
const responseStatuses = new Set<CalendarResponseStatus>(['none', 'organizer', 'tentativelyAccepted', 'accepted', 'declined', 'notResponded']);

export function normalizeAvailabilityStatus(value: unknown): CalendarAvailabilityStatus | undefined {
  if (typeof value !== 'string') return undefined;
  return availabilityStatuses.has(value as CalendarAvailabilityStatus) ? value as CalendarAvailabilityStatus : 'unknown';
}

export function normalizeResponseStatus(value: unknown): CalendarResponseStatus | undefined {
  return typeof value === 'string' && responseStatuses.has(value as CalendarResponseStatus)
    ? value as CalendarResponseStatus
    : undefined;
}
