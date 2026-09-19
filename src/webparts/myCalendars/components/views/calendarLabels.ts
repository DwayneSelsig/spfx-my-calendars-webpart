import * as strings from 'MyCalendarsWebPartStrings';
import type { CalendarAvailabilityStatus, CalendarResponseStatus } from '../../models/ICalendarEvent';

export interface ICalendarLabels {
  today: string;
  previous: string;
  next: string;
  day: string;
  week: string;
  month: string;
  allDay: string;
  open: string;
  join: string;
  close: string;
  details: string;
  location: string;
  organizer: string;
  attendees: string;
  source: string;
  description: string;
  noEvents: string;
  navigation: string;
  view: string;
  site: string;
  siteNameUnavailable: string;
  calendarAvailability: string;
  yourResponse: string;
}

export function getCalendarLabels(): ICalendarLabels {
  return {
    today: strings.TodayLabel,
    previous: strings.PreviousLabel,
    next: strings.NextLabel,
    day: strings.DayLabel,
    week: strings.WeekLabel,
    month: strings.MonthLabel,
    allDay: strings.AllDayLabel,
    open: strings.OpenLabel,
    join: strings.JoinLabel,
    close: strings.CloseLabel,
    details: strings.AppointmentDetailsLabel,
    location: strings.LocationLabel,
    organizer: strings.OrganizerLabel,
    attendees: strings.AttendeesLabel,
    source: strings.SourceLabel,
    description: strings.DescriptionLabel,
    noEvents: strings.NoAppointmentsLabel,
    navigation: strings.CalendarNavigationLabel,
    view: strings.CalendarViewLabel,
    site: strings.SiteLabel,
    siteNameUnavailable: strings.SiteNameUnavailableLabel,
    calendarAvailability: strings.CalendarAvailabilityLabel,
    yourResponse: strings.YourResponseLabel
  };
}

export function getAvailabilityLabel(status: CalendarAvailabilityStatus): string {
  return {
    free: strings.AvailabilityFreeLabel,
    tentative: strings.AvailabilityTentativeLabel,
    busy: strings.AvailabilityBusyLabel,
    oof: strings.AvailabilityOofLabel,
    workingElsewhere: strings.AvailabilityWorkingElsewhereLabel,
    unknown: strings.AvailabilityUnknownLabel
  }[status];
}

export function getResponseLabel(status: CalendarResponseStatus): string {
  return {
    none: strings.ResponseNoneLabel,
    organizer: strings.ResponseOrganizerLabel,
    tentativelyAccepted: strings.ResponseTentativelyAcceptedLabel,
    accepted: strings.ResponseAcceptedLabel,
    declined: strings.ResponseDeclinedLabel,
    notResponded: strings.ResponseNotRespondedLabel
  }[status];
}
