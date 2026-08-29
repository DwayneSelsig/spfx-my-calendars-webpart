import * as strings from 'MyCalendarsWebPartStrings';

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
    view: strings.CalendarViewLabel
  };
}
