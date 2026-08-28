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

const en: ICalendarLabels = {
  today: 'Today', previous: 'Previous', next: 'Next', day: 'Day', week: 'Week', month: 'Month',
  allDay: 'All day', open: 'Open', join: 'Join', close: 'Close', details: 'Appointment details',
  location: 'Location', organizer: 'Organizer', attendees: 'Attendees', source: 'Source',
  description: 'Description', noEvents: 'No appointments', navigation: 'Calendar navigation', view: 'Calendar view'
};

const nl: ICalendarLabels = {
  today: 'Vandaag', previous: 'Vorige', next: 'Volgende', day: 'Dag', week: 'Week', month: 'Maand',
  allDay: 'Hele dag', open: 'Openen', join: 'Deelnemen', close: 'Sluiten', details: 'Afspraakdetails',
  location: 'Locatie', organizer: 'Organisator', attendees: 'Deelnemers', source: 'Bron',
  description: 'Beschrijving', noEvents: 'Geen afspraken', navigation: 'Kalendernavigatie', view: 'Kalenderweergave'
};

export function getCalendarLabels(): ICalendarLabels {
  const language = typeof navigator === 'undefined' ? 'en' : navigator.language.toLowerCase();
  return language.indexOf('nl') === 0 ? nl : en;
}
