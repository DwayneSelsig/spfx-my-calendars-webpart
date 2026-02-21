export interface IAppointment {
  id: string;
  title: string;
  description?: string;
  location?: string;
  startDate: Date;
  endDate: Date;
  isAllDay: boolean;
  sourceId: string;
  color: string;
  organizer?: string;
  attendees?: string[];
}
