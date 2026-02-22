import { CalendarSourceType } from './ICalendarSettings';

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
  // Source metadata for display
  sourceType?: CalendarSourceType;
  showSourceLogo?: boolean;
  // Planner-specific fields
  percentComplete?: number;
  // Teams Shifts-specific fields
  isDraft?: boolean;
}
