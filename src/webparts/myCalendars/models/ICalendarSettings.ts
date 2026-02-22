export type CalendarViewType = 'day' | 'week' | 'month' | 'schedule' | 'search';
export type CalendarSourceType = 'ics' | 'exchange' | 'sharepoint' | 'planner';

export interface ISharePointFieldMapping {
  titleField?: string;        // e.g., "Title"
  startDateField?: string;    // e.g., "EventDate" or "StartDate"
  endDateField?: string;      // e.g., "EndDate"
  descriptionField?: string;  // e.g., "Description"
  locationField?: string;     // e.g., "Location"
  allDayField?: string;       // e.g., "fAllDayEvent"
}

export interface ICalendarSource {
  id: string;
  name: string;
  color: string;
  isEnabled: boolean;
  sourceType: CalendarSourceType;
  // ICS source properties
  url?: string;
  rawContent?: string;
  // Exchange source properties
  exchangeMailbox?: string; // user principal name or email, defaults to current user
  exchangeCalendarId?: string;
  // SharePoint source properties
  sharePointSiteId?: string;
  sharePointListId?: string;
  sharePointFieldMapping?: ISharePointFieldMapping;
  // Planner source properties
  plannerPlanId?: string;
  plannerPlanTitle?: string; // for display purposes
  plannerAssignedToMeOnly?: boolean;
  showCompletedTasks?: boolean; // default true
  // Source display options
  showSourceLogo?: boolean; // default true
}

export interface ICalendarSettings {
  defaultView: CalendarViewType;
  sources: ICalendarSource[];
  showWeekends: boolean;
  startHour: number;
  endHour: number;
  slotDuration: number;
  firstDayOfWeek: number;
  // User-level overrides for start/end hour (overrides webpart settings)
  userStartHour?: number;
  userEndHour?: number;
  // Proxy settings for fetching ICS files (to mitigate CORS)
  useCustomProxy?: boolean;
  customProxyUrl?: string;
  useWhateverOrigin?: boolean;
  // Proxy order preferences (direct is always attempted first)
  proxyPriority1?: 'custom' | 'whateverorigin';
  proxyPriority2?: 'custom' | 'whateverorigin';
  // Theme-based primary color for organization branding (set on first load)
  organizationPrimaryColor?: string;
  // Exchange calendar enabled states (keyed by calendar ID)
  exchangeCalendarStates?: { [calendarId: string]: boolean };
  // Service-level logo display settings
  exchangeShowSourceLogo?: boolean;
  sharePointShowSourceLogo?: boolean;
  plannerShowSourceLogo?: boolean;
}

export const defaultCalendarSettings: ICalendarSettings = {
  defaultView: 'month',
  sources: [],
  showWeekends: true,
  startHour: 8,
  endHour: 18,
  slotDuration: 30,
  firstDayOfWeek: 1,
  useCustomProxy: false,
  customProxyUrl: '',
  useWhateverOrigin: true,
  proxyPriority1: 'custom',
  proxyPriority2: 'whateverorigin',
  organizationPrimaryColor: '#0078d4', // M365 theme primary
  exchangeCalendarStates: {}, // All Exchange calendars enabled by default
  exchangeShowSourceLogo: true,
  sharePointShowSourceLogo: true,
  plannerShowSourceLogo: true
};
