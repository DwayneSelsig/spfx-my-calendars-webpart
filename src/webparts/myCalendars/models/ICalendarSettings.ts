export type CalendarViewType = 'day' | 'week' | 'month' | 'search';
export type CalendarSourceType = 'ics' | 'exchange' | 'sharepoint' | 'planner' | 'teamsShifts' | 'unifiedGroup';
export type CalendarSourceOrigin = 'admin' | 'user';
export type ProxyType = 'custom' | 'whateverorigin';

export const CALENDAR_SETTINGS_SCHEMA_VERSION = 2;

export interface ISharePointFieldMapping {
  titleField?: string;
  startDateField?: string;
  endDateField?: string;
  descriptionField?: string;
  locationField?: string;
  allDayField?: string;
}

export interface ICalendarSourceBase {
  name: string;
  color: string;
  isEnabled: boolean;
  sourceType: CalendarSourceType;
  exchangeMailbox?: string;
  exchangeCalendarId?: string;
  sharePointSiteId?: string;
  sharePointListId?: string;
  sharePointFieldMapping?: ISharePointFieldMapping;
  plannerPlanId?: string;
  plannerPlanTitle?: string;
  plannerAssignedToMeOnly?: boolean;
  showCompletedTasks?: boolean;
  groupId?: string;
  showSourceLogo?: boolean;
  icsUrl?: string;
}

export interface IUserCalendarSource extends ICalendarSourceBase {
  userSourceId: string;
}

export interface IAudienceGroup {
  groupId: string;
  displayName: string;
}

export interface IAdminAssignedSource {
  adminSourceId: string;
  source: ICalendarSourceBase;
  audienceGroups: IAudienceGroup[];
}

export interface IAdminIcsCatalogItem {
  adminIcsId: string;
  displayName: string;
  icsUrl: string;
  audienceGroups: IAudienceGroup[];
}

export interface IAdminSourceOverride {
  removed?: boolean;
  isEnabled?: boolean;
  name?: string;
  color?: string;
}

export interface IAdminWebPartSettings {
  schemaVersion: number;
  defaultView: CalendarViewType;
  showWeekends: boolean;
  startHour: number;
  endHour: number;
  slotDuration: number;
  firstDayOfWeek: number;
  useCustomProxy: boolean;
  customProxyUrl: string;
  useWhateverOrigin: boolean;
  proxyPriority1?: ProxyType;
  proxyPriority2?: ProxyType;
  organizationPrimaryColor?: string;
  exchangeShowSourceLogo: boolean;
  sharePointShowSourceLogo: boolean;
  plannerShowSourceLogo: boolean;
  unifiedGroupShowSourceLogo: boolean;
  teamsShiftsShowSourceLogo: boolean;
  plannerShowAllCalendars: boolean;
  plannerShowAllAssignedToMeOnly: boolean;
  unifiedGroupShowAllCalendars: boolean;
  teamsShiftsShowAllCalendars: boolean;
  assignedSources: IAdminAssignedSource[];
  icsCatalog: IAdminIcsCatalogItem[];
}

export interface IUserCalendarSettings {
  schemaVersion: number;
  defaultView?: CalendarViewType;
  userStartHour?: number;
  userEndHour?: number;
  exchangeCalendarStates: { [calendarId: string]: boolean };
  exchangeShowSourceLogo?: boolean;
  sharePointShowSourceLogo?: boolean;
  plannerShowSourceLogo?: boolean;
  unifiedGroupShowSourceLogo?: boolean;
  teamsShiftsShowSourceLogo?: boolean;
  plannerShowAllCalendars?: boolean;
  plannerShowAllAssignedToMeOnly?: boolean;
  unifiedGroupShowAllCalendars?: boolean;
  teamsShiftsShowAllCalendars?: boolean;
  personalSources: IUserCalendarSource[];
  adminSourceOverridesById: { [adminSourceId: string]: IAdminSourceOverride };
}

export interface ICalendarSource extends ICalendarSourceBase {
  id: string;
  origin: CalendarSourceOrigin;
  adminSourceId?: string;
  userSourceId?: string;
  audienceGroupNames?: string[];
}

export interface ICalendarSettings {
  schemaVersion: number;
  defaultView: CalendarViewType;
  sources: ICalendarSource[];
  availableAdminIcsCatalogItems: IAdminIcsCatalogItem[];
  showWeekends: boolean;
  startHour: number;
  endHour: number;
  slotDuration: number;
  firstDayOfWeek: number;
  userStartHour?: number;
  userEndHour?: number;
  useCustomProxy: boolean;
  customProxyUrl: string;
  useWhateverOrigin: boolean;
  proxyPriority1?: ProxyType;
  proxyPriority2?: ProxyType;
  organizationPrimaryColor?: string;
  exchangeCalendarStates: { [calendarId: string]: boolean };
  exchangeShowSourceLogo: boolean;
  sharePointShowSourceLogo: boolean;
  plannerShowSourceLogo: boolean;
  unifiedGroupShowSourceLogo: boolean;
  teamsShiftsShowSourceLogo: boolean;
  plannerShowAllCalendars: boolean;
  plannerShowAllAssignedToMeOnly: boolean;
  unifiedGroupShowAllCalendars: boolean;
  teamsShiftsShowAllCalendars: boolean;
}

export interface ILegacyCalendarSource extends ICalendarSourceBase {
  id: string;
}

export interface ILegacyCalendarSettings {
  defaultView: CalendarViewType;
  sources: ILegacyCalendarSource[];
  showWeekends: boolean;
  startHour: number;
  endHour: number;
  slotDuration: number;
  firstDayOfWeek: number;
  userStartHour?: number;
  userEndHour?: number;
  useCustomProxy?: boolean;
  customProxyUrl?: string;
  useWhateverOrigin?: boolean;
  proxyPriority1?: ProxyType;
  proxyPriority2?: ProxyType;
  organizationPrimaryColor?: string;
  exchangeCalendarStates?: { [calendarId: string]: boolean };
  exchangeShowSourceLogo?: boolean;
  sharePointShowSourceLogo?: boolean;
  plannerShowSourceLogo?: boolean;
  unifiedGroupShowSourceLogo?: boolean;
  teamsShiftsShowSourceLogo?: boolean;
  plannerShowAllCalendars?: boolean;
  plannerShowAllAssignedToMeOnly?: boolean;
  unifiedGroupShowAllCalendars?: boolean;
  teamsShiftsShowAllCalendars?: boolean;
}

export const defaultAdminWebPartSettings: IAdminWebPartSettings = {
  schemaVersion: CALENDAR_SETTINGS_SCHEMA_VERSION,
  defaultView: 'month',
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
  organizationPrimaryColor: '#0078d4',
  exchangeShowSourceLogo: true,
  sharePointShowSourceLogo: true,
  plannerShowSourceLogo: true,
  unifiedGroupShowSourceLogo: true,
  teamsShiftsShowSourceLogo: true,
  plannerShowAllCalendars: false,
  plannerShowAllAssignedToMeOnly: false,
  unifiedGroupShowAllCalendars: false,
  teamsShiftsShowAllCalendars: false,
  assignedSources: [],
  icsCatalog: []
};

export const defaultUserCalendarSettings: IUserCalendarSettings = {
  schemaVersion: CALENDAR_SETTINGS_SCHEMA_VERSION,
  exchangeCalendarStates: {},
  personalSources: [],
  adminSourceOverridesById: {}
};

export const defaultCalendarSettings: ICalendarSettings = {
  schemaVersion: CALENDAR_SETTINGS_SCHEMA_VERSION,
  defaultView: 'month',
  sources: [],
  availableAdminIcsCatalogItems: [],
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
  organizationPrimaryColor: '#0078d4',
  exchangeCalendarStates: {},
  exchangeShowSourceLogo: true,
  sharePointShowSourceLogo: true,
  plannerShowSourceLogo: true,
  unifiedGroupShowSourceLogo: true,
  teamsShiftsShowSourceLogo: true,
  plannerShowAllCalendars: false,
  plannerShowAllAssignedToMeOnly: false,
  unifiedGroupShowAllCalendars: false,
  teamsShiftsShowAllCalendars: false
};
