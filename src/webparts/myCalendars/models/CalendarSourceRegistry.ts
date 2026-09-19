import type { CalendarSourceType } from './ICalendarSettings';

export interface ICalendarSourceDefinition {
  type: CalendarSourceType;
  displayNameKey: keyof IMyCalendarsWebPartStrings;
  descriptionKey: keyof IMyCalendarsWebPartStrings;
  iconName: string;
  defaultColor: string;
  userSelectable: boolean;
  adminSelectable: boolean;
  adminCatalogOnly?: boolean;
}

export const calendarSourceRegistry: ICalendarSourceDefinition[] = [
  {
    type: 'sharepoint',
    displayNameKey: 'SharePointCalendarLabel',
    descriptionKey: 'AddCalendarSharePointDescription',
    iconName: 'SharepointLogo',
    defaultColor: '#038186',
    userSelectable: true,
    adminSelectable: true
  },
  {
    type: 'exchange',
    displayNameKey: 'OutlookCalendarLabel',
    descriptionKey: 'OutlookCalendarDescription',
    iconName: 'OutlookLogo',
    defaultColor: '#0078d4',
    userSelectable: true,
    adminSelectable: true
  },
  {
    type: 'planner',
    displayNameKey: 'MicrosoftPlannerLabel',
    descriptionKey: 'AddFromPlannerLabel',
    iconName: 'PlannerLogo',
    defaultColor: '#107c41',
    userSelectable: true,
    adminSelectable: true
  },
  {
    type: 'unifiedGroup',
    displayNameKey: 'Microsoft365GroupTeamsLabel',
    descriptionKey: 'AddGroupAndTeamCalendarDescription',
    iconName: 'Group',
    defaultColor: '#5b5fc7',
    userSelectable: true,
    adminSelectable: true
  },
  {
    type: 'teamsShifts',
    displayNameKey: 'TeamsShiftsLabel',
    descriptionKey: 'TeamsShiftsDescription',
    iconName: 'Clock',
    defaultColor: '#4a4fbe',
    userSelectable: true,
    adminSelectable: true
  },
  {
    type: 'ics',
    displayNameKey: 'InternetCalendarLabel',
    descriptionKey: 'AddIcsCalendarDescription',
    iconName: 'World',
    defaultColor: '#605e5c',
    userSelectable: true,
    adminSelectable: true,
    adminCatalogOnly: true
  }
];

export function getCalendarSourceDefinition(type: CalendarSourceType): ICalendarSourceDefinition | undefined {
  return calendarSourceRegistry.find(definition => definition.type === type);
}
