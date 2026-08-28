import type { CalendarSourceType } from './ICalendarSettings';

export interface ICalendarSourceDefinition {
  type: CalendarSourceType;
  displayName: string;
  description: string;
  iconName: string;
  defaultColor: string;
  userSelectable: boolean;
  adminSelectable: boolean;
  adminCatalogOnly?: boolean;
}

export const calendarSourceRegistry: ICalendarSourceDefinition[] = [
  {
    type: 'sharepoint',
    displayName: 'SharePoint Calendar',
    description: 'Select a calendar from a SharePoint site',
    iconName: 'SharepointLogo',
    defaultColor: '#038186',
    userSelectable: true,
    adminSelectable: true
  },
  {
    type: 'exchange',
    displayName: 'Outlook Calendar',
    description: 'Manage Outlook calendars (opens in new window)',
    iconName: 'OutlookLogo',
    defaultColor: '#0078d4',
    userSelectable: true,
    adminSelectable: true
  },
  {
    type: 'planner',
    displayName: 'Microsoft Planner',
    description: 'Voeg taken toe vanuit een Planner plan',
    iconName: 'PlannerLogo',
    defaultColor: '#107c41',
    userSelectable: true,
    adminSelectable: true
  },
  {
    type: 'unifiedGroup',
    displayName: 'Microsoft 365 Group / Teams',
    description: 'Add a calendar from a Microsoft 365 group or Team',
    iconName: 'Group',
    defaultColor: '#5b5fc7',
    userSelectable: true,
    adminSelectable: true
  },
  {
    type: 'teamsShifts',
    displayName: 'Teams Shifts',
    description: 'Toon diensten uit Teams shifts',
    iconName: 'Clock',
    defaultColor: '#4a4fbe',
    userSelectable: true,
    adminSelectable: true
  },
  {
    type: 'ics',
    displayName: 'Internet Calendar',
    description: 'Add calendar from URL or paste ICS content',
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
