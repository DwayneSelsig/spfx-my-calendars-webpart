import { CalendarSourceType } from '../models/ICalendarSettings';

/**
 * Get the Fluent UI icon name for a calendar source type
 */
export function getSourceIcon(sourceType: CalendarSourceType | undefined): string {
  switch (sourceType) {
    case 'exchange':
      return 'OutlookLogo';
    case 'sharepoint':
      return 'SharepointLogo';
    case 'planner':
      return 'PlannerLogo';
    case 'teamsShifts':
      return 'Clock';
    case 'ics':
      return 'Calendar';
    default:
      return 'Calendar';
  }
}
