import type { CalendarSourceType } from '../models/ICalendarSettings';
import { getCalendarSourceDefinition } from '../models/CalendarSourceRegistry';
import * as strings from 'MyCalendarsWebPartStrings';

/**
 * Get the Fluent UI icon name for a calendar source type
 */
export function getSourceIconName(sourceType: CalendarSourceType | undefined, explicitIconName?: string): string {
  if (explicitIconName) return explicitIconName;
  return sourceType ? getCalendarSourceDefinition(sourceType)?.iconName || 'Calendar' : 'Calendar';
}

/** User-facing source type name. Never exposes the internal source type key. */
export function getSourceTypeDisplayName(sourceType: CalendarSourceType | undefined, iconName?: string): string {
  if (!sourceType) return strings.CalendarLabel;
  if (sourceType === 'unifiedGroup' && iconName === 'TeamsLogo') return strings.TeamsLabel;
  const definition = getCalendarSourceDefinition(sourceType);
  return definition ? strings[definition.displayNameKey] : strings.CalendarLabel;
}

export function getSourceTypeDescription(sourceType: CalendarSourceType): string {
  const definition = getCalendarSourceDefinition(sourceType);
  return definition ? strings[definition.descriptionKey] : '';
}
