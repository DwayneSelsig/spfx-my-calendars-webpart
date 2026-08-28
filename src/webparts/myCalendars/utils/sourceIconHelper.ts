import type { CalendarSourceType } from '../models/ICalendarSettings';
import { getCalendarSourceDefinition } from '../models/CalendarSourceRegistry';

/**
 * Get the Fluent UI icon name for a calendar source type
 */
export function getSourceIconName(sourceType: CalendarSourceType | undefined, explicitIconName?: string): string {
  if (explicitIconName) return explicitIconName;
  return sourceType ? getCalendarSourceDefinition(sourceType)?.iconName || 'Calendar' : 'Calendar';
}

/** User-facing source type name. Never exposes the internal source type key. */
export function getSourceTypeDisplayName(sourceType: CalendarSourceType | undefined, iconName?: string): string {
  if (!sourceType) return 'Calendar';
  if (sourceType === 'unifiedGroup' && iconName === 'TeamsLogo') return 'Microsoft Teams';
  return getCalendarSourceDefinition(sourceType)?.displayName || 'Calendar';
}
