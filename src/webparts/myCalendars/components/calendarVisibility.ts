import type { ICalendarSettings } from '../models/ICalendarSettings';

export type GroupVisibilityState = 'on' | 'off' | 'mixed';

export function getGroupVisibilityState(values: boolean[]): GroupVisibilityState {
  if (values.length === 0 || values.every(Boolean)) return 'on';
  if (values.every(value => !value)) return 'off';
  return 'mixed';
}

export function getBulkVisibilityTarget(state: GroupVisibilityState): boolean {
  return state !== 'on';
}

export function setOutlookVisibility(settings: ICalendarSettings, discoveredCalendarIds: string[], visible: boolean): ICalendarSettings {
  const exchangeCalendarStates = { ...settings.exchangeCalendarStates };
  discoveredCalendarIds.forEach(calendarId => { exchangeCalendarStates[calendarId] = visible; });
  return {
    ...settings,
    exchangeCalendarStates,
    sources: settings.sources.map(source => source.sourceType === 'exchange' ? { ...source, isEnabled: visible } : source)
  };
}

export function setSharePointVisibility(settings: ICalendarSettings, visible: boolean): ICalendarSettings {
  return {
    ...settings,
    sources: settings.sources.map(source => source.sourceType === 'sharepoint' ? { ...source, isEnabled: visible } : source)
  };
}
