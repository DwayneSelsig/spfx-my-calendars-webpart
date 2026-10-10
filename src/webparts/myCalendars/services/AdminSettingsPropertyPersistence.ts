import { serializeAdminWebPartSettings } from './CalendarSettingsService';
import type { IAdminWebPartSettings } from '../models/ICalendarSettings';

export interface IAdminSettingsPropertyBag {
  adminSettings?: string;
  adminSettingsBackup?: string;
}

export type AdminSettingsPropertyChangeNotifier = (serializedSettings: string) => void;
export type PropertyPaneChangeCallback = (targetProperty?: string, newValue?: unknown, isValidEntry?: boolean) => void;

export function notifyAdminSettingsPropertyChanges(
  changeCallback: PropertyPaneChangeCallback,
  targetProperty: string,
  backupTargetProperty: string,
  serializedSettings: string
): void {
  changeCallback(targetProperty, serializedSettings, true);
  changeCallback(backupTargetProperty, serializedSettings, true);
}

export function persistAdminWebPartSettings(
  properties: IAdminSettingsPropertyBag,
  settings: IAdminWebPartSettings,
  notifyPropertyChange: AdminSettingsPropertyChangeNotifier
): string {
  const serialized = serializeAdminWebPartSettings(settings);
  // SPFx must observe each transition before the local property bag is synchronized.
  // Prewriting both values can hide the change from a freshly initialized host snapshot.
  notifyPropertyChange(serialized);
  properties.adminSettings = serialized;
  properties.adminSettingsBackup = serialized;
  return serialized;
}
