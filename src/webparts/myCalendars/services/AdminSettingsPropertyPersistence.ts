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
  const serialized = JSON.stringify(settings);
  properties.adminSettings = serialized;
  properties.adminSettingsBackup = serialized;
  notifyPropertyChange(serialized);
  return serialized;
}
