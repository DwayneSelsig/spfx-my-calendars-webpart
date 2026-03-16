import { ICalendarSettings } from '../models/ICalendarSettings';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IMyCalendarsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  settings: ICalendarSettings;
  onSettingsChange: (settings: ICalendarSettings) => void;
  onResetSettings?: () => void;
  context: WebPartContext;
}
