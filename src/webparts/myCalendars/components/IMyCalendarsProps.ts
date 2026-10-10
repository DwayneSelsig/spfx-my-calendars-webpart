import { CalendarViewType, ICalendarSettings } from '../models/ICalendarSettings';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IMyCalendarsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  locale?: string;
  settings: ICalendarSettings;
  onSettingsChange: (settings: ICalendarSettings) => Promise<void>;
  onPreviewSettings: (settings: ICalendarSettings) => void;
  onCancelSettingsPreview: () => void;
  isSettingsWritePending: boolean;
  onDefaultViewChange: (view: CalendarViewType) => Promise<void>;
  onResetSettings: () => Promise<void>;
  onRefreshAdminSources?: () => Promise<void>;
  context: WebPartContext;
  tenantId?: string;
  userId?: string;
  webPartInstanceId: string;
}
