import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneCheckbox,
  PropertyPaneTextField,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import PnPTelemetry from '@pnp/telemetry-js';

import * as strings from 'MyCalendarsWebPartStrings';
import MyCalendars from './components/MyCalendars';
import { IMyCalendarsProps } from './components/IMyCalendarsProps';
import { ICalendarSettings, defaultCalendarSettings, CalendarViewType } from './models/ICalendarSettings';
import { SettingsStorageService } from './services/SettingsStorageService';

export interface IMyCalendarsWebPartProps {
  settings: string;
  disablePnpTelemetry?: boolean;
}

export default class MyCalendarsWebPart extends BaseClientSideWebPart<IMyCalendarsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _currentSettings: ICalendarSettings = { ...defaultCalendarSettings };
  private _storageService: SettingsStorageService | null = null;
  private _themeVariant: IReadonlyTheme | undefined;

  public render(): void {
    const settings = this.getSettings();
    
    const element: React.ReactElement<IMyCalendarsProps> = React.createElement(
      MyCalendars,
      {
        description: '',
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        settings: settings,
        onSettingsChange: this.handleSettingsChange,
        onResetSettings: this.handleResetSettings,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private getSettings(): ICalendarSettings {
    // Always merge defaults with current persisted settings to ensure
    // newly added fields (like proxy options) have values, and changes
    // from the property pane are reflected immediately.
    let persisted: Partial<ICalendarSettings> = {};
    if (this.properties.settings) {
      try {
        persisted = JSON.parse(this.properties.settings);
      } catch {
        persisted = {};
      }
    }

    // Merge order: defaults < previous current < persisted
    const merged: ICalendarSettings = {
      ...defaultCalendarSettings,
      ...(this._currentSettings || {}),
      ...(persisted as ICalendarSettings)
    };

    this._currentSettings = merged;
    return merged;
  }

  private handleSettingsChange = (settings: ICalendarSettings): void => {
    this._currentSettings = settings;
    // Save to both web part properties (for backward compatibility) and App Folder (persistent)
    this.properties.settings = JSON.stringify(settings);
    
    // Also persist to App Folder
    if (this._storageService) {
      this._storageService.saveSettings(settings).then(success => {
        if (success) {
          console.log('Settings persisted to App Folder');
        } else {
          console.error('Failed to persist settings to App Folder');
        }
      }).catch(err => console.error('Error saving settings:', err));
    }
    
    this.render();
  };

  private handleResetSettings = (): void => {
    // Delete user settings from App Folder
    if (this._storageService) {
      this._storageService.deleteSettings().then(success => {
        if (success) {
          console.log('User settings removed from App Folder, reverting to defaults');
          // Reset to default admin settings
          this._currentSettings = { ...defaultCalendarSettings };
          // Clear web part property storage as well for consistency
          this.properties.settings = JSON.stringify(defaultCalendarSettings);
          this.render();
        } else {
          console.error('Failed to remove user settings from App Folder');
        }
      }).catch(err => console.error('Error deleting settings:', err));
    }
  };

  protected onInit(): Promise<void> {
    this.applyTelemetryPreference();

    // Initialize storage service
    this._storageService = new SettingsStorageService(
      this.context.msGraphClientFactory
    );

    // Load persisted settings from App Folder
    return this._storageService.loadSettings().then(appFolderSettings => {
      if (appFolderSettings) {
        this._currentSettings = appFolderSettings;
      }
      return this._getEnvironmentMessage().then(message => {
        this._environmentMessage = message;
      });
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  private applyTelemetryPreference(): void {
    if (this.properties.disablePnpTelemetry === true) {
      const telemetry = PnPTelemetry.getInstance();
      telemetry.optOut();
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    this._themeVariant = currentTheme;
    const {
      semanticColors,
      palette
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

    // Update organization primary color from theme
    if (palette && palette.themePrimary) {
      this._currentSettings.organizationPrimaryColor = palette.themePrimary;
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const settings = this.getSettings();
    
    return {
      pages: [
        {
          header: {
            description: 'Configure your calendar sources and settings'
          },
          groups: [
            {
              groupName: 'View Settings',
              groupFields: [
                PropertyPaneDropdown('defaultView', {
                  label: 'Default View',
                  options: [
                    { key: 'day', text: 'Day' },
                    { key: 'week', text: 'Week' },
                    { key: 'month', text: 'Month' },
                    { key: 'schedule', text: 'Schedule' }
                  ],
                  selectedKey: settings.defaultView
                }),
                PropertyPaneToggle('showWeekends', {
                  label: 'Show Weekends',
                  checked: settings.showWeekends,
                  onText: 'Yes',
                  offText: 'No'
                }),
                PropertyPaneSlider('startHour', {
                  label: 'Start Hour',
                  min: 0,
                  max: 23,
                  value: settings.startHour,
                  showValue: true
                }),
                PropertyPaneSlider('endHour', {
                  label: 'End Hour',
                  min: 0,
                  max: 23,
                  value: settings.endHour,
                  showValue: true
                }),
                PropertyPaneSlider('slotDuration', {
                  label: 'Slot Duration (minutes)',
                  min: 15,
                  max: 60,
                  step: 15,
                  value: settings.slotDuration,
                  showValue: true
                }),
                PropertyPaneSlider('firstDayOfWeek', {
                  label: 'First Day of Week (0=Sunday, 1=Monday)',
                  min: 0,
                  max: 6,
                  value: settings.firstDayOfWeek,
                  showValue: true
                })
              ]
            },
            {
              groupName: 'PnP Controls',
              groupFields: [
                PropertyPaneToggle('disablePnpTelemetry', {
                  label: 'Schakel PnP Controls telemetry uit',
                  checked: this.properties.disablePnpTelemetry === true,
                  onText: 'Uit',
                  offText: 'Aan'
                }),
                PropertyPaneLabel('telemetryInfo', {
                  text: 'Standaard staat telemetry aan. Bij uitschakelen wordt PnP telemetry opt-out toegepast voor deze webpart.'
                })
              ]
            },
            {
              groupName: 'ICS Proxy',
              groupFields: [
                PropertyPaneLabel('icsProxyInfo', {
                  text: 'Sommige ICS-bronnen vereisen een proxy wegens CORS.'
                }),
                PropertyPaneCheckbox('useCustomProxy', {
                  text: 'Gebruik eigen proxy',
                  checked: !!settings.useCustomProxy
                }),
                PropertyPaneTextField('customProxyUrl', {
                  label: 'Eigen proxy URL',
                  placeholder: 'https://proxy.example.com/https://doel-host/...',
                  description: 'Wordt gebruikt als "Eigen proxy" is ingeschakeld',
                  value: settings.customProxyUrl || '',
                  disabled: !settings.useCustomProxy
                }),
                PropertyPaneCheckbox('useWhateverOrigin', {
                  text: 'Gebruik https://www.whateverorigin.org/',
                  checked: !!settings.useWhateverOrigin
                }),
                PropertyPaneLabel('corsTip', {
                  text: 'Tip: zie ook https://github.com/Zibri/cloudflare-cors-anywhere voor een eigen proxy.'
                }),
                PropertyPaneLabel('proxyOrderHeader', { text: 'Proxy volgorde (na directe poging):' }),
                PropertyPaneDropdown('proxyPriority1', {
                  label: 'Eerste proxy',
                  options: [
                    { key: 'custom', text: 'Eigen proxy' },
                    { key: 'whateverorigin', text: 'whateverorigin.org' }
                  ],
                  selectedKey: settings.proxyPriority1
                }),
                PropertyPaneDropdown('proxyPriority2', {
                  label: 'Tweede proxy',
                  options: [
                    { key: 'custom', text: 'Eigen proxy' },
                    { key: 'whateverorigin', text: 'whateverorigin.org' }
                  ],
                  selectedKey: settings.proxyPriority2
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    const settings = this.getSettings();
    
    switch (propertyPath) {
      case 'defaultView':
        settings.defaultView = newValue as CalendarViewType;
        break;
      case 'showWeekends':
        settings.showWeekends = newValue as boolean;
        break;
      case 'startHour':
        settings.startHour = newValue as number;
        break;
      case 'endHour':
        settings.endHour = newValue as number;
        break;
      case 'slotDuration':
        settings.slotDuration = newValue as number;
        break;
      case 'firstDayOfWeek':
        settings.firstDayOfWeek = newValue as number;
        break;
      case 'useCustomProxy': {
        const v = !!newValue;
        settings.useCustomProxy = v;
        if (v) {
          settings.useWhateverOrigin = false;
        }
        break;
      }
      case 'customProxyUrl':
        settings.customProxyUrl = (newValue as string) || '';
        break;
      case 'useWhateverOrigin': {
        const v = !!newValue;
        settings.useWhateverOrigin = v;
        if (v) {
          settings.useCustomProxy = false;
        }
        break;
      }
      case 'proxyPriority1':
        settings.proxyPriority1 = newValue as 'custom' | 'whateverorigin';
        break;
      case 'proxyPriority2':
        settings.proxyPriority2 = newValue as 'custom' | 'whateverorigin';
        break;
      case 'disablePnpTelemetry':
        this.properties.disablePnpTelemetry = newValue === true;
        this.applyTelemetryPreference();
        break;
    }
    
    this.handleSettingsChange(settings);
  }
}
