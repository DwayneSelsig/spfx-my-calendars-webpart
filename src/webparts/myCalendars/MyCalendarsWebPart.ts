import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import MyCalendars from './components/MyCalendars';
import type { IMyCalendarsProps } from './components/IMyCalendarsProps';
import {
  type IAdminWebPartSettings,
  type ICalendarSettings,
  type CalendarViewType,
  type IUserCalendarSettings,
  defaultAdminWebPartSettings,
  defaultCalendarSettings,
  defaultUserCalendarSettings
} from './models/ICalendarSettings';
import { PropertyPaneAdminCalendarManager } from './propertyPane/PropertyPaneAdminCalendarManager';
import { AudienceService } from './services/AudienceService';
import {
  deriveUserCalendarSettings,
  loadAdminWebPartSettings,
  migrateLegacyUserSettings,
  resolveCalendarSettings
} from './services/CalendarSettingsService';
import { SettingsStorageService } from './services/SettingsStorageService';
import * as strings from 'MyCalendarsWebPartStrings';

export interface IMyCalendarsWebPartProps {
  settings?: string;
  adminSettings?: string;
  adminSettingsBackup?: string;
}

export default class MyCalendarsWebPart extends BaseClientSideWebPart<IMyCalendarsWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _resolvedSettings: ICalendarSettings = { ...defaultCalendarSettings };
  private _adminSettings: IAdminWebPartSettings = { ...defaultAdminWebPartSettings };
  private _userSettings: IUserCalendarSettings = { ...defaultUserCalendarSettings };
  private _matchedGroupIds: Set<string> = new Set<string>();
  private _storageService: SettingsStorageService | null = null;
  private _themeVariant: IReadonlyTheme | undefined;
  private _graphClient: MSGraphClientV3 | undefined;
  private _audienceService: AudienceService | null = null;
  private _adminLoadNotice: string | undefined;

  public render(): void {
    const element: React.ReactElement<IMyCalendarsProps> = React.createElement(
      MyCalendars,
      {
        description: '',
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        settings: this._resolvedSettings,
        onSettingsChange: this.handleUserSettingsChange,
        onDefaultViewChange: this.handleDefaultViewChange,
        onResetSettings: this.handleResetUserSettings,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this._storageService = new SettingsStorageService(this.context.msGraphClientFactory);
    try {
      this._graphClient = await this.context.msGraphClientFactory.getClient('3');
      this._audienceService = new AudienceService(this._graphClient);
    } catch (error) {
      console.error('Failed to initialize Microsoft Graph client:', error);
    }

    const adminLoadResult = loadAdminWebPartSettings({
      current: this.properties.adminSettings,
      backup: this.properties.adminSettingsBackup,
      legacy: this.properties.settings
    });
    this._adminSettings = adminLoadResult.settings;
    this._adminLoadNotice = adminLoadResult.notice;

    if (this._storageService) {
      const persistedUserSettings = await this._storageService.loadUserSettings();
      if (persistedUserSettings) {
        this._userSettings = persistedUserSettings;
      } else {
        const legacyUserSettings = await this._storageService.loadLegacySettings();
        if (legacyUserSettings) {
          this._userSettings = migrateLegacyUserSettings(legacyUserSettings);
          const saved = await this._storageService.saveUserSettings(this._userSettings);
          if (!saved) {
            console.error('Failed to persist migrated user settings.');
          }
        }
      }
    }

    await this.rebuildResolvedSettings();
    this._environmentMessage = await this._getEnvironmentMessage();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    this._themeVariant = currentTheme;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

    this._resolvedSettings = resolveCalendarSettings({
      adminSettings: this._adminSettings,
      userSettings: this._userSettings,
      matchedGroupIds: this._matchedGroupIds,
      organizationPrimaryColor: currentTheme.palette?.themePrimary
    });

    this.render();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure administrator defaults for this web part'
          },
          groups: [
            {
              groupName: 'Administrator Defaults',
              groupFields: [
                PropertyPaneAdminCalendarManager('adminSettingsManager', {
                  label: 'Default calendars and ICS catalog',
                  adminSettings: this._adminSettings,
                  adminLoadNotice: this._adminLoadNotice,
                  context: this.context,
                  onSave: this.handleAdminSettingsSave
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private async rebuildResolvedSettings(): Promise<void> {
    this._matchedGroupIds = await this.resolveMatchedGroupIds();
    this._resolvedSettings = resolveCalendarSettings({
      adminSettings: this._adminSettings,
      userSettings: this._userSettings,
      matchedGroupIds: this._matchedGroupIds,
      organizationPrimaryColor: this._themeVariant?.palette?.themePrimary
    });
  }

  private async resolveMatchedGroupIds(): Promise<Set<string>> {
    if (!this._audienceService) {
      return new Set<string>();
    }

    const groupIds = new Set<string>();
    this._adminSettings.assignedSources.forEach(item => item.audienceGroups.forEach(group => groupIds.add(group.groupId)));
    this._adminSettings.icsCatalog.forEach(item => item.audienceGroups.forEach(group => groupIds.add(group.groupId)));

    return this._audienceService.getMatchingGroupIds(Array.from(groupIds));
  }

  private handleUserSettingsChange = (settings: ICalendarSettings): void => {
    this._userSettings = deriveUserCalendarSettings({
      nextResolvedSettings: settings,
      adminSettings: this._adminSettings,
      matchedGroupIds: this._matchedGroupIds,
      existingUserSettings: this._userSettings
    });

    this._resolvedSettings = resolveCalendarSettings({
      adminSettings: this._adminSettings,
      userSettings: this._userSettings,
      matchedGroupIds: this._matchedGroupIds,
      organizationPrimaryColor: this._themeVariant?.palette?.themePrimary
    });

    if (this._storageService) {
      this._storageService.saveUserSettings(this._userSettings).then(success => {
        if (!success) {
          console.error('Failed to persist user settings.');
        }
      }).catch(error => console.error('Error saving user settings:', error));
    }

    this.render();
  };

  private handleDefaultViewChange = (defaultView: CalendarViewType): void => {
    this._userSettings = {
      ...this._userSettings,
      defaultView
    };

    this._resolvedSettings = resolveCalendarSettings({
      adminSettings: this._adminSettings,
      userSettings: this._userSettings,
      matchedGroupIds: this._matchedGroupIds,
      organizationPrimaryColor: this._themeVariant?.palette?.themePrimary
    });

    if (this._storageService) {
      this._storageService.saveUserSettings(this._userSettings).then(success => {
        if (!success) {
          console.error('Failed to persist the personal default calendar view.');
        }
      }).catch(error => console.error('Error saving the personal default calendar view:', error));
    }

    this.render();
  };

  private handleResetUserSettings = (): void => {
    if (this._storageService) {
      this._storageService.deleteUserSettings().then(success => {
        if (!success) {
          console.error('Failed to delete user settings.');
          return;
        }

        this._userSettings = { ...defaultUserCalendarSettings };
        this._resolvedSettings = resolveCalendarSettings({
          adminSettings: this._adminSettings,
          userSettings: this._userSettings,
          matchedGroupIds: this._matchedGroupIds,
          organizationPrimaryColor: this._themeVariant?.palette?.themePrimary
        });
        this.render();
      }).catch(error => console.error('Error deleting user settings:', error));
    }
  };

  private handleAdminSettingsSave = async (settings: IAdminWebPartSettings): Promise<void> => {
    this._adminSettings = settings;
    const serialized = JSON.stringify(settings);
    this.properties.adminSettings = serialized;
    this.properties.adminSettingsBackup = serialized;
    this._adminLoadNotice = undefined;

    await this.rebuildResolvedSettings();
    this.context.propertyPane.refresh();
    this.render();
  };

  private async _getEnvironmentMessage(): Promise<string> {
    if (this.context.sdks.microsoftTeams) {
      const context = await this.context.sdks.microsoftTeams.teamsJs.app.getContext();
      switch (context.app.host.name) {
        case 'Office':
          return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
        case 'Outlook':
          return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
        case 'Teams':
        case 'TeamsModern':
          return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
        default:
          return strings.UnknownEnvironment;
      }
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }
}
