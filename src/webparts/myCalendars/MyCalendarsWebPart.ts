import { ExchangeCalendarService, getExchangeDiscoveryErrorMessage } from './services/ExchangeCalendarService';
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
  type IExchangeMailboxDiscovery,
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
  audienceApplies,
  canonicalizeExchangeSourceIdentities,
  cleanAdminSourceOverrides,
  normalizeAdminWebPartSettings,
  loadAdminWebPartSettings,
  migrateLegacyUserSettings,
  resolveCalendarSettings
} from './services/CalendarSettingsService';
import { SettingsStorageService } from './services/SettingsStorageService';
import {
  persistAdminWebPartSettings,
  type AdminSettingsPropertyChangeNotifier
} from './services/AdminSettingsPropertyPersistence';
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
  private _mailboxDiscoveries: IExchangeMailboxDiscovery[] = [];
  private _settingsGeneration = 0;
  private _adminSaveGeneration = 0;

  public render(): void {
    const element: React.ReactElement<IMyCalendarsProps> = React.createElement(
      MyCalendars,
      {
        description: '',
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        locale: this.context.pageContext.cultureInfo.currentCultureName,
        settings: this._resolvedSettings,
        onSettingsChange: this.handleUserSettingsChange,
        onDefaultViewChange: this.handleDefaultViewChange,
        onResetSettings: this.handleResetUserSettings,
        onRefreshAdminSources: this.handleRefreshAdminSources,
        context: this.context,
        tenantId: this.getAadContextId(this.context.pageContext.aadInfo?.tenantId),
        userId: this.getAadContextId(this.context.pageContext.aadInfo?.userId),
        webPartInstanceId: this.instanceId
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
      mailboxDiscoveries: this._mailboxDiscoveries,
      currentUserMailboxId: this.getAadContextId(this.context.pageContext.aadInfo?.userId),
      organizationPrimaryColor: currentTheme.palette?.themePrimary
    });
    // SPFx renders after a theme change. Rendering here would also mount React
    // during framework initialization, before onInit has completed.
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
            description: strings.PropertyPaneAdminDescription
          },
          groups: [
            {
              groupName: strings.AdminDefaultsGroupName,
              groupFields: [
                PropertyPaneAdminCalendarManager('adminSettings', {
                  label: strings.AdminCalendarManagerLabel,
                  adminSettings: this._adminSettings,
                  backupTargetProperty: 'adminSettingsBackup',
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
    const generation = ++this._settingsGeneration;
    const matched = await this.resolveMatchedGroupIds();
    const rules = (this._adminSettings.exchangeMailboxAssignments || []).filter(rule => audienceApplies(rule.audienceGroups, matched));
    const service = this._graphClient ? new ExchangeCalendarService(this._graphClient) : undefined;
    const aliases = Array.from(new Set(this._adminSettings.assignedSources.filter(item => audienceApplies(item.audienceGroups, matched) && item.source.sourceType === 'exchange' && item.source.exchangeMailbox).map(item => item.source.exchangeMailbox as string)));
    const resolvedAliases = await Promise.all(aliases.map(async alias => {
      try { return { alias, id: (await service?.resolveMailbox(alias))?.id }; }
      catch { return { alias, id: undefined }; }
    }));
    if (generation !== this._settingsGeneration) return;
    this._adminSettings = canonicalizeExchangeSourceIdentities(this._adminSettings, new Map(resolvedAliases.filter(item => item.id).map(item => [item.alias.toLowerCase(), item.id as string])));
    const mailboxIds = Array.from(new Set(rules.map(rule => rule.mailboxId)));
    const discoveries = await Promise.all(mailboxIds.map(async mailboxId => {
      try {
        if (!service) throw new Error('GraphClient not initialized');
        const calendars = await service.getCalendars(mailboxId);
        return { mailboxId, sources: calendars.map(calendar => ({ sourceType: 'exchange' as const, exchangeMailbox: mailboxId,
          exchangeCalendarId: calendar.id, name: calendar.name, color: calendar.hexColor, isEnabled: true })) };
      } catch (error) { return { mailboxId, sources: [], error: getExchangeDiscoveryErrorMessage(error) }; }
    }));
    if (generation !== this._settingsGeneration) return;
    this._matchedGroupIds = matched;
    this._mailboxDiscoveries = discoveries;
    this._resolvedSettings = resolveCalendarSettings({
      adminSettings: this._adminSettings,
      userSettings: this._userSettings,
      matchedGroupIds: this._matchedGroupIds,
      mailboxDiscoveries: this._mailboxDiscoveries,
      currentUserMailboxId: this.getAadContextId(this.context.pageContext.aadInfo?.userId),
      organizationPrimaryColor: this._themeVariant?.palette?.themePrimary
    });
    this._userSettings = cleanAdminSourceOverrides(this._userSettings, this._resolvedSettings.applicableAdminSources || [], this._resolvedSettings.unresolvedAdminSourceIds);
  }

  private getAadContextId(value: unknown): string | undefined {
    try {
      if (typeof value === 'string') return value.trim() || undefined;
      if (value && typeof value === 'object' && typeof (value as { toString?: unknown }).toString === 'function') {
        const stringValue = (value as { toString: () => string }).toString().trim();
        return stringValue && stringValue !== '[object Object]' ? stringValue : undefined;
      }
    } catch (error) {
      console.warn('Could not read an Azure Active Directory context identifier; appointment caching is disabled.', error);
    }
    return undefined;
  }

  private async resolveMatchedGroupIds(): Promise<Set<string>> {
    if (!this._audienceService) {
      return new Set<string>();
    }

    const groupIds = new Set<string>();
    this._adminSettings.assignedSources.forEach(item => item.audienceGroups.forEach(group => groupIds.add(group.groupId)));
    (this._adminSettings.exchangeMailboxAssignments || []).forEach(item => item.audienceGroups.forEach(group => groupIds.add(group.groupId)));
    this._adminSettings.icsCatalog.forEach(item => item.audienceGroups.forEach(group => groupIds.add(group.groupId)));

    return this._audienceService.getMatchingGroupIds(Array.from(groupIds));
  }

  private handleUserSettingsChange = (settings: ICalendarSettings): void => {
    this._userSettings = deriveUserCalendarSettings({
      nextResolvedSettings: settings,
      adminSettings: this._adminSettings,
      matchedGroupIds: this._matchedGroupIds,
      existingUserSettings: this._userSettings,
      mailboxDiscoveries: this._mailboxDiscoveries,
      currentUserMailboxId: this.getAadContextId(this.context.pageContext.aadInfo?.userId)
    });

    this._resolvedSettings = resolveCalendarSettings({
      adminSettings: this._adminSettings,
      userSettings: this._userSettings,
      matchedGroupIds: this._matchedGroupIds,
      mailboxDiscoveries: this._mailboxDiscoveries,
      currentUserMailboxId: this.getAadContextId(this.context.pageContext.aadInfo?.userId),
      organizationPrimaryColor: this._themeVariant?.palette?.themePrimary
    });

    this._userSettings = cleanAdminSourceOverrides(this._userSettings, this._resolvedSettings.applicableAdminSources || [], this._resolvedSettings.unresolvedAdminSourceIds);

    if (this._storageService) {
      this._storageService.saveUserSettings(this._userSettings).then(success => {
        if (!success) {
          console.error('Failed to persist user settings.');
        }
      }).catch(error => console.error('Error saving user settings:', error));
    }

    this.render();
  };

  private handleRefreshAdminSources = async (): Promise<void> => {
    await this.rebuildResolvedSettings();
    this.render();
  };

  private handleDefaultViewChange = (defaultView: CalendarViewType): void => {
    this._userSettings = {
      ...cleanAdminSourceOverrides(this._userSettings, this._resolvedSettings.applicableAdminSources || [], this._resolvedSettings.unresolvedAdminSourceIds),
      defaultView
    };

    this._resolvedSettings = resolveCalendarSettings({
      adminSettings: this._adminSettings,
      userSettings: this._userSettings,
      matchedGroupIds: this._matchedGroupIds,
      mailboxDiscoveries: this._mailboxDiscoveries,
      currentUserMailboxId: this.getAadContextId(this.context.pageContext.aadInfo?.userId),
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
          mailboxDiscoveries: this._mailboxDiscoveries,
          currentUserMailboxId: this.getAadContextId(this.context.pageContext.aadInfo?.userId),
          organizationPrimaryColor: this._themeVariant?.palette?.themePrimary
        });
        this.render();
      }).catch(error => console.error('Error deleting user settings:', error));
    }
  };

  private handleAdminSettingsSave = async (
    settings: IAdminWebPartSettings,
    notifyPropertyChange: AdminSettingsPropertyChangeNotifier
  ): Promise<void> => {
    const normalized = normalizeAdminWebPartSettings(settings);
    if (!normalized) throw new Error('Invalid administrator settings.');
    const generation = ++this._adminSaveGeneration;
    persistAdminWebPartSettings(this.properties, normalized, notifyPropertyChange);
    this._adminSettings = normalized;
    this._adminLoadNotice = undefined;

    await this.rebuildResolvedSettings();
    if (generation !== this._adminSaveGeneration) return;
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
