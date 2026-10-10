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
import { PropertyPaneAdminDefaultsManager } from './propertyPane/PropertyPaneAdminDefaultsManager';
import { AudienceService } from './services/AudienceService';
import {
  deriveUserCalendarSettings,
  audienceApplies,
  canonicalizeExchangeSourceIdentities,
  cleanAdminSourceOverrides,
  normalizeAdminWebPartSettings,
  serializeAdminWebPartSettings,
  loadAdminWebPartSettings,
  migrateLegacyUserSettings,
  resolveCalendarSettings
} from './services/CalendarSettingsService';
import { SettingsStorageService } from './services/SettingsStorageService';
import * as strings from 'MyCalendarsWebPartStrings';

export interface IMyCalendarsWebPartProps {
  settings?: string;
  adminSettings?: string;
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
  private _adminSavePending = false;
  private _userSavePending = false;
  private _userPreview: IUserCalendarSettings | undefined;
  private _disposed = false;

  public render(): void {
    if (this._disposed) return;
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
        onPreviewSettings: this.handleUserSettingsPreview,
        onCancelSettingsPreview: this.handleCancelUserSettingsPreview,
        isSettingsWritePending: this._userSavePending,
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
      userSettings: this._userPreview || this._userSettings,
      matchedGroupIds: this._matchedGroupIds,
      mailboxDiscoveries: this._mailboxDiscoveries,
      currentUserMailboxId: this.getAadContextId(this.context.pageContext.aadInfo?.userId),
      organizationPrimaryColor: currentTheme.palette?.themePrimary
    });
    // SPFx renders after a theme change. Rendering here would also mount React
    // during framework initialization, before onInit has completed.
  }

  protected onAfterDeserialize(properties: IMyCalendarsWebPartProps): IMyCalendarsWebPartProps {
    return { settings: properties.settings, adminSettings: properties.adminSettings };
  }

  protected onDispose(): void {
    this._disposed = true;
    this._settingsGeneration++;
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
                PropertyPaneAdminDefaultsManager('adminSettings', {
                  label: strings.AdminCalendarManagerLabel,
                  getAdminSettings: () => this._adminSettings,
                  getAdminLoadNotice: () => this._adminLoadNotice,
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
      userSettings: this._userPreview || this._userSettings,
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

  private resolveUserSettings(): void {
    this._resolvedSettings = resolveCalendarSettings({
      adminSettings: this._adminSettings,
      userSettings: this._userPreview || this._userSettings,
      matchedGroupIds: this._matchedGroupIds,
      mailboxDiscoveries: this._mailboxDiscoveries,
      currentUserMailboxId: this.getAadContextId(this.context.pageContext.aadInfo?.userId),
      organizationPrimaryColor: this._themeVariant?.palette?.themePrimary
    });
  }

  private deriveUserSettings(settings: ICalendarSettings): IUserCalendarSettings {
    return deriveUserCalendarSettings({
      nextResolvedSettings: settings,
      adminSettings: this._adminSettings,
      matchedGroupIds: this._matchedGroupIds,
      existingUserSettings: this._userSettings,
      mailboxDiscoveries: this._mailboxDiscoveries,
      currentUserMailboxId: this.getAadContextId(this.context.pageContext.aadInfo?.userId)
    });
  }

  private handleUserSettingsPreview = (settings: ICalendarSettings): void => {
    if (this._userSavePending || this._disposed) return;
    this._userPreview = this.deriveUserSettings(settings);
    this.resolveUserSettings();
    this.render();
  };

  private handleCancelUserSettingsPreview = (): void => {
    if (this._userSavePending) return;
    this._userPreview = undefined;
    this.resolveUserSettings();
    this.render();
  };

  private handleUserSettingsChange = async (settings: ICalendarSettings): Promise<void> => {
    await this.persistUserSettings(this.deriveUserSettings(structuredClone(settings)), false);
  };

  private handleDefaultViewChange = async (defaultView: CalendarViewType): Promise<void> => {
    await this.persistUserSettings({
      ...cleanAdminSourceOverrides(this._userSettings, this._resolvedSettings.applicableAdminSources || [], this._resolvedSettings.unresolvedAdminSourceIds),
      defaultView
    }, false);
  };

  private handleResetUserSettings = async (): Promise<void> => {
    await this.persistUserSettings(structuredClone(defaultUserCalendarSettings), true);
  };

  private async persistUserSettings(snapshot: IUserCalendarSettings, reset: boolean): Promise<void> {
    if (this._userSavePending) throw new Error('A personal settings write is already in progress.');
    this._userSavePending = true;
    this._userPreview = snapshot;
    try {
      this.resolveUserSettings();
      this.render();
      if (!this._storageService) throw new Error('Personal settings storage is unavailable.');
      const success = reset
        ? await this._storageService.deleteUserSettings()
        : await this._storageService.saveUserSettings(snapshot);
      if (!success) throw new Error(reset ? 'Failed to reset personal settings.' : 'Failed to save personal settings.');
      this._userSettings = snapshot;
    } catch (error) {
      console.error(reset ? 'Failed to reset personal settings.' : 'Failed to save personal settings.', error);
      throw error;
    } finally {
      this._userPreview = undefined;
      this._userSavePending = false;
      if (!this._disposed) {
        try { this.resolveUserSettings(); this.render(); }
        catch (error) { console.error('Personal settings storage completed, but runtime settings could not be refreshed.', error); }
      }
    }
  }

  private handleRefreshAdminSources = async (): Promise<void> => {
    await this.rebuildResolvedSettings();
    this.render();
  };

  private handleAdminSettingsSave = async (
    settings: IAdminWebPartSettings,
    commitProperty: (serialized: string | undefined) => void
  ): Promise<void> => {
    if (this._adminSavePending) throw new Error('An administrator save is already in progress.');
    this._adminSavePending = true;
    try {
      const normalized = normalizeAdminWebPartSettings(settings);
      if (!normalized) throw new Error('Invalid administrator defaults.');
      const serialized = serializeAdminWebPartSettings(normalized);
      const previous = this.properties.adminSettings;
      try {
        commitProperty(serialized);
        if (this.properties.adminSettings !== serialized) {
          throw new Error('SPFx did not accept the administrator settings property.');
        }
      } catch (error) {
        if (this.properties.adminSettings !== previous) {
          try {
            commitProperty(previous);
            if (this.properties.adminSettings !== previous) {
              throw new Error('SPFx did not restore the administrator settings property.');
            }
          }
          catch (restoreError) {
            console.error('Could not restore administrator properties after a rejected save.', restoreError);
            this._adminLoadNotice = strings.AdminSettingsRestoreErrorLabel;
          }
        }
        throw error;
      }
      this._adminSettings = normalized;
      this._adminLoadNotice = undefined;
      // The SPFx hand-off succeeded. Runtime failures must not reject this save.
      try {
        await this.rebuildResolvedSettings();
        this.render();
      } catch (error) {
        console.error('Administrator defaults were saved, but resolved settings could not be rebuilt.', error);
        this._adminLoadNotice = strings.AdminRuntimeRebuildErrorLabel;
      }
    } finally {
      this._adminSavePending = false;
    }
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
