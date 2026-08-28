import {
  CALENDAR_SETTINGS_SCHEMA_VERSION,
  type IAdminAssignedSource,
  type IAdminIcsCatalogItem,
  type IAdminSourceOverride,
  type IAdminWebPartSettings,
  type IAudienceGroup,
  type ICalendarSettings,
  type ICalendarSource,
  type ICalendarSourceBase,
  type ILegacyCalendarSettings,
  type ISharePointFieldMapping,
  type IUserCalendarSettings,
  type IUserCalendarSource,
  defaultAdminWebPartSettings,
  defaultCalendarSettings,
  defaultUserCalendarSettings
} from '../models/ICalendarSettings';

export type AdminSettingsLoadSource = 'current' | 'backup' | 'defaults' | 'legacy';

export interface IAdminSettingsLoadResult {
  settings: IAdminWebPartSettings;
  source: AdminSettingsLoadSource;
  notice?: string;
}

interface IParsedSourceShape extends Partial<ICalendarSourceBase> {
  audienceGroupNames?: string[];
  adminSourceId?: string;
  userSourceId?: string;
}

const SOURCE_TYPES = new Set(['ics', 'exchange', 'sharepoint', 'planner', 'teamsShifts', 'unifiedGroup']);

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isValidSourceType(value: unknown): value is ICalendarSourceBase['sourceType'] {
  return typeof value === 'string' && SOURCE_TYPES.has(value);
}

function normalizeFieldMapping(value: unknown): ISharePointFieldMapping | undefined {
  if (!isRecord(value)) {
    return undefined;
  }

  const mapping: ISharePointFieldMapping = {};
  if (typeof value.titleField === 'string') mapping.titleField = value.titleField;
  if (typeof value.startDateField === 'string') mapping.startDateField = value.startDateField;
  if (typeof value.endDateField === 'string') mapping.endDateField = value.endDateField;
  if (typeof value.descriptionField === 'string') mapping.descriptionField = value.descriptionField;
  if (typeof value.locationField === 'string') mapping.locationField = value.locationField;
  if (typeof value.allDayField === 'string') mapping.allDayField = value.allDayField;

  return Object.keys(mapping).length > 0 ? mapping : undefined;
}

function normalizeCalendarSourceBase(value: unknown): ICalendarSourceBase | undefined {
  if (!isRecord(value) || !isValidSourceType(value.sourceType) || typeof value.name !== 'string' || typeof value.color !== 'string') {
    return undefined;
  }

  const source: ICalendarSourceBase = {
    sourceType: value.sourceType,
    name: value.name,
    color: value.color,
    isEnabled: value.isEnabled !== false
  };

  if (typeof value.exchangeMailbox === 'string' && value.exchangeMailbox.trim()) {
    source.exchangeMailbox = value.exchangeMailbox.trim();
  }
  if (typeof value.exchangeCalendarId === 'string' && value.exchangeCalendarId.trim()) {
    source.exchangeCalendarId = value.exchangeCalendarId.trim();
  }
  if (typeof value.sharePointSiteId === 'string' && value.sharePointSiteId.trim()) {
    source.sharePointSiteId = value.sharePointSiteId.trim();
  }
  if (typeof value.sharePointListId === 'string' && value.sharePointListId.trim()) {
    source.sharePointListId = value.sharePointListId.trim();
  }
  const fieldMapping = normalizeFieldMapping(value.sharePointFieldMapping);
  if (fieldMapping) {
    source.sharePointFieldMapping = fieldMapping;
  }
  if (typeof value.plannerPlanId === 'string' && value.plannerPlanId.trim()) {
    source.plannerPlanId = value.plannerPlanId.trim();
  }
  if (typeof value.plannerPlanTitle === 'string' && value.plannerPlanTitle.trim()) {
    source.plannerPlanTitle = value.plannerPlanTitle.trim();
  }
  if (typeof value.plannerAssignedToMeOnly === 'boolean') {
    source.plannerAssignedToMeOnly = value.plannerAssignedToMeOnly;
  }
  if (typeof value.showCompletedTasks === 'boolean') {
    source.showCompletedTasks = value.showCompletedTasks;
  }
  if (typeof value.groupId === 'string' && value.groupId.trim()) {
    source.groupId = value.groupId.trim();
  }
  if (typeof value.showSourceLogo === 'boolean') {
    source.showSourceLogo = value.showSourceLogo;
  }
  if (typeof value.icsUrl === 'string' && value.icsUrl.trim()) {
    source.icsUrl = value.icsUrl.trim();
  }

  return source;
}

function normalizeAudienceGroups(value: unknown): IAudienceGroup[] {
  if (!Array.isArray(value)) {
    return [];
  }

  return value
    .map(item => {
      if (!isRecord(item) || typeof item.groupId !== 'string' || typeof item.displayName !== 'string') {
        return undefined;
      }

      return {
        groupId: item.groupId.trim(),
        displayName: item.displayName.trim()
      };
    })
    .filter((item): item is IAudienceGroup => !!item && !!item.groupId && !!item.displayName);
}

function normalizeOverride(value: unknown): IAdminSourceOverride | undefined {
  if (!isRecord(value)) {
    return undefined;
  }

  const override: IAdminSourceOverride = {};
  if (typeof value.removed === 'boolean') override.removed = value.removed;
  if (typeof value.isEnabled === 'boolean') override.isEnabled = value.isEnabled;
  if (typeof value.name === 'string' && value.name.trim()) override.name = value.name.trim();
  if (typeof value.color === 'string' && value.color.trim()) override.color = value.color.trim();

  return Object.keys(override).length > 0 ? override : undefined;
}

export function generateStableId(prefix: string = 'source'): string {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 11)}`;
}

export function stripRuntimeSource(source: ICalendarSource | IUserCalendarSource | ICalendarSourceBase): ICalendarSourceBase {
  const candidate = source as IParsedSourceShape;
  const normalized = normalizeCalendarSourceBase(candidate);
  if (!normalized) {
    throw new Error('Unable to normalize calendar source.');
  }
  return normalized;
}

export function createUserCalendarSource(source: ICalendarSourceBase, userSourceId?: string): IUserCalendarSource {
  return {
    userSourceId: userSourceId || generateStableId('userSource'),
    ...stripRuntimeSource(source)
  };
}

export function createAdminAssignedSource(source: ICalendarSourceBase, audienceGroups: IAudienceGroup[], adminSourceId?: string): IAdminAssignedSource {
  return {
    adminSourceId: adminSourceId || generateStableId('adminSource'),
    source: stripRuntimeSource(source),
    audienceGroups: audienceGroups.map(group => ({
      groupId: group.groupId,
      displayName: group.displayName
    }))
  };
}

function getSourceIdentityKey(source: ICalendarSourceBase): string {
  switch (source.sourceType) {
    case 'sharepoint':
      return `${source.sourceType}|${source.sharePointSiteId || ''}|${source.sharePointListId || ''}`;
    case 'exchange':
      return `${source.sourceType}|${source.exchangeMailbox || 'me'}|${source.exchangeCalendarId || 'calendar'}`;
    case 'planner':
      return `${source.sourceType}|${source.plannerPlanId || ''}`;
    case 'unifiedGroup':
      return `${source.sourceType}|${source.groupId || ''}`;
    case 'teamsShifts':
      return `${source.sourceType}|teamsShifts`;
    case 'ics':
      return `${source.sourceType}|${source.icsUrl || ''}`;
    default:
      return `${source.sourceType}|${source.name}`;
  }
}

function dedupeAudienceGroups(groups: IAudienceGroup[]): IAudienceGroup[] {
  const byId = new Map<string, IAudienceGroup>();
  groups.forEach(group => {
    if (group.groupId && group.displayName) {
      byId.set(group.groupId, group);
    }
  });
  return Array.from(byId.values());
}

export function normalizeAdminWebPartSettings(value: unknown): IAdminWebPartSettings | undefined {
  if (!isRecord(value)) {
    return undefined;
  }

  const assignedSources = Array.isArray(value.assignedSources)
    ? value.assignedSources
      .map(item => {
        if (!isRecord(item) || typeof item.adminSourceId !== 'string') {
          return undefined;
        }

        const source = normalizeCalendarSourceBase(item.source);
        const audienceGroups = dedupeAudienceGroups(normalizeAudienceGroups(item.audienceGroups));
        if (!source || audienceGroups.length === 0) {
          return undefined;
        }

        return {
          adminSourceId: item.adminSourceId.trim(),
          source,
          audienceGroups
        } as IAdminAssignedSource;
      })
      .filter((item): item is IAdminAssignedSource => !!item && !!item.adminSourceId)
    : [];

  const sourceIdentityKeys = new Set<string>();
  for (const assignedSource of assignedSources) {
    const identityKey = getSourceIdentityKey(assignedSource.source);
    if (sourceIdentityKeys.has(identityKey)) {
      return undefined;
    }
    sourceIdentityKeys.add(identityKey);
  }

  const icsCatalog = Array.isArray(value.icsCatalog)
    ? value.icsCatalog
      .map(item => {
        if (!isRecord(item) || typeof item.adminIcsId !== 'string' || typeof item.displayName !== 'string' || typeof item.icsUrl !== 'string') {
          return undefined;
        }

        const audienceGroups = dedupeAudienceGroups(normalizeAudienceGroups(item.audienceGroups));
        if (audienceGroups.length === 0 || !item.icsUrl.trim() || !item.displayName.trim()) {
          return undefined;
        }

        return {
          adminIcsId: item.adminIcsId.trim(),
          displayName: item.displayName.trim(),
          icsUrl: item.icsUrl.trim(),
          audienceGroups
        } as IAdminIcsCatalogItem;
      })
      .filter((item): item is IAdminIcsCatalogItem => !!item && !!item.adminIcsId)
    : [];

  const icsIdentityKeys = new Set<string>();
  for (const item of icsCatalog) {
    const identityKey = item.icsUrl.toLowerCase();
    if (icsIdentityKeys.has(identityKey)) {
      return undefined;
    }
    icsIdentityKeys.add(identityKey);
  }

  return {
    ...defaultAdminWebPartSettings,
    schemaVersion: typeof value.schemaVersion === 'number' ? value.schemaVersion : CALENDAR_SETTINGS_SCHEMA_VERSION,
    defaultView: value.defaultView === 'day' || value.defaultView === 'week' || value.defaultView === 'month'
      ? value.defaultView
      : defaultAdminWebPartSettings.defaultView,
    showWeekends: typeof value.showWeekends === 'boolean' ? value.showWeekends : defaultAdminWebPartSettings.showWeekends,
    startHour: typeof value.startHour === 'number' ? value.startHour : defaultAdminWebPartSettings.startHour,
    endHour: typeof value.endHour === 'number' ? value.endHour : defaultAdminWebPartSettings.endHour,
    slotDuration: typeof value.slotDuration === 'number' ? value.slotDuration : defaultAdminWebPartSettings.slotDuration,
    firstDayOfWeek: typeof value.firstDayOfWeek === 'number' ? value.firstDayOfWeek : defaultAdminWebPartSettings.firstDayOfWeek,
    useCustomProxy: typeof value.useCustomProxy === 'boolean' ? value.useCustomProxy : defaultAdminWebPartSettings.useCustomProxy,
    customProxyUrl: typeof value.customProxyUrl === 'string' ? value.customProxyUrl : defaultAdminWebPartSettings.customProxyUrl,
    useWhateverOrigin: typeof value.useWhateverOrigin === 'boolean' ? value.useWhateverOrigin : defaultAdminWebPartSettings.useWhateverOrigin,
    proxyPriority1: value.proxyPriority1 === 'custom' || value.proxyPriority1 === 'whateverorigin'
      ? value.proxyPriority1
      : defaultAdminWebPartSettings.proxyPriority1,
    proxyPriority2: value.proxyPriority2 === 'custom' || value.proxyPriority2 === 'whateverorigin'
      ? value.proxyPriority2
      : defaultAdminWebPartSettings.proxyPriority2,
    organizationPrimaryColor: typeof value.organizationPrimaryColor === 'string' && value.organizationPrimaryColor.trim()
      ? value.organizationPrimaryColor.trim()
      : defaultAdminWebPartSettings.organizationPrimaryColor,
    exchangeShowSourceLogo: typeof value.exchangeShowSourceLogo === 'boolean' ? value.exchangeShowSourceLogo : defaultAdminWebPartSettings.exchangeShowSourceLogo,
    sharePointShowSourceLogo: typeof value.sharePointShowSourceLogo === 'boolean' ? value.sharePointShowSourceLogo : defaultAdminWebPartSettings.sharePointShowSourceLogo,
    plannerShowSourceLogo: typeof value.plannerShowSourceLogo === 'boolean' ? value.plannerShowSourceLogo : defaultAdminWebPartSettings.plannerShowSourceLogo,
    unifiedGroupShowSourceLogo: typeof value.unifiedGroupShowSourceLogo === 'boolean' ? value.unifiedGroupShowSourceLogo : defaultAdminWebPartSettings.unifiedGroupShowSourceLogo,
    teamsShiftsShowSourceLogo: typeof value.teamsShiftsShowSourceLogo === 'boolean' ? value.teamsShiftsShowSourceLogo : defaultAdminWebPartSettings.teamsShiftsShowSourceLogo,
    plannerShowAllCalendars: typeof value.plannerShowAllCalendars === 'boolean' ? value.plannerShowAllCalendars : defaultAdminWebPartSettings.plannerShowAllCalendars,
    plannerShowAllAssignedToMeOnly: typeof value.plannerShowAllAssignedToMeOnly === 'boolean' ? value.plannerShowAllAssignedToMeOnly : defaultAdminWebPartSettings.plannerShowAllAssignedToMeOnly,
    unifiedGroupShowAllCalendars: typeof value.unifiedGroupShowAllCalendars === 'boolean' ? value.unifiedGroupShowAllCalendars : defaultAdminWebPartSettings.unifiedGroupShowAllCalendars,
    teamsShiftsShowAllCalendars: typeof value.teamsShiftsShowAllCalendars === 'boolean' ? value.teamsShiftsShowAllCalendars : defaultAdminWebPartSettings.teamsShiftsShowAllCalendars,
    assignedSources,
    icsCatalog
  };
}

export function parseAdminWebPartSettingsJson(raw: string | undefined): IAdminWebPartSettings | undefined {
  if (!raw) {
    return undefined;
  }

  try {
    return normalizeAdminWebPartSettings(JSON.parse(raw));
  } catch {
    return undefined;
  }
}

export function normalizeUserCalendarSettings(value: unknown): IUserCalendarSettings | undefined {
  if (!isRecord(value)) {
    return undefined;
  }

  const personalSources = Array.isArray(value.personalSources)
    ? value.personalSources
      .map(item => {
        if (!isRecord(item) || typeof item.userSourceId !== 'string') {
          return undefined;
        }
        const source = normalizeCalendarSourceBase(item);
        if (!source) {
          return undefined;
        }
        return {
          userSourceId: item.userSourceId.trim(),
          ...source
        } as IUserCalendarSource;
      })
      .filter((item): item is IUserCalendarSource => !!item && !!item.userSourceId)
    : [];

  const overrides: { [adminSourceId: string]: IAdminSourceOverride } = {};
  if (isRecord(value.adminSourceOverridesById)) {
    const overridesRecord = value.adminSourceOverridesById as Record<string, unknown>;
    Object.keys(overridesRecord).forEach(key => {
      const override = normalizeOverride(overridesRecord[key]);
      if (override) {
        overrides[key] = override;
      }
    });
  }

  const exchangeCalendarStates: { [calendarId: string]: boolean } = {};
  if (isRecord(value.exchangeCalendarStates)) {
    const exchangeStatesRecord = value.exchangeCalendarStates as Record<string, unknown>;
    Object.keys(exchangeStatesRecord).forEach(key => {
      const state = exchangeStatesRecord[key];
      if (typeof state === 'boolean') {
        exchangeCalendarStates[key] = state;
      }
    });
  }

  return {
    ...defaultUserCalendarSettings,
    schemaVersion: typeof value.schemaVersion === 'number' ? value.schemaVersion : CALENDAR_SETTINGS_SCHEMA_VERSION,
    defaultView: value.defaultView === 'day' || value.defaultView === 'week' || value.defaultView === 'month' ? value.defaultView : undefined,
    userStartHour: typeof value.userStartHour === 'number' ? value.userStartHour : undefined,
    userEndHour: typeof value.userEndHour === 'number' ? value.userEndHour : undefined,
    exchangeCalendarStates,
    exchangeShowSourceLogo: typeof value.exchangeShowSourceLogo === 'boolean' ? value.exchangeShowSourceLogo : undefined,
    sharePointShowSourceLogo: typeof value.sharePointShowSourceLogo === 'boolean' ? value.sharePointShowSourceLogo : undefined,
    plannerShowSourceLogo: typeof value.plannerShowSourceLogo === 'boolean' ? value.plannerShowSourceLogo : undefined,
    unifiedGroupShowSourceLogo: typeof value.unifiedGroupShowSourceLogo === 'boolean' ? value.unifiedGroupShowSourceLogo : undefined,
    teamsShiftsShowSourceLogo: typeof value.teamsShiftsShowSourceLogo === 'boolean' ? value.teamsShiftsShowSourceLogo : undefined,
    plannerShowAllCalendars: typeof value.plannerShowAllCalendars === 'boolean' ? value.plannerShowAllCalendars : undefined,
    plannerShowAllAssignedToMeOnly: typeof value.plannerShowAllAssignedToMeOnly === 'boolean' ? value.plannerShowAllAssignedToMeOnly : undefined,
    unifiedGroupShowAllCalendars: typeof value.unifiedGroupShowAllCalendars === 'boolean' ? value.unifiedGroupShowAllCalendars : undefined,
    teamsShiftsShowAllCalendars: typeof value.teamsShiftsShowAllCalendars === 'boolean' ? value.teamsShiftsShowAllCalendars : undefined,
    personalSources,
    adminSourceOverridesById: overrides
  };
}

function isLegacyCalendarSettings(value: unknown): value is ILegacyCalendarSettings {
  return isRecord(value) && Array.isArray(value.sources);
}

function normalizeLegacySource(value: unknown): IUserCalendarSource | undefined {
  if (!isRecord(value) || typeof value.id !== 'string') {
    return undefined;
  }

  const source = normalizeCalendarSourceBase(value);
  if (!source) {
    return undefined;
  }

  return {
    userSourceId: value.id.trim(),
    ...source
  };
}

export function migrateLegacyAdminSettings(value: unknown): IAdminWebPartSettings {
  if (!isLegacyCalendarSettings(value)) {
    return { ...defaultAdminWebPartSettings };
  }

  return {
    ...defaultAdminWebPartSettings,
    defaultView: value.defaultView === 'day' || value.defaultView === 'week' || value.defaultView === 'month'
      ? value.defaultView
      : defaultAdminWebPartSettings.defaultView,
    showWeekends: typeof value.showWeekends === 'boolean' ? value.showWeekends : defaultAdminWebPartSettings.showWeekends,
    startHour: typeof value.startHour === 'number' ? value.startHour : defaultAdminWebPartSettings.startHour,
    endHour: typeof value.endHour === 'number' ? value.endHour : defaultAdminWebPartSettings.endHour,
    slotDuration: typeof value.slotDuration === 'number' ? value.slotDuration : defaultAdminWebPartSettings.slotDuration,
    firstDayOfWeek: typeof value.firstDayOfWeek === 'number' ? value.firstDayOfWeek : defaultAdminWebPartSettings.firstDayOfWeek,
    useCustomProxy: typeof value.useCustomProxy === 'boolean' ? value.useCustomProxy : defaultAdminWebPartSettings.useCustomProxy,
    customProxyUrl: typeof value.customProxyUrl === 'string' ? value.customProxyUrl : defaultAdminWebPartSettings.customProxyUrl,
    useWhateverOrigin: typeof value.useWhateverOrigin === 'boolean' ? value.useWhateverOrigin : defaultAdminWebPartSettings.useWhateverOrigin,
    proxyPriority1: value.proxyPriority1 === 'custom' || value.proxyPriority1 === 'whateverorigin'
      ? value.proxyPriority1
      : defaultAdminWebPartSettings.proxyPriority1,
    proxyPriority2: value.proxyPriority2 === 'custom' || value.proxyPriority2 === 'whateverorigin'
      ? value.proxyPriority2
      : defaultAdminWebPartSettings.proxyPriority2,
    organizationPrimaryColor: typeof value.organizationPrimaryColor === 'string' ? value.organizationPrimaryColor : defaultAdminWebPartSettings.organizationPrimaryColor
  };
}

export function migrateLegacyUserSettings(value: unknown): IUserCalendarSettings {
  if (!isLegacyCalendarSettings(value)) {
    return { ...defaultUserCalendarSettings };
  }

  const personalSources = (value.sources || [])
    .map(item => normalizeLegacySource(item))
    .filter((item): item is IUserCalendarSource => !!item);

  return {
    ...defaultUserCalendarSettings,
    defaultView: value.defaultView,
    userStartHour: typeof value.userStartHour === 'number' ? value.userStartHour : undefined,
    userEndHour: typeof value.userEndHour === 'number' ? value.userEndHour : undefined,
    exchangeCalendarStates: isRecord(value.exchangeCalendarStates)
      ? Object.keys(value.exchangeCalendarStates).reduce<{ [calendarId: string]: boolean }>((acc, key) => {
        const state = value.exchangeCalendarStates?.[key];
        if (typeof state === 'boolean') {
          acc[key] = state;
        }
        return acc;
      }, {})
      : {},
    exchangeShowSourceLogo: typeof value.exchangeShowSourceLogo === 'boolean' ? value.exchangeShowSourceLogo : undefined,
    sharePointShowSourceLogo: typeof value.sharePointShowSourceLogo === 'boolean' ? value.sharePointShowSourceLogo : undefined,
    plannerShowSourceLogo: typeof value.plannerShowSourceLogo === 'boolean' ? value.plannerShowSourceLogo : undefined,
    unifiedGroupShowSourceLogo: typeof value.unifiedGroupShowSourceLogo === 'boolean' ? value.unifiedGroupShowSourceLogo : undefined,
    teamsShiftsShowSourceLogo: typeof value.teamsShiftsShowSourceLogo === 'boolean' ? value.teamsShiftsShowSourceLogo : undefined,
    plannerShowAllCalendars: typeof value.plannerShowAllCalendars === 'boolean' ? value.plannerShowAllCalendars : undefined,
    plannerShowAllAssignedToMeOnly: typeof value.plannerShowAllAssignedToMeOnly === 'boolean' ? value.plannerShowAllAssignedToMeOnly : undefined,
    unifiedGroupShowAllCalendars: typeof value.unifiedGroupShowAllCalendars === 'boolean' ? value.unifiedGroupShowAllCalendars : undefined,
    teamsShiftsShowAllCalendars: typeof value.teamsShiftsShowAllCalendars === 'boolean' ? value.teamsShiftsShowAllCalendars : undefined,
    personalSources,
    adminSourceOverridesById: {}
  };
}

export function loadAdminWebPartSettings(params: {
  current?: string;
  backup?: string;
  legacy?: string;
}): IAdminSettingsLoadResult {
  const current = parseAdminWebPartSettingsJson(params.current);
  if (current) {
    return { settings: current, source: 'current' };
  }

  const backup = parseAdminWebPartSettingsJson(params.backup);
  if (backup) {
    return {
      settings: backup,
      source: 'backup',
      notice: 'Current admin settings were invalid. The last known good configuration was loaded from backup.'
    };
  }

  if (params.legacy) {
    try {
      const legacy = JSON.parse(params.legacy);
      if (isLegacyCalendarSettings(legacy)) {
        return {
          settings: migrateLegacyAdminSettings(legacy),
          source: 'legacy',
          notice: 'Legacy web part settings were migrated to the new admin settings format.'
        };
      }
    } catch {
      // Ignore invalid legacy payloads.
    }
  }

  return {
    settings: { ...defaultAdminWebPartSettings },
    source: 'defaults',
    notice: params.current || params.backup
      ? 'Admin settings could not be recovered from current or backup data. Hardcoded defaults were loaded.'
      : undefined
  };
}

function applySourceOverride(source: ICalendarSourceBase, override: IAdminSourceOverride | undefined): ICalendarSourceBase | undefined {
  if (override?.removed) {
    return undefined;
  }

  return {
    ...source,
    name: override?.name || source.name,
    color: override?.color || source.color,
    isEnabled: typeof override?.isEnabled === 'boolean' ? override.isEnabled : source.isEnabled
  };
}

function createResolvedAdminSource(assignedSource: IAdminAssignedSource, override: IAdminSourceOverride | undefined): ICalendarSource | undefined {
  const source = applySourceOverride(assignedSource.source, override);
  if (!source) {
    return undefined;
  }

  return {
    id: assignedSource.adminSourceId,
    origin: 'admin',
    adminSourceId: assignedSource.adminSourceId,
    audienceGroupNames: assignedSource.audienceGroups.map(group => group.displayName),
    ...source
  };
}

function createResolvedUserSource(source: IUserCalendarSource): ICalendarSource {
  return {
    id: source.userSourceId,
    origin: 'user',
    userSourceId: source.userSourceId,
    ...stripRuntimeSource(source)
  };
}

export function resolveCalendarSettings(params: {
  adminSettings: IAdminWebPartSettings;
  userSettings: IUserCalendarSettings;
  matchedGroupIds: Set<string>;
  organizationPrimaryColor?: string;
}): ICalendarSettings {
  const { adminSettings, userSettings, matchedGroupIds, organizationPrimaryColor } = params;

  const resolvedAdminSources = adminSettings.assignedSources
    .filter(item => item.audienceGroups.some(group => matchedGroupIds.has(group.groupId)))
    .map(item => createResolvedAdminSource(item, userSettings.adminSourceOverridesById[item.adminSourceId]))
    .filter((item): item is ICalendarSource => !!item);

  const resolvedUserSources = userSettings.personalSources.map(createResolvedUserSource);

  const availableAdminIcsCatalogItems = adminSettings.icsCatalog
    .filter(item => item.audienceGroups.some(group => matchedGroupIds.has(group.groupId)));

  return {
    ...defaultCalendarSettings,
    schemaVersion: CALENDAR_SETTINGS_SCHEMA_VERSION,
    defaultView: userSettings.defaultView || adminSettings.defaultView,
    sources: [...resolvedAdminSources, ...resolvedUserSources],
    availableAdminIcsCatalogItems,
    showWeekends: adminSettings.showWeekends,
    startHour: adminSettings.startHour,
    endHour: adminSettings.endHour,
    slotDuration: adminSettings.slotDuration,
    firstDayOfWeek: adminSettings.firstDayOfWeek,
    userStartHour: userSettings.userStartHour,
    userEndHour: userSettings.userEndHour,
    useCustomProxy: adminSettings.useCustomProxy,
    customProxyUrl: adminSettings.customProxyUrl,
    useWhateverOrigin: adminSettings.useWhateverOrigin,
    proxyPriority1: adminSettings.proxyPriority1,
    proxyPriority2: adminSettings.proxyPriority2,
    organizationPrimaryColor: organizationPrimaryColor || adminSettings.organizationPrimaryColor || defaultCalendarSettings.organizationPrimaryColor,
    exchangeCalendarStates: { ...userSettings.exchangeCalendarStates },
    exchangeShowSourceLogo: userSettings.exchangeShowSourceLogo ?? adminSettings.exchangeShowSourceLogo,
    sharePointShowSourceLogo: userSettings.sharePointShowSourceLogo ?? adminSettings.sharePointShowSourceLogo,
    plannerShowSourceLogo: userSettings.plannerShowSourceLogo ?? adminSettings.plannerShowSourceLogo,
    unifiedGroupShowSourceLogo: userSettings.unifiedGroupShowSourceLogo ?? adminSettings.unifiedGroupShowSourceLogo,
    teamsShiftsShowSourceLogo: userSettings.teamsShiftsShowSourceLogo ?? adminSettings.teamsShiftsShowSourceLogo,
    plannerShowAllCalendars: userSettings.plannerShowAllCalendars ?? adminSettings.plannerShowAllCalendars,
    plannerShowAllAssignedToMeOnly: userSettings.plannerShowAllAssignedToMeOnly ?? adminSettings.plannerShowAllAssignedToMeOnly,
    unifiedGroupShowAllCalendars: userSettings.unifiedGroupShowAllCalendars ?? adminSettings.unifiedGroupShowAllCalendars,
    teamsShiftsShowAllCalendars: userSettings.teamsShiftsShowAllCalendars ?? adminSettings.teamsShiftsShowAllCalendars
  };
}

function createAdminSourceMap(adminSettings: IAdminWebPartSettings, matchedGroupIds: Set<string>): Record<string, ICalendarSourceBase> {
  return adminSettings.assignedSources.reduce<Record<string, ICalendarSourceBase>>((acc, item) => {
    if (item.audienceGroups.some(group => matchedGroupIds.has(group.groupId))) {
      acc[item.adminSourceId] = stripRuntimeSource(item.source);
    }
    return acc;
  }, {});
}

function copyExchangeCalendarStates(states: { [calendarId: string]: boolean } | undefined): { [calendarId: string]: boolean } {
  return states ? { ...states } : {};
}

export function deriveUserCalendarSettings(params: {
  nextResolvedSettings: ICalendarSettings;
  adminSettings: IAdminWebPartSettings;
  matchedGroupIds: Set<string>;
}): IUserCalendarSettings {
  const { nextResolvedSettings, adminSettings, matchedGroupIds } = params;
  const adminSourceMap = createAdminSourceMap(adminSettings, matchedGroupIds);
  const seenAdminSourceIds = new Set<string>();
  const personalSources: IUserCalendarSource[] = [];
  const adminSourceOverridesById: { [adminSourceId: string]: IAdminSourceOverride } = {};

  nextResolvedSettings.sources.forEach(source => {
    if (source.origin === 'admin' && source.adminSourceId) {
      const baseSource = adminSourceMap[source.adminSourceId];
      if (!baseSource) {
        return;
      }

      seenAdminSourceIds.add(source.adminSourceId);
      const override: IAdminSourceOverride = {};

      if (source.name !== baseSource.name) {
        override.name = source.name;
      }
      if (source.color !== baseSource.color) {
        override.color = source.color;
      }
      if (source.isEnabled !== baseSource.isEnabled) {
        override.isEnabled = source.isEnabled;
      }

      if (Object.keys(override).length > 0) {
        adminSourceOverridesById[source.adminSourceId] = override;
      }
      return;
    }

    if (source.origin !== 'admin') {
      personalSources.push(createUserCalendarSource(stripRuntimeSource(source), source.userSourceId || source.id));
    }
  });

  Object.keys(adminSourceMap).forEach(adminSourceId => {
    if (!seenAdminSourceIds.has(adminSourceId)) {
      adminSourceOverridesById[adminSourceId] = { removed: true };
    }
  });

  return {
    schemaVersion: CALENDAR_SETTINGS_SCHEMA_VERSION,
    defaultView: nextResolvedSettings.defaultView !== adminSettings.defaultView ? nextResolvedSettings.defaultView : undefined,
    userStartHour: nextResolvedSettings.userStartHour,
    userEndHour: nextResolvedSettings.userEndHour,
    exchangeCalendarStates: copyExchangeCalendarStates(nextResolvedSettings.exchangeCalendarStates),
    exchangeShowSourceLogo: nextResolvedSettings.exchangeShowSourceLogo !== adminSettings.exchangeShowSourceLogo ? nextResolvedSettings.exchangeShowSourceLogo : undefined,
    sharePointShowSourceLogo: nextResolvedSettings.sharePointShowSourceLogo !== adminSettings.sharePointShowSourceLogo ? nextResolvedSettings.sharePointShowSourceLogo : undefined,
    plannerShowSourceLogo: nextResolvedSettings.plannerShowSourceLogo !== adminSettings.plannerShowSourceLogo ? nextResolvedSettings.plannerShowSourceLogo : undefined,
    unifiedGroupShowSourceLogo: nextResolvedSettings.unifiedGroupShowSourceLogo !== adminSettings.unifiedGroupShowSourceLogo ? nextResolvedSettings.unifiedGroupShowSourceLogo : undefined,
    teamsShiftsShowSourceLogo: nextResolvedSettings.teamsShiftsShowSourceLogo !== adminSettings.teamsShiftsShowSourceLogo ? nextResolvedSettings.teamsShiftsShowSourceLogo : undefined,
    plannerShowAllCalendars: nextResolvedSettings.plannerShowAllCalendars !== adminSettings.plannerShowAllCalendars ? nextResolvedSettings.plannerShowAllCalendars : undefined,
    plannerShowAllAssignedToMeOnly: nextResolvedSettings.plannerShowAllAssignedToMeOnly !== adminSettings.plannerShowAllAssignedToMeOnly
      ? nextResolvedSettings.plannerShowAllAssignedToMeOnly
      : undefined,
    unifiedGroupShowAllCalendars: nextResolvedSettings.unifiedGroupShowAllCalendars !== adminSettings.unifiedGroupShowAllCalendars
      ? nextResolvedSettings.unifiedGroupShowAllCalendars
      : undefined,
    teamsShiftsShowAllCalendars: nextResolvedSettings.teamsShiftsShowAllCalendars !== adminSettings.teamsShiftsShowAllCalendars
      ? nextResolvedSettings.teamsShiftsShowAllCalendars
      : undefined,
    personalSources,
    adminSourceOverridesById
  };
}
