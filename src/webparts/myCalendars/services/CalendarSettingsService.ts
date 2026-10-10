import {
  CALENDAR_SETTINGS_SCHEMA_VERSION,
  type IAdminAssignedSource,
  type IAdminAllowedOverrides,
  type IExchangeMailboxAssignment,
  type IExchangeMailboxDiscovery,
  type IAdminSourceDefinition,
  defaultAllowedOverrides,
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
import { normalizeCacheDuration } from './CalendarEventCache';

export type AdminSettingsLoadSource = 'current' | 'defaults' | 'legacy';

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
const MIN_SUPPORTED_CALENDAR_SETTINGS_SCHEMA_VERSION = 2;

function hasSupportedSettingsSchemaVersion(value: Record<string, unknown>): boolean {
  return typeof value.schemaVersion === 'number' &&
    Number.isInteger(value.schemaVersion) &&
    value.schemaVersion >= MIN_SUPPORTED_CALENDAR_SETTINGS_SCHEMA_VERSION &&
    value.schemaVersion <= CALENDAR_SETTINGS_SCHEMA_VERSION;
}

function normalizeSlotDuration(value: unknown): 15 | 30 | 60 {
  return value === 15 || value === 30 || value === 60 ? value : 30;
}

function normalizeVisibleHourCount(value: unknown, legacyStart?: unknown, legacyEnd?: unknown): number {
  const migrated = typeof legacyStart === 'number' && typeof legacyEnd === 'number' ? legacyEnd - legacyStart : undefined;
  const candidate = typeof value === 'number' ? value : migrated;
  return Math.max(1, Math.min(24, Number.isFinite(candidate) ? Math.round(candidate as number) : 10));
}

function normalizePreferredStartMinutes(
  value: unknown,
  legacyStart: unknown,
  visibleHourCount: number,
  slotDuration: 15 | 30 | 60
): number {
  const migrated = typeof legacyStart === 'number' ? legacyStart * 60 : 8 * 60;
  const candidate = typeof value === 'number' ? value : migrated;
  const latestStart = Math.max(0, 24 * 60 - visibleHourCount * 60);
  const clamped = Math.max(0, Math.min(latestStart, Number.isFinite(candidate) ? candidate as number : migrated));
  return Math.floor(clamped / slotDuration) * slotDuration;
}

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
  if (typeof value.sharePointSiteName === 'string' && value.sharePointSiteName.trim()) {
    source.sharePointSiteName = value.sharePointSiteName.trim();
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
    .map((item): IAudienceGroup | undefined => {
      if (!isRecord(item) || typeof item.groupId !== 'string' || typeof item.displayName !== 'string') {
        return undefined;
      }

      return {
        groupId: item.groupId.trim(),
        displayName: item.displayName.trim(),
        groupType: item.groupType === 'microsoft365' || item.groupType === 'security' || item.groupType === 'mailEnabledSecurity' ? item.groupType : undefined
      };
    })
    .filter((item): item is IAudienceGroup => !!item && !!item.groupId && !!item.displayName);
}

function normalizeOverride(value: unknown): IAdminSourceOverride | undefined {
  if (!isRecord(value)) {
    return undefined;
  }

  const override: IAdminSourceOverride = {};
  ['showSourceLogo', 'plannerAssignedToMeOnly', 'showCompletedTasks'].forEach(key => {
    if (typeof value[key] === 'boolean') (override as Record<string, unknown>)[key] = value[key];
  });
  if (typeof value.removed === 'boolean') override.removed = value.removed;
  if (typeof value.isEnabled === 'boolean') override.isEnabled = value.isEnabled;
  if (typeof value.name === 'string' && value.name.trim()) override.name = value.name.trim();
  if (typeof value.color === 'string' && value.color.trim()) override.color = value.color.trim();

  return Object.keys(override).length > 0 ? override : undefined;
}


export function normalizeAllowedOverrides(value: unknown): IAdminAllowedOverrides {
  const raw = isRecord(value) ? value : {};
  return Object.keys(defaultAllowedOverrides).reduce((result, key) => {
    (result as unknown as Record<string, boolean>)[key] = raw[key] !== false;
    return result;
  }, { ...defaultAllowedOverrides });
}

export function normalizeAssignmentPolicy(value: { isMandatory?: unknown; defaultEnabled?: unknown }): { isMandatory: boolean; defaultEnabled: boolean } {
  return { isMandatory: value.isMandatory === true, defaultEnabled: value.isMandatory === true || value.defaultEnabled !== false };
}

export function audienceApplies(groups: IAudienceGroup[], matched: Set<string>): boolean {
  return groups.length === 0 || groups.some(group => matched.has(group.groupId));
}

/** Wire format keeps shared source and audience metadata out of assignment records. */
export function serializeAdminWebPartSettings(settings: IAdminWebPartSettings): string {
  const normalized = normalizeAdminWebPartSettings(settings);
  if (!normalized) throw new Error('Invalid administrator settings.');
  const { assignedSources, exchangeMailboxAssignments, icsCatalog, ...base } = normalized;
  return JSON.stringify({
    ...base,
    assignedSources: assignedSources.map(item => ({ assignmentId: item.assignmentId, adminSourceId: item.adminSourceId,
      audienceGroupIds: item.audienceGroups.map(group => group.groupId), ...normalizeAssignmentPolicy(item) })),
    exchangeMailboxAssignments: (exchangeMailboxAssignments || []).map(({ audienceGroups, ...item }) => ({
      ...item, audienceGroupIds: audienceGroups.map(group => group.groupId)
    })),
    icsCatalog: icsCatalog.map(({ audienceGroups, ...item }) => ({ ...item, audienceGroupIds: audienceGroups.map(group => group.groupId) }))
  });
}

export function canonicalizeExchangeSourceIdentities(settings: IAdminWebPartSettings, mailboxIds: Map<string, string>): IAdminWebPartSettings {
  const canonical = (source: ICalendarSourceBase): ICalendarSourceBase => {
    const id = source.sourceType === 'exchange' && source.exchangeMailbox ? mailboxIds.get(source.exchangeMailbox.toLowerCase()) : undefined;
    return id ? { ...source, exchangeMailbox: id } : source;
  };
  return { ...settings, assignedSources: settings.assignedSources.map(item => ({ ...item, source: canonical(item.source) })),
    sourceCatalog: (settings.sourceCatalog || []).map(item => ({ ...item, source: canonical(item.source) })) };
}

export function addExchangeMailboxAssignment(settings: IAdminWebPartSettings, rule: IExchangeMailboxAssignment): IAdminWebPartSettings {
  const normalized = normalizeAdminWebPartSettings(settings);
  if (!normalized) throw new Error('Invalid administrator settings.');
  const audienceKey = rule.audienceGroups.map(group => group.groupId).sort().join(',');
  const existing = (normalized.exchangeMailboxAssignments || []).find(item => item.mailboxId.toLowerCase() === rule.mailboxId.toLowerCase() && item.audienceGroups.map(group => group.groupId).sort().join(',') === audienceKey);
  if (existing) return { ...normalized, exchangeMailboxAssignments: (normalized.exchangeMailboxAssignments || []).map(item => item.assignmentId === existing.assignmentId ? { ...item, ...normalizeAssignmentPolicy(rule) } : item) };
  const transferred = normalized.assignedSources.filter(item => item.source.sourceType === 'exchange' && item.source.exchangeMailbox?.toLowerCase() === rule.mailboxId.toLowerCase() && item.audienceGroups.map(group => group.groupId).sort().join(',') === audienceKey);
  const exceptions = [...rule.exceptions];
  for (const item of transferred) {
    if (item.source.exchangeCalendarId && !exceptions.some(entry => entry.calendarId === item.source.exchangeCalendarId)) exceptions.push({ calendarId: item.source.exchangeCalendarId, ...normalizeAssignmentPolicy(item) });
  }
  return { ...normalized, assignedSources: normalized.assignedSources.filter(item => transferred.indexOf(item) < 0),
    exchangeMailboxAssignments: [...(normalized.exchangeMailboxAssignments || []), { ...rule, exceptions }] };
}

export function addAdministratorAssignments(settings: IAdminWebPartSettings, entries: Array<{ source: ICalendarSourceBase; policy: { isMandatory: boolean; defaultEnabled: boolean } }>, groups: IAudienceGroup[]): IAdminWebPartSettings {
  const assignedSources = settings.assignedSources.slice();
  const rules = (settings.exchangeMailboxAssignments || []).map(rule => ({ ...rule, exceptions: rule.exceptions.slice() }));
  const audienceKey = groups.map(group => group.groupId).sort().join(',');
  for (const entry of entries) {
    const identity = getSourceIdentityKey(entry.source);
    const existing = (settings.sourceCatalog || []).find(item => getSourceIdentityKey(item.source) === identity) || assignedSources.find(item => getSourceIdentityKey(item.source) === identity);
    const rule = entry.source.sourceType === 'exchange' ? rules.find(item => item.mailboxId.toLowerCase() === entry.source.exchangeMailbox?.toLowerCase() && item.audienceGroups.map(group => group.groupId).sort().join(',') === audienceKey) : undefined;
    if (rule) {
      rule.exceptions = rule.exceptions.filter(item => item.calendarId !== entry.source.exchangeCalendarId);
      rule.exceptions.push({ calendarId: entry.source.exchangeCalendarId as string, ...normalizeAssignmentPolicy(entry.policy) });
      continue;
    }
    const index = assignedSources.findIndex(item => getSourceIdentityKey(item.source) === identity && item.audienceGroups.map(group => group.groupId).sort().join(',') === audienceKey);
    if (index >= 0) assignedSources[index] = { ...assignedSources[index], ...normalizeAssignmentPolicy(entry.policy) };
    else assignedSources.push({ ...createAdminAssignedSource(existing?.source || entry.source, groups, existing?.adminSourceId), ...normalizeAssignmentPolicy(entry.policy), allowedOverrides: existing?.allowedOverrides || { ...defaultAllowedOverrides } });
  }
  return { ...settings, assignedSources, exchangeMailboxAssignments: rules };
}

export function getApplicableAdminSources(adminSettings: IAdminWebPartSettings, matched: Set<string>, discoveries: IExchangeMailboxDiscovery[] = [], currentUserMailboxId?: string): ICalendarSource[] {
  const candidates = adminSettings.assignedSources.filter(item => audienceApplies(item.audienceGroups, matched)).slice();
  for (const rule of adminSettings.exchangeMailboxAssignments || []) {
    if (!audienceApplies(rule.audienceGroups, matched)) continue;
    const discovery = discoveries.find(item => item.mailboxId.toLowerCase() === rule.mailboxId.toLowerCase());
    if (!discovery || discovery.error) continue;
    for (const source of discovery.sources) {
      const exception = rule.exceptions.find(item => item.calendarId === source.exchangeCalendarId);
      if (exception?.excluded) continue;
      const definition = (adminSettings.sourceCatalog || []).find(item => getSourceIdentityKey(item.source) === getSourceIdentityKey(source)) ||
        adminSettings.assignedSources.find(item => getSourceIdentityKey(item.source) === getSourceIdentityKey(source));
      candidates.push({ adminSourceId: definition?.adminSourceId || 'exchange|' + encodeURIComponent(rule.mailboxId.toLowerCase()) + '|' + encodeURIComponent(source.exchangeCalendarId || ''),
        source: definition?.source || source, audienceGroups: rule.audienceGroups,
        allowedOverrides: definition?.allowedOverrides || rule.allowedOverrides,
        ...normalizeAssignmentPolicy(exception && (exception.isMandatory !== undefined || exception.defaultEnabled !== undefined) ? exception : rule) });
    }
  }
  const result = new Map<string, ICalendarSource>();
  for (const item of candidates) {
    const key = getSourceIdentityKey(item.source.sourceType === 'exchange' && !item.source.exchangeMailbox && currentUserMailboxId ? { ...item.source, exchangeMailbox: currentUserMailboxId } : item.source);
    const policy = normalizeAssignmentPolicy({ isMandatory: item.isMandatory, defaultEnabled: item.defaultEnabled ?? item.source.isEnabled });
    const names = item.audienceGroups.length ? item.audienceGroups.filter(group => matched.has(group.groupId)).map(group => group.displayName) : ['__everyone__'];
    const previous = result.get(key);
    if (previous) {
      previous.isMandatory = previous.isMandatory || policy.isMandatory;
      previous.defaultEnabled = previous.isMandatory || previous.defaultEnabled || policy.defaultEnabled;
      previous.isEnabled = !!previous.defaultEnabled;
      previous.audienceGroupNames = Array.from(new Set([...(previous.audienceGroupNames || []), ...names]));
    } else result.set(key, { ...item.source, id: item.adminSourceId, adminSourceId: item.adminSourceId, origin: 'admin',
      ...policy, isEnabled: policy.defaultEnabled, allowedOverrides: item.allowedOverrides || { ...defaultAllowedOverrides }, audienceGroupNames: names });
  }
  return Array.from(result.values()).map(source => {
    if (source.allowedOverrides?.showSourceLogo !== false || source.showSourceLogo !== undefined) return source;
    const flag = source.sourceType === 'sharepoint' ? adminSettings.sharePointShowSourceLogo : source.sourceType === 'planner' ? adminSettings.plannerShowSourceLogo : source.sourceType === 'teamsShifts' ? adminSettings.teamsShiftsShowSourceLogo : source.sourceType === 'unifiedGroup' ? adminSettings.unifiedGroupShowSourceLogo : adminSettings.exchangeShowSourceLogo;
    return { ...source, showSourceLogo: flag };
  });
}

export function cleanAdminSourceOverrides(userSettings: IUserCalendarSettings, sources: ICalendarSource[], unresolvedIds: string[] = []): IUserCalendarSettings {
  const byId = new Map(sources.map(source => [source.id, source]));
  const overrides: Record<string, IAdminSourceOverride> = {};
  for (const id of Object.keys(userSettings.adminSourceOverridesById)) {
    const source = byId.get(id);
    if (!source) {
      if (unresolvedIds.indexOf(id) >= 0) overrides[id] = { ...userSettings.adminSourceOverridesById[id] };
      continue;
    }
    const raw = userSettings.adminSourceOverridesById[id];
    const clean: IAdminSourceOverride = {};
    if (!source.isMandatory) {
      if (typeof raw.isEnabled === 'boolean') clean.isEnabled = raw.isEnabled;
      if (raw.removed) clean.removed = true;
    }
    for (const key of Object.keys(defaultAllowedOverrides) as Array<keyof IAdminAllowedOverrides>) {
      if (source.allowedOverrides?.[key] !== false && raw[key] !== undefined) (clean as Record<string, unknown>)[key] = raw[key];
    }
    if (Object.keys(clean).length) overrides[id] = clean;
  }
  return { ...userSettings, adminSourceOverridesById: overrides };
}

export function isAutomaticExchangeCalendarAssigned(settings: ICalendarSettings, calendarId: string): boolean {
  return (settings.applicableAdminSources || []).some(source => source.sourceType === 'exchange' && source.exchangeCalendarId === calendarId &&
    (!source.exchangeMailbox || source.exchangeMailbox.toLowerCase() === settings.currentUserMailboxId?.toLowerCase()));
}

export function applyAdminSourceChanges(source: ICalendarSource, updates: Partial<ICalendarSource>): ICalendarSource {
  if (source.origin !== 'admin') return { ...source, ...updates };
  const accepted: Partial<ICalendarSource> = {};
  for (const key of Object.keys(defaultAllowedOverrides) as Array<keyof IAdminAllowedOverrides>) {
    if (source.allowedOverrides?.[key] !== false && updates[key] !== undefined) (accepted as Record<string, unknown>)[key] = updates[key];
  }
  if (!source.isMandatory && typeof updates.isEnabled === 'boolean') {
    accepted.isEnabled = updates.isEnabled;
    accepted.visibilityOverride = updates.isEnabled;
  }
  if (!source.isMandatory && Object.prototype.hasOwnProperty.call(updates, 'visibilityOverride') && updates.visibilityOverride === undefined) {
    accepted.isEnabled = !!source.defaultEnabled;
    accepted.visibilityOverride = undefined;
  }
  return { ...source, ...accepted };
}

/** An open personal draft must observe Mandatory membership without losing other edits. */
export function restoreMandatoryAdminSourcesInDraft(draft: ICalendarSettings, effective: ICalendarSettings): ICalendarSettings {
  const mandatory = new Map(effective.sources.filter(source => source.origin === 'admin' && source.isMandatory).map(source => [source.id, source]));
  const sources = draft.sources.map(source => {
    const current = source.origin === 'admin' ? mandatory.get(source.id) : undefined;
    if (!current) return source;
    mandatory.delete(source.id);
    return applyAdminSourceChanges(current, source);
  });
  mandatory.forEach(source => sources.push({ ...source }));
  return JSON.stringify(sources) === JSON.stringify(draft.sources) ? draft : { ...draft, sources };
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
    assignmentId: generateStableId('assignment'),
    isMandatory: false,
    defaultEnabled: source.isEnabled,
    allowedOverrides: { ...defaultAllowedOverrides },
    adminSourceId: adminSourceId || generateStableId('adminSource'),
    source: stripRuntimeSource(source),
    audienceGroups: audienceGroups.map(group => ({
      groupId: group.groupId,
      displayName: group.displayName,
      groupType: group.groupType
    }))
  };
}

export function getSourceIdentityKey(source: ICalendarSourceBase): string {
  switch (source.sourceType) {
    case 'sharepoint':
      return `${source.sourceType}|${source.sharePointSiteId || ''}|${source.sharePointListId || ''}`;
    case 'exchange':
      return `${source.sourceType}|${(source.exchangeMailbox || 'me').toLowerCase()}|${source.exchangeCalendarId || 'calendar'}`;
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
  if (!isRecord(value) || !hasSupportedSettingsSchemaVersion(value)) {
    return undefined;
  }

  const groupRegistry = dedupeAudienceGroups(normalizeAudienceGroups(value.audienceGroups));
  const storedCatalog = Array.isArray(value.sourceCatalog) ? value.sourceCatalog : [];
  const catalog: IAdminSourceDefinition[] = [];
  for (const item of storedCatalog) {
    if (!isRecord(item) || typeof item.adminSourceId !== 'string') continue;
    const source = normalizeCalendarSourceBase(item.source);
    if (source) catalog.push({ adminSourceId: item.adminSourceId, source, allowedOverrides: normalizeAllowedOverrides(item.allowedOverrides) });
  }
  const readGroups = (item: Record<string, unknown>): IAudienceGroup[] | undefined => {
    if (Array.isArray(item.audienceGroupIds)) {
      const ids = item.audienceGroupIds;
      if (ids.some(id => typeof id !== 'string' || !groupRegistry.some(group => group.groupId === id))) return undefined;
      return groupRegistry.filter(group => ids.indexOf(group.groupId) >= 0);
    }
    if (!Array.isArray(item.audienceGroups)) return undefined;
    const groups = dedupeAudienceGroups(normalizeAudienceGroups(item.audienceGroups));
    return item.audienceGroups.length && !groups.length ? undefined : groups;
  };
  const assignedSources = Array.isArray(value.assignedSources)
    ? value.assignedSources
      .map(item => {
        if (!isRecord(item) || typeof item.adminSourceId !== 'string') {
          return undefined;
        }

        const definition = catalog.find(entry => entry.adminSourceId === item.adminSourceId);
        const source = normalizeCalendarSourceBase(item.source) || definition?.source;
        const audienceGroups = readGroups(item);
        if (!source || !audienceGroups) {
          return undefined;
        }

        return {
          adminSourceId: item.adminSourceId.trim(),
          source,
          audienceGroups,
          assignmentId: typeof item.assignmentId === 'string' ? item.assignmentId : 'assignment_' + item.adminSourceId.trim(),
          isMandatory: item.isMandatory === true,
          defaultEnabled: item.isMandatory === true || (typeof item.defaultEnabled === 'boolean' ? item.defaultEnabled : source.isEnabled),
          allowedOverrides: normalizeAllowedOverrides(item.allowedOverrides || definition?.allowedOverrides)
        } as IAdminAssignedSource;
      })
      .filter((item): item is IAdminAssignedSource => !!item && !!item.adminSourceId)
    : [];

  const assignmentKeys = new Set<string>();
  for (const item of assignedSources) {
    const identityKey = getSourceIdentityKey(item.source);
    for (const audienceId of item.audienceGroups.length ? item.audienceGroups.map(group => group.groupId) : ['__everyone__']) {
      const key = identityKey + '|' + audienceId;
      if (assignmentKeys.has(key)) return undefined;
      assignmentKeys.add(key);
    }
    const definition = catalog.find(entry => getSourceIdentityKey(entry.source) === identityKey);
    if (definition) {
      item.adminSourceId = definition.adminSourceId;
      // Draft changes update all occurrences of the shared source together.
      if (!Array.isArray(value.sourceCatalog) || isRecord((value.assignedSources as unknown[]).find(raw => isRecord(raw) && raw.adminSourceId === item.adminSourceId && raw.source))) {
        definition.source = item.source;
        definition.allowedOverrides = item.allowedOverrides || normalizeAllowedOverrides(undefined);
      }
      item.source = definition.source;
      item.allowedOverrides = definition.allowedOverrides;
    } else catalog.push({ adminSourceId: item.adminSourceId, source: item.source, allowedOverrides: item.allowedOverrides || normalizeAllowedOverrides(undefined) });
  }

  const exchangeMailboxAssignments: IExchangeMailboxAssignment[] = [];
  for (const raw of Array.isArray(value.exchangeMailboxAssignments) ? value.exchangeMailboxAssignments : []) {
    if (!isRecord(raw) || typeof raw.assignmentId !== 'string' || typeof raw.mailboxId !== 'string' || !raw.mailboxId.trim()) continue;
    const audienceGroups = readGroups(raw);
    if (!audienceGroups) continue;
    exchangeMailboxAssignments.push({
      assignmentId: raw.assignmentId, mailboxId: raw.mailboxId.trim(),
      mailboxDisplayName: typeof raw.mailboxDisplayName === 'string' ? raw.mailboxDisplayName : raw.mailboxId,
      audienceGroups, ...normalizeAssignmentPolicy(raw), allowedOverrides: normalizeAllowedOverrides(raw.allowedOverrides),
      exceptions: (Array.isArray(raw.exceptions) ? raw.exceptions : []).filter(item => isRecord(item) && typeof item.calendarId === 'string')
        .map(item => ({ calendarId: item.calendarId as string, excluded: item.excluded === true,
          ...(typeof item.isMandatory === 'boolean' || typeof item.defaultEnabled === 'boolean' ? normalizeAssignmentPolicy(item) : {}) }))
    });
  }

  const icsCatalog = Array.isArray(value.icsCatalog)
    ? value.icsCatalog
      .map(item => {
        if (!isRecord(item) || typeof item.adminIcsId !== 'string' || typeof item.displayName !== 'string' || typeof item.icsUrl !== 'string') {
          return undefined;
        }

        const audienceGroups = readGroups(item);
        if (!audienceGroups || !item.icsUrl.trim() || !item.displayName.trim()) {
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

  const slotDurationMinutes = normalizeSlotDuration(value.slotDurationMinutes ?? value.slotDuration);
  const visibleHourCount = normalizeVisibleHourCount(value.visibleHourCount, value.startHour, value.endHour);

  return {
    ...defaultAdminWebPartSettings,
    schemaVersion: CALENDAR_SETTINGS_SCHEMA_VERSION,
    defaultView: value.defaultView === 'day' || value.defaultView === 'week' || value.defaultView === 'month'
      ? value.defaultView
      : defaultAdminWebPartSettings.defaultView,
    showWeekends: typeof value.showWeekends === 'boolean' ? value.showWeekends : defaultAdminWebPartSettings.showWeekends,
    preferredStartMinutes: normalizePreferredStartMinutes(value.preferredStartMinutes, value.startHour, visibleHourCount, slotDurationMinutes),
    visibleHourCount,
    slotDurationMinutes,
    enableCache: typeof value.enableCache === 'boolean' ? value.enableCache : defaultAdminWebPartSettings.enableCache,
    cacheDurationMinutes: normalizeCacheDuration(value.cacheDurationMinutes),
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
    sourceCatalog: catalog,
    audienceGroups: dedupeAudienceGroups([...groupRegistry, ...assignedSources.reduce<IAudienceGroup[]>((all, item) => all.concat(item.audienceGroups), []), ...exchangeMailboxAssignments.reduce<IAudienceGroup[]>((all, item) => all.concat(item.audienceGroups), []), ...icsCatalog.reduce<IAudienceGroup[]>((all, item) => all.concat(item.audienceGroups), [])]),
    exchangeMailboxAssignments,
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
  if (!isRecord(value) || !hasSupportedSettingsSchemaVersion(value)) {
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
    schemaVersion: CALENDAR_SETTINGS_SCHEMA_VERSION,
    defaultView: value.defaultView === 'day' || value.defaultView === 'week' || value.defaultView === 'month' ? value.defaultView : undefined,
    userShowWeekends: typeof value.userShowWeekends === 'boolean' ? value.userShowWeekends : undefined,
    userPreferredStartMinutes: typeof value.userPreferredStartMinutes === 'number'
      ? value.userPreferredStartMinutes
      : typeof value.userStartHour === 'number' ? value.userStartHour * 60 : undefined,
    userVisibleHourCount: typeof value.userVisibleHourCount === 'number'
      ? normalizeVisibleHourCount(value.userVisibleHourCount)
      : typeof value.userStartHour === 'number' && typeof value.userEndHour === 'number'
        ? normalizeVisibleHourCount(undefined, value.userStartHour, value.userEndHour)
        : undefined,
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

  const slotDurationMinutes = normalizeSlotDuration(value.slotDuration);
  const visibleHourCount = normalizeVisibleHourCount(undefined, value.startHour, value.endHour);

  return {
    ...defaultAdminWebPartSettings,
    defaultView: value.defaultView === 'day' || value.defaultView === 'week' || value.defaultView === 'month'
      ? value.defaultView
      : defaultAdminWebPartSettings.defaultView,
    showWeekends: typeof value.showWeekends === 'boolean' ? value.showWeekends : defaultAdminWebPartSettings.showWeekends,
    preferredStartMinutes: normalizePreferredStartMinutes(undefined, value.startHour, visibleHourCount, slotDurationMinutes),
    visibleHourCount,
    slotDurationMinutes,
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
    userShowWeekends: undefined,
    userPreferredStartMinutes: typeof value.userStartHour === 'number' ? value.userStartHour * 60 : undefined,
    userVisibleHourCount: typeof value.userStartHour === 'number' && typeof value.userEndHour === 'number'
      ? normalizeVisibleHourCount(undefined, value.userStartHour, value.userEndHour)
      : undefined,
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
  legacy?: string;
}): IAdminSettingsLoadResult {
  const current = parseAdminWebPartSettingsJson(params.current);
  if (current) {
    return { settings: current, source: 'current' };
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
    notice: params.current
      ? 'Admin settings could not be recovered from current data. Hardcoded defaults were loaded.'
      : undefined
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
  mailboxDiscoveries?: IExchangeMailboxDiscovery[];
  currentUserMailboxId?: string;
}): ICalendarSettings {
  const { adminSettings, userSettings, matchedGroupIds, organizationPrimaryColor } = params;

  const applicableAdminSources = getApplicableAdminSources(adminSettings, matchedGroupIds, params.mailboxDiscoveries, params.currentUserMailboxId);
  const cleaned = cleanAdminSourceOverrides(userSettings, applicableAdminSources);
  const resolvedAdminSources = applicableAdminSources.filter(source => !cleaned.adminSourceOverridesById[source.id]?.removed).map(source => {
    const override = cleaned.adminSourceOverridesById[source.id];
    return { ...source, ...override, isEnabled: source.isMandatory ? true : override?.isEnabled ?? source.defaultEnabled ?? true,
      visibilityOverride: source.isMandatory ? undefined : override?.isEnabled };
  });
  const identity = (source: ICalendarSourceBase): string => getSourceIdentityKey(source.sourceType === 'exchange' && !source.exchangeMailbox && params.currentUserMailboxId ? { ...source, exchangeMailbox: params.currentUserMailboxId } : source);
  const adminIdentities = new Set(applicableAdminSources.map(identity));
  const resolvedUserSources = userSettings.personalSources.filter(source => !adminIdentities.has(identity(source))).map(createResolvedUserSource);
  const availableAdminIcsCatalogItems = adminSettings.icsCatalog.filter(item => audienceApplies(item.audienceGroups, matchedGroupIds));
  const failedMailboxes = (params.mailboxDiscoveries || []).filter(item => item.error);
  const unresolvedAdminSourceIds = Object.keys(userSettings.adminSourceOverridesById).filter(id => failedMailboxes.some(mailbox =>
    id.indexOf('exchange|' + encodeURIComponent(mailbox.mailboxId.toLowerCase()) + '|') === 0 ||
    adminSettings.sourceCatalog?.some(source => source.adminSourceId === id && source.source.exchangeMailbox?.toLowerCase() === mailbox.mailboxId.toLowerCase())));
  const userVisibleHourCount = userSettings.userVisibleHourCount === undefined
    ? undefined
    : normalizeVisibleHourCount(userSettings.userVisibleHourCount);
  const userPreferredStartMinutes = userSettings.userPreferredStartMinutes === undefined
    ? undefined
    : normalizePreferredStartMinutes(
      userSettings.userPreferredStartMinutes,
      undefined,
      userVisibleHourCount ?? adminSettings.visibleHourCount,
      adminSettings.slotDurationMinutes
    );

  return {
    ...defaultCalendarSettings,
    applicableAdminSources,
    currentUserMailboxId: params.currentUserMailboxId,
    unresolvedAdminSourceIds,
    adminExchangeDiscoveryErrors: failedMailboxes.map(item => (adminSettings.exchangeMailboxAssignments?.find(rule => rule.mailboxId === item.mailboxId)?.mailboxDisplayName || item.mailboxId) + ': ' + item.error),
    schemaVersion: CALENDAR_SETTINGS_SCHEMA_VERSION,
    defaultView: userSettings.defaultView || adminSettings.defaultView,
    sources: [...resolvedAdminSources, ...resolvedUserSources],
    availableAdminIcsCatalogItems,
    adminShowWeekends: adminSettings.showWeekends,
    showWeekends: userSettings.userShowWeekends ?? adminSettings.showWeekends,
    userShowWeekends: userSettings.userShowWeekends,
    preferredStartMinutes: adminSettings.preferredStartMinutes,
    visibleHourCount: adminSettings.visibleHourCount,
    slotDurationMinutes: adminSettings.slotDurationMinutes,
    enableCache: adminSettings.enableCache,
    cacheDurationMinutes: adminSettings.cacheDurationMinutes,
    userPreferredStartMinutes,
    userVisibleHourCount,
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

function copyExchangeCalendarStates(states: { [calendarId: string]: boolean } | undefined): { [calendarId: string]: boolean } {
  return states ? { ...states } : {};
}

export function deriveUserCalendarSettings(params: {
  nextResolvedSettings: ICalendarSettings;
  adminSettings: IAdminWebPartSettings;
  matchedGroupIds: Set<string>;
  existingUserSettings?: IUserCalendarSettings;
  mailboxDiscoveries?: IExchangeMailboxDiscovery[];
  currentUserMailboxId?: string;
}): IUserCalendarSettings {
  const { nextResolvedSettings, adminSettings, matchedGroupIds, existingUserSettings } = params;
  const baseline = params.mailboxDiscoveries !== undefined
    ? getApplicableAdminSources(adminSettings, matchedGroupIds, params.mailboxDiscoveries, params.currentUserMailboxId)
    : nextResolvedSettings.applicableAdminSources || getApplicableAdminSources(adminSettings, matchedGroupIds);
  const adminSourceMap = new Map(baseline.map(source => [source.id, source]));
  const cleaned = cleanAdminSourceOverrides(existingUserSettings || defaultUserCalendarSettings, baseline, nextResolvedSettings.unresolvedAdminSourceIds);
  const seen = new Set<string>();
  const personalSources: IUserCalendarSource[] = [];
  const adminSourceOverridesById: Record<string, IAdminSourceOverride> = {};
  for (const id of nextResolvedSettings.unresolvedAdminSourceIds || []) {
    if (cleaned.adminSourceOverridesById[id]) adminSourceOverridesById[id] = cleaned.adminSourceOverridesById[id];
  }
  for (const source of nextResolvedSettings.sources) {
    if (source.origin !== 'admin') { personalSources.push(createUserCalendarSource(stripRuntimeSource(source), source.userSourceId || source.id)); continue; }
    const base = adminSourceMap.get(source.id);
    if (!base) continue;
    seen.add(source.id);
    const override: IAdminSourceOverride = {};
    for (const key of Object.keys(defaultAllowedOverrides) as Array<keyof IAdminAllowedOverrides>) {
      if (base.allowedOverrides?.[key] !== false && source[key] !== base[key] && source[key] !== undefined) (override as Record<string, unknown>)[key] = source[key];
    }
    if (!base.isMandatory && (source.visibilityOverride !== undefined || source.isEnabled !== base.defaultEnabled)) override.isEnabled = source.isEnabled;
    if (Object.keys(override).length) adminSourceOverridesById[source.id] = override;
  }
  for (const base of baseline) {
    if (!seen.has(base.id) && !base.isMandatory) adminSourceOverridesById[base.id] = { ...cleaned.adminSourceOverridesById[base.id], removed: true };
  }
  // Preserve personal entries suppressed by a matching administrator source.
  const identity = (source: ICalendarSourceBase): string => getSourceIdentityKey(source.sourceType === 'exchange' && !source.exchangeMailbox && nextResolvedSettings.currentUserMailboxId ? { ...source, exchangeMailbox: nextResolvedSettings.currentUserMailboxId } : source);
  const adminIdentities = new Set(baseline.map(identity));
  for (const source of existingUserSettings?.personalSources || []) {
    if (adminIdentities.has(identity(source)) && !personalSources.some(item => item.userSourceId === source.userSourceId)) personalSources.push(source);
  }

  return {
    schemaVersion: CALENDAR_SETTINGS_SCHEMA_VERSION,
    defaultView: existingUserSettings?.defaultView !== undefined || nextResolvedSettings.defaultView !== adminSettings.defaultView
      ? nextResolvedSettings.defaultView
      : undefined,
    userShowWeekends: nextResolvedSettings.userShowWeekends !== undefined || nextResolvedSettings.showWeekends !== adminSettings.showWeekends
      ? nextResolvedSettings.showWeekends
      : undefined,
    userPreferredStartMinutes: nextResolvedSettings.userPreferredStartMinutes,
    userVisibleHourCount: nextResolvedSettings.userVisibleHourCount,
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
