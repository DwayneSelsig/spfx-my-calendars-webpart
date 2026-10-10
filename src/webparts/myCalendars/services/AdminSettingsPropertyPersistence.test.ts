import { defaultAdminWebPartSettings, defaultUserCalendarSettings, defaultAllowedOverrides, type IAdminWebPartSettings } from '../models/ICalendarSettings';
import { loadAdminWebPartSettings, normalizeAdminWebPartSettings, resolveCalendarSettings, serializeAdminWebPartSettings } from './CalendarSettingsService';
import {
  notifyAdminSettingsPropertyChanges,
  persistAdminWebPartSettings,
  type IAdminSettingsPropertyBag
} from './AdminSettingsPropertyPersistence';

describe('administrator settings property persistence', () => {
  it('notifies SPFx before synchronizing the complete draft to current and backup', () => {
    const settings = {
      ...defaultAdminWebPartSettings,
      defaultView: 'week' as const,
      showWeekends: false,
      visibleHourCount: 8,
      slotDurationMinutes: 60 as const,
      enableCache: true,
      cacheDurationMinutes: 37,
      plannerShowAllCalendars: false
    };
    const properties: IAdminSettingsPropertyBag = {};
    let propertiesAtNotification: IAdminSettingsPropertyBag | undefined;

    const serialized = persistAdminWebPartSettings(properties, settings, value => {
      expect(JSON.parse(value).cacheDurationMinutes).toBe(37);
      propertiesAtNotification = { ...properties };
    });

    expect(properties.adminSettings).toBe(serialized);
    expect(properties.adminSettingsBackup).toBe(serialized);
    expect(propertiesAtNotification).toEqual({});
    expect(JSON.parse(serialized)).toMatchObject({
      defaultView: 'week',
      showWeekends: false,
      visibleHourCount: 8,
      slotDurationMinutes: 60,
      enableCache: true,
      cacheDurationMinutes: 37,
      plannerShowAllCalendars: false
    });
  });

  it('reports both persisted properties through the SPFx change callback', () => {
    const changeCallback = jest.fn();

    notifyAdminSettingsPropertyChanges(changeCallback, 'adminSettings', 'adminSettingsBackup', '{"cacheDurationMinutes":37}');

    expect(changeCallback).toHaveBeenNthCalledWith(1, 'adminSettings', '{"cacheDurationMinutes":37}', true);
    expect(changeCallback).toHaveBeenNthCalledWith(2, 'adminSettingsBackup', '{"cacheDurationMinutes":37}', true);
  });
});


function existingAssignmentSettings(isMandatory = false): IAdminWebPartSettings {
  const normalized = normalizeAdminWebPartSettings({
    ...defaultAdminWebPartSettings,
    assignedSources: [{
      assignmentId: 'existing-assignment', adminSourceId: 'existing-source',
      source: { sourceType: 'exchange', exchangeMailbox: 'mailbox-object-id', exchangeCalendarId: 'calendar-id',
        name: 'Existing calendar', color: '#0078d4', isEnabled: true },
      audienceGroups: [{ groupId: 'target-group', displayName: 'Target group', groupType: 'security' }],
      isMandatory, defaultEnabled: true, allowedOverrides: { ...defaultAllowedOverrides }
    }]
  });
  if (!normalized) throw new Error('Invalid test assignment');
  return normalized;
}

// Model SPFx's host snapshot: after a reset, the first notification establishes
// its baseline; a subsequent distinct state is delivered to the page via setDirty.
function saveToPage(properties: IAdminSettingsPropertyBag, draft: IAdminWebPartSettings): IAdminSettingsPropertyBag {
  let previousSnapshot: string | undefined;
  let pageProperties: IAdminSettingsPropertyBag | undefined;
  persistAdminWebPartSettings(properties, draft, serialized => {
    notifyAdminSettingsPropertyChanges((targetProperty, value) => {
      if (targetProperty !== 'adminSettings' && targetProperty !== 'adminSettingsBackup') throw new Error('Unexpected property');
      properties[targetProperty] = value as string;
      const snapshot = JSON.stringify(properties);
      if (previousSnapshot !== undefined && previousSnapshot !== snapshot) pageProperties = { ...properties };
      previousSnapshot = snapshot;
    }, 'adminSettings', 'adminSettingsBackup', serialized);
  });
  if (!pageProperties) throw new Error('The host did not receive changed page properties');
  return pageProperties;
}

describe('existing administrator assignment persistence', () => {
  it.each([
    [false, true, true],
    [true, false, true],
    [false, false, false]
  ])('persists mandatory %s -> %s and default enabled %s after a host snapshot reset', (initialMandatory, isMandatory, defaultEnabled) => {
    const initial = existingAssignmentSettings(initialMandatory);
    const oldJson = serializeAdminWebPartSettings(initial);
    const properties = { adminSettings: oldJson, adminSettingsBackup: oldJson };
    const draft = { ...initial, assignedSources: initial.assignedSources.map(item => ({ ...item, isMandatory, defaultEnabled })) };
    const pageProperties = saveToPage(properties, draft);
    expect(pageProperties.adminSettings).not.toBe(oldJson);
    expect(pageProperties.adminSettingsBackup).toBe(pageProperties.adminSettings);
    const reloaded = loadAdminWebPartSettings({ current: pageProperties.adminSettings, backup: pageProperties.adminSettingsBackup });
    expect(reloaded.source).toBe('current');
    expect(reloaded.settings.assignedSources).toHaveLength(1);
    expect(reloaded.settings.assignedSources[0]).toMatchObject({
      assignmentId: 'existing-assignment', adminSourceId: 'existing-source', isMandatory, defaultEnabled,
      source: { exchangeMailbox: 'mailbox-object-id', exchangeCalendarId: 'calendar-id' },
      audienceGroups: [{ groupId: 'target-group' }]
    });
    const backup = loadAdminWebPartSettings({ current: '{invalid', backup: pageProperties.adminSettingsBackup });
    expect(backup.source).toBe('backup');
    expect(backup.settings).toEqual(reloaded.settings);
  });

  it('publishes changed name and color on the existing shared definition', () => {
    const initial = existingAssignmentSettings();
    const oldJson = serializeAdminWebPartSettings(initial);
    const updates = { name: 'Renamed calendar', color: '#ff0000' };
    const draft = { ...initial,
      assignedSources: initial.assignedSources.map(item => ({ ...item, source: { ...item.source, ...updates } })),
      sourceCatalog: initial.sourceCatalog?.map(item => ({ ...item, source: { ...item.source, ...updates } }))
    };
    const page = saveToPage({ adminSettings: oldJson, adminSettingsBackup: oldJson }, draft);
    const reloaded = loadAdminWebPartSettings({ current: page.adminSettings }).settings;
    expect(reloaded.assignedSources[0]).toMatchObject({ adminSourceId: 'existing-source', assignmentId: 'existing-assignment', source: updates });
    expect(reloaded.sourceCatalog?.[0]).toMatchObject({ adminSourceId: 'existing-source', source: updates });
  });

  it('enables a published mandatory calendar despite previous personal hiding and removal', () => {
    const initial = existingAssignmentSettings();
    const oldJson = serializeAdminWebPartSettings(initial);
    const draft = { ...initial, assignedSources: initial.assignedSources.map(item => ({ ...item, isMandatory: true, defaultEnabled: true })) };
    const page = saveToPage({ adminSettings: oldJson, adminSettingsBackup: oldJson }, draft);
    const reloaded = loadAdminWebPartSettings({ current: page.adminSettings }).settings;
    const user = { ...defaultUserCalendarSettings, adminSourceOverridesById: { 'existing-source': { isEnabled: false, removed: true } } };
    const result = resolveCalendarSettings({ adminSettings: reloaded, userSettings: user, matchedGroupIds: new Set(['target-group']) });
    expect(result.sources).toHaveLength(1);
    expect(result.sources[0]).toMatchObject({ id: 'existing-source', isMandatory: true, isEnabled: true });
    expect(resolveCalendarSettings({ adminSettings: reloaded, userSettings: user, matchedGroupIds: new Set() }).sources).toHaveLength(0);
    expect(user.adminSourceOverridesById['existing-source']).toEqual({ isEnabled: false, removed: true });
  });

  it('lets SPFx observe each property transition before local synchronization', () => {
    const initial = existingAssignmentSettings();
    const oldJson = serializeAdminWebPartSettings(initial);
    const properties = { adminSettings: oldJson, adminSettingsBackup: oldJson };
    const observed: IAdminSettingsPropertyBag[] = [];
    const draft = { ...initial, assignedSources: initial.assignedSources.map(item => ({ ...item, isMandatory: true })) };
    const serialized = persistAdminWebPartSettings(properties, draft, value => {
      notifyAdminSettingsPropertyChanges((targetProperty, newValue) => {
        observed.push({ ...properties });
        if (targetProperty === 'adminSettings' || targetProperty === 'adminSettingsBackup') properties[targetProperty] = newValue as string;
      }, 'adminSettings', 'adminSettingsBackup', value);
    });
    expect(observed).toEqual([
      { adminSettings: oldJson, adminSettingsBackup: oldJson },
      { adminSettings: serialized, adminSettingsBackup: oldJson }
    ]);
    expect(properties).toEqual({ adminSettings: serialized, adminSettingsBackup: serialized });
  });

  it('does not synchronize new properties when the SPFx notification is rejected', () => {
    const initial = existingAssignmentSettings();
    const oldJson = serializeAdminWebPartSettings(initial);
    const properties = { adminSettings: oldJson, adminSettingsBackup: oldJson };
    const draft = { ...initial, assignedSources: initial.assignedSources.map(item => ({ ...item, isMandatory: true })) };
    expect(() => persistAdminWebPartSettings(properties, draft, () => { throw new Error('Notification rejected'); })).toThrow('Notification rejected');
    expect(properties).toEqual({ adminSettings: oldJson, adminSettingsBackup: oldJson });
  });
});
