import {
  defaultAdminWebPartSettings, defaultAllowedOverrides, defaultUserCalendarSettings,
  type IAdminWebPartSettings, type ICalendarSourceBase, type IExchangeMailboxDiscovery,
  type IUserCalendarSettings
} from '../models/ICalendarSettings';
import {
  serializeAdminWebPartSettings, applyAdminSourceChanges, cleanAdminSourceOverrides, deriveUserCalendarSettings,
  loadAdminWebPartSettings, normalizeUserCalendarSettings, resolveCalendarSettings
} from './CalendarSettingsService';

const matchedGroupIds = new Set(['target-group']);
const audienceGroups = [{ groupId: 'target-group', displayName: 'Target group' }];
const exchange: ICalendarSourceBase = {
  sourceType: 'exchange', exchangeMailbox: 'mailbox-id', exchangeCalendarId: 'calendar-id',
  name: 'Assigned calendar', color: '#0078d4', isEnabled: true
};
const sharePoint: ICalendarSourceBase = {
  sourceType: 'sharepoint', sharePointSiteId: 'site-id', sharePointListId: 'list-id',
  name: 'Assigned calendar', color: '#0078d4', isEnabled: true
};

function reloadAdmin(settings: IAdminWebPartSettings): IAdminWebPartSettings {
  const serialized = serializeAdminWebPartSettings(settings);
  const loaded = loadAdminWebPartSettings({ current: serialized });
  expect(loaded.source).toBe('current');
  return loaded.settings;
}

function reloadUser(settings: IUserCalendarSettings): IUserCalendarSettings {
  const loaded = normalizeUserCalendarSettings(JSON.parse(JSON.stringify(settings)));
  if (!loaded) throw new Error('Expected valid persisted personal settings');
  return loaded;
}

describe('previously removed calendars becoming mandatory', () => {
  it.each([exchange, sharePoint])('restores a removed $sourceType calendar across saved policy and personal settings', source => {
    const optional = reloadAdmin({
      ...defaultAdminWebPartSettings,
      assignedSources: [{
        assignmentId: 'assignment-id', adminSourceId: 'source-id', source, audienceGroups,
        isMandatory: false, defaultEnabled: true, allowedOverrides: { ...defaultAllowedOverrides }
      }]
    });
    const initial = resolveCalendarSettings({ adminSettings: optional, userSettings: defaultUserCalendarSettings, matchedGroupIds });
    expect(initial.sources).toHaveLength(1);
    initial.sources[0] = applyAdminSourceChanges(initial.sources[0], { name: 'Personal name' });
    const personalized = deriveUserCalendarSettings({ nextResolvedSettings: initial, adminSettings: optional, matchedGroupIds });
    initial.sources = [];
    const removed = reloadUser(deriveUserCalendarSettings({
      nextResolvedSettings: initial, adminSettings: optional, matchedGroupIds, existingUserSettings: personalized
    }));
    expect(removed.adminSourceOverridesById['source-id']).toEqual({ removed: true, name: 'Personal name' });
    expect(resolveCalendarSettings({ adminSettings: optional, userSettings: removed, matchedGroupIds }).sources).toEqual([]);

    const mandatory = reloadAdmin({
      ...optional, assignedSources: optional.assignedSources.map(item => ({ ...item, isMandatory: true }))
    });
    // Reload the old personal file: restoration must not require a personal save or reset.
    const restored = resolveCalendarSettings({ adminSettings: mandatory, userSettings: removed, matchedGroupIds });
    expect(restored.sources).toHaveLength(1);
    expect(restored.sources[0]).toMatchObject({ id: 'source-id', isMandatory: true, isEnabled: true, name: 'Personal name' });
    expect(restored.sources[0].visibilityOverride).toBeUndefined();

    // Composition cleans observed disallowed choices; a toolbar save also persists that cleanup.
    const cleaned = cleanAdminSourceOverrides(removed, restored.applicableAdminSources || []);
    const toolbarSaved = reloadUser({ ...cleaned, defaultView: 'week' });
    expect(toolbarSaved.adminSourceOverridesById['source-id']).toEqual({ name: 'Personal name' });
    const saved = reloadUser(deriveUserCalendarSettings({
      nextResolvedSettings: restored, adminSettings: mandatory, matchedGroupIds, existingUserSettings: cleaned
    }));
    expect(saved.adminSourceOverridesById['source-id']).toEqual({ name: 'Personal name' });
    const relaxed = resolveCalendarSettings({ adminSettings: optional, userSettings: saved, matchedGroupIds });
    expect(relaxed.sources[0]).toMatchObject({ isEnabled: true, name: 'Personal name' });
  });

  it.each([2, 3, 4, 5, 6, 7])('ignores saved removal and disabled state from personal schema %s', schemaVersion => {
    const mandatory = reloadAdmin({ ...defaultAdminWebPartSettings, assignedSources: [{
      adminSourceId: 'source-id', source: exchange, audienceGroups, isMandatory: true, defaultEnabled: true
    }] });
    const user = reloadUser({ ...defaultUserCalendarSettings, schemaVersion,
      adminSourceOverridesById: { 'source-id': { removed: true, isEnabled: false } },
      exchangeCalendarStates: { 'calendar-id': false }
    });
    const restored = resolveCalendarSettings({ adminSettings: mandatory, userSettings: user, matchedGroupIds, currentUserMailboxId: 'mailbox-id' });
    expect(restored.sources).toHaveLength(1);
    expect(restored.sources[0]).toMatchObject({ isMandatory: true, isEnabled: true });
    expect(cleanAdminSourceOverrides(user, restored.applicableAdminSources || []).adminSourceOverridesById).toEqual({});
  });

  it.each([false, true])('restores a removed automatic mailbox calendar with a calendar exception=%s', useException => {
    const discoveries: IExchangeMailboxDiscovery[] = [{ mailboxId: 'mailbox-id', sources: [exchange] }];
    const optional = reloadAdmin({ ...defaultAdminWebPartSettings, exchangeMailboxAssignments: [{
      assignmentId: 'rule-id', mailboxId: 'mailbox-id', mailboxDisplayName: 'Mailbox', audienceGroups,
      isMandatory: false, defaultEnabled: true, allowedOverrides: { ...defaultAllowedOverrides }, exceptions: []
    }] });
    const initial = resolveCalendarSettings({ adminSettings: optional, userSettings: defaultUserCalendarSettings, matchedGroupIds, mailboxDiscoveries: discoveries });
    const sourceId = initial.sources[0].id;
    initial.sources = [];
    const removed = reloadUser(deriveUserCalendarSettings({ nextResolvedSettings: initial, adminSettings: optional, matchedGroupIds, mailboxDiscoveries: discoveries }));
    expect(removed.adminSourceOverridesById[sourceId]).toEqual({ removed: true });
    const mandatory = reloadAdmin({ ...optional, exchangeMailboxAssignments: optional.exchangeMailboxAssignments?.map(rule => ({
      ...rule, isMandatory: !useException,
      exceptions: useException ? [{ calendarId: 'calendar-id', isMandatory: true, defaultEnabled: true }] : []
    })) });
    const restored = resolveCalendarSettings({ adminSettings: mandatory, userSettings: removed, matchedGroupIds, mailboxDiscoveries: discoveries });
    expect(restored.sources).toHaveLength(1);
    expect(restored.sources[0]).toMatchObject({ id: sourceId, isMandatory: true, isEnabled: true });
    const saved = reloadUser(deriveUserCalendarSettings({ nextResolvedSettings: restored, adminSettings: mandatory, matchedGroupIds, existingUserSettings: removed, mailboxDiscoveries: discoveries }));
    expect(saved.adminSourceOverridesById).toEqual({});
  });
});
