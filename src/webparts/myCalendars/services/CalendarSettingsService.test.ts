import { CALENDAR_SETTINGS_SCHEMA_VERSION, defaultAdminWebPartSettings, defaultUserCalendarSettings } from '../models/ICalendarSettings';
import {
  loadAdminWebPartSettings,
  migrateLegacyUserSettings,
  normalizeAdminWebPartSettings,
  normalizeUserCalendarSettings,
  resolveCalendarSettings,
  serializeAdminWebPartSettings
} from './CalendarSettingsService';

describe('administrator settings normalization and loading', () => {
  it('round-trips the same ICS URL for Everyone and distinct group audiences', () => {
    const url = 'https://example.com/calendar.ics';
    const catalog = [
      { adminIcsId: 'all', displayName: 'All holidays', icsUrl: url, audienceGroups: [] },
      { adminIcsId: 'mt', displayName: 'MT holidays', icsUrl: url, audienceGroups: [{ groupId: 'mt', displayName: 'MT' }] },
      { adminIcsId: 'hr', displayName: 'HR holidays', icsUrl: url, audienceGroups: [{ groupId: 'hr', displayName: 'HR' }] }
    ];
    const normalized = normalizeAdminWebPartSettings({ ...defaultAdminWebPartSettings, icsCatalog: catalog });
    expect(normalized?.icsCatalog).toEqual(catalog);
    const loaded = loadAdminWebPartSettings({ current: serializeAdminWebPartSettings(normalized!) });
    expect(loaded.source).toBe('current');
    expect(loaded.settings.icsCatalog).toEqual(catalog);
    expect(resolveCalendarSettings({ adminSettings: loaded.settings, userSettings: defaultUserCalendarSettings, matchedGroupIds: new Set(['mt']) }).availableAdminIcsCatalogItems.map(item => item.adminIcsId)).toEqual(['all', 'mt']);
    expect(resolveCalendarSettings({ adminSettings: loaded.settings, userSettings: defaultUserCalendarSettings, matchedGroupIds: new Set(['mt', 'hr']) }).availableAdminIcsCatalogItems.map(item => item.adminIcsId)).toEqual(['all', 'mt', 'hr']);
  });

  it.each([{ audienceGroups: [] }, { audienceGroups: [{ groupId: 'a', displayName: 'A' }, { groupId: 'b', displayName: 'B' }] }])('rejects the same ICS URL and audience set regardless of group order', ({ audienceGroups }) => {
    expect(normalizeAdminWebPartSettings({ ...defaultAdminWebPartSettings, icsCatalog: [
      { adminIcsId: 'one', displayName: 'One', icsUrl: 'https://example.com/feed', audienceGroups },
      { adminIcsId: 'two', displayName: 'Two', icsUrl: ' HTTPS://EXAMPLE.COM/FEED ', audienceGroups: [...audienceGroups].reverse() }
    ] })).toBeUndefined();
  });
  it('preserves cache and other administrator fields during normalization', () => {
    const normalized = normalizeAdminWebPartSettings({
      ...defaultAdminWebPartSettings,
      defaultView: 'day',
      showWeekends: false,
      preferredStartMinutes: 7 * 60,
      visibleHourCount: 12,
      slotDurationMinutes: 15,
      enableCache: false,
      cacheDurationMinutes: 43,
      exchangeShowSourceLogo: false,
      plannerShowAllAssignedToMeOnly: true
    });

    expect(normalized).toMatchObject({
      defaultView: 'day',
      showWeekends: false,
      preferredStartMinutes: 7 * 60,
      visibleHourCount: 12,
      slotDurationMinutes: 15,
      enableCache: false,
      cacheDurationMinutes: 43,
      exchangeShowSourceLogo: false,
      plannerShowAllAssignedToMeOnly: true
    });
  });

  it('round-trips persisted administrator settings without returning cache duration to ten', () => {
    const persisted = {
      ...defaultAdminWebPartSettings,
      defaultView: 'week' as const,
      showWeekends: false,
      cacheDurationMinutes: 27,
      unifiedGroupShowAllCalendars: false
    };

    const loaded = loadAdminWebPartSettings({
      current: JSON.stringify(persisted)
    });

    expect(loaded.source).toBe('current');
    expect(loaded.settings).toMatchObject({
      defaultView: 'week',
      showWeekends: false,
      cacheDurationMinutes: 27,
      unifiedGroupShowAllCalendars: false
    });
  });

  it('migrates missing cache fields to the documented defaults', () => {
    const legacyShape = { ...defaultAdminWebPartSettings } as Record<string, unknown>;
    delete legacyShape.enableCache;
    delete legacyShape.cacheDurationMinutes;

    const normalized = normalizeAdminWebPartSettings(legacyShape);

    expect(normalized?.enableCache).toBe(true);
    expect(normalized?.cacheDurationMinutes).toBe(10);
  });

  it('migrates schema 5 SharePoint site metadata into schema 6 and keeps missing names valid', () => {
    const admin = normalizeAdminWebPartSettings({
      ...defaultAdminWebPartSettings,
      schemaVersion: 5,
      assignedSources: [{
        adminSourceId: 'admin-sp', audienceGroups: [{ groupId: 'group', displayName: 'Group' }],
        source: { sourceType: 'sharepoint', name: 'Events', color: '#0078d4', isEnabled: true, sharePointSiteId: 'site', sharePointSiteName: '  Contoso  ', sharePointListId: 'list' }
      }]
    });
    const user = normalizeUserCalendarSettings({
      ...defaultUserCalendarSettings,
      schemaVersion: 5,
      personalSources: [{ userSourceId: 'user-sp', sourceType: 'sharepoint', name: 'Personal', color: '#0078d4', isEnabled: true, sharePointSiteId: 'site-2', sharePointListId: 'list-2' }]
    });
    if (!admin || !user) throw new Error('Expected settings to normalize.');
    expect(admin.schemaVersion).toBe(CALENDAR_SETTINGS_SCHEMA_VERSION);
    expect(admin.assignedSources[0].source.sharePointSiteName).toBe('Contoso');
    expect(user?.personalSources[0].sharePointSiteName).toBeUndefined();
    const resolved = resolveCalendarSettings({ adminSettings: admin, userSettings: user, matchedGroupIds: new Set(['group']) });
    expect(resolved.sources.map(source => source.sharePointSiteName)).toEqual(['Contoso', undefined]);
  });

  it.each([2, 3, 4, 5, 6, 7])('accepts administrator and personal schema version %s and normalizes it to version 7', schemaVersion => {
    const admin = normalizeAdminWebPartSettings({ ...defaultAdminWebPartSettings, schemaVersion });
    const user = normalizeUserCalendarSettings({ ...defaultUserCalendarSettings, schemaVersion });

    expect(admin?.schemaVersion).toBe(CALENDAR_SETTINGS_SCHEMA_VERSION);
    expect(user?.schemaVersion).toBe(CALENDAR_SETTINGS_SCHEMA_VERSION);
  });

  it('migrates version 2 hour fields into the current settings shape', () => {
    const admin = normalizeAdminWebPartSettings({
      ...defaultAdminWebPartSettings,
      schemaVersion: 2,
      startHour: 7,
      endHour: 19,
      slotDuration: 15,
      preferredStartMinutes: undefined,
      visibleHourCount: undefined,
      slotDurationMinutes: undefined
    });
    const user = normalizeUserCalendarSettings({
      ...defaultUserCalendarSettings,
      schemaVersion: 2,
      userStartHour: 9,
      userEndHour: 17,
      userPreferredStartMinutes: undefined,
      userVisibleHourCount: undefined
    });

    expect(admin).toMatchObject({ schemaVersion: 7, preferredStartMinutes: 420, visibleHourCount: 12, slotDurationMinutes: 15 });
    expect(user).toMatchObject({ schemaVersion: 7, userPreferredStartMinutes: 540, userVisibleHourCount: 8 });
  });

  it.each([undefined, 1, 1.5, 8, '6'])('rejects unsupported current schema version %p', schemaVersion => {
    expect(normalizeAdminWebPartSettings({ ...defaultAdminWebPartSettings, schemaVersion })).toBeUndefined();
    expect(normalizeUserCalendarSettings({ ...defaultUserCalendarSettings, schemaVersion })).toBeUndefined();
  });

  it('falls back from an unknown current administrator version to legacy and defaults', () => {
    const invalidCurrent = JSON.stringify({ ...defaultAdminWebPartSettings, schemaVersion: 8 });
    expect(loadAdminWebPartSettings({ current: invalidCurrent }).source).toBe('defaults');

    expect(loadAdminWebPartSettings({
      current: invalidCurrent,
      legacy: JSON.stringify({ sources: [], defaultView: 'week' })
    })).toMatchObject({ source: 'legacy', settings: { defaultView: 'week' } });
  });

  it('accepts unversioned personal data only through the explicit legacy migration path', () => {
    const legacy = { sources: [], defaultView: 'day' as const, userStartHour: 6, userEndHour: 18 };

    expect(normalizeUserCalendarSettings(legacy)).toBeUndefined();
    expect(migrateLegacyUserSettings(legacy)).toMatchObject({
      schemaVersion: CALENDAR_SETTINGS_SCHEMA_VERSION,
      defaultView: 'day',
      userPreferredStartMinutes: 360,
      userVisibleHourCount: 12
    });
  });
});
