import { CALENDAR_SETTINGS_SCHEMA_VERSION, defaultAdminWebPartSettings, defaultUserCalendarSettings } from '../models/ICalendarSettings';
import {
  loadAdminWebPartSettings,
  normalizeAdminWebPartSettings,
  normalizeUserCalendarSettings,
  resolveCalendarSettings
} from './CalendarSettingsService';

describe('administrator settings normalization and loading', () => {
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
      current: JSON.stringify(persisted),
      backup: JSON.stringify(defaultAdminWebPartSettings)
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
});
