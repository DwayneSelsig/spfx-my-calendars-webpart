import { defaultAdminWebPartSettings } from '../models/ICalendarSettings';
import {
  loadAdminWebPartSettings,
  normalizeAdminWebPartSettings
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
});
