import { defaultAdminWebPartSettings } from '../models/ICalendarSettings';
import {
  notifyAdminSettingsPropertyChanges,
  persistAdminWebPartSettings,
  type IAdminSettingsPropertyBag
} from './AdminSettingsPropertyPersistence';

describe('administrator settings property persistence', () => {
  it('writes the complete draft to current and backup before notifying SPFx', () => {
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
      expect(value).toBe(properties.adminSettings);
      propertiesAtNotification = { ...properties };
    });

    expect(properties.adminSettings).toBe(serialized);
    expect(properties.adminSettingsBackup).toBe(serialized);
    expect(propertiesAtNotification).toEqual(properties);
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
