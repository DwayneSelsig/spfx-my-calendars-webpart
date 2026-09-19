import { getBulkVisibilityTarget, getGroupVisibilityState, setOutlookVisibility, setSharePointVisibility } from './calendarVisibility';
import { defaultCalendarSettings, type ICalendarSettings } from '../models/ICalendarSettings';

describe('calendar group visibility', () => {
  it('derives on, off and mixed and makes mixed visible', () => {
    expect(getGroupVisibilityState([true, true])).toBe('on');
    expect(getGroupVisibilityState([false, false])).toBe('off');
    expect(getGroupVisibilityState([true, false])).toBe('mixed');
    expect(getBulkVisibilityTarget('on')).toBe(false);
    expect(getBulkVisibilityTarget('off')).toBe(true);
    expect(getBulkVisibilityTarget('mixed')).toBe(true);
  });

  it('updates Outlook discovered and configured calendars without touching discovery modes', () => {
    const settings: ICalendarSettings = {
      ...defaultCalendarSettings,
      plannerShowAllCalendars: true,
      sources: [
        { id: 'exchange-source', origin: 'user', sourceType: 'exchange', name: 'Shared', color: '#123456', isEnabled: false },
        { id: 'planner-source', origin: 'user', sourceType: 'planner', name: 'Plan', color: '#654321', isEnabled: false }
      ]
    };
    const updated = setOutlookVisibility(settings, ['calendar-a', 'calendar-b'], true);
    expect(updated.exchangeCalendarStates).toEqual({ 'calendar-a': true, 'calendar-b': true });
    expect(updated.sources[0].isEnabled).toBe(true);
    expect(updated.sources[1].isEnabled).toBe(false);
    expect(updated.plannerShowAllCalendars).toBe(true);
  });

  it('updates only SharePoint configured sources', () => {
    const settings: ICalendarSettings = {
      ...defaultCalendarSettings,
      sources: [
        { id: 'sp', origin: 'admin', adminSourceId: 'sp', sourceType: 'sharepoint', name: 'Events', color: '#123456', isEnabled: true },
        { id: 'exchange', origin: 'user', sourceType: 'exchange', name: 'Mailbox', color: '#654321', isEnabled: true }
      ]
    };
    const updated = setSharePointVisibility(settings, false);
    expect(updated.sources[0].isEnabled).toBe(false);
    expect(updated.sources[1].isEnabled).toBe(true);
  });
});
