jest.mock('@microsoft/sp-http', () => ({}));
jest.mock('../services/ExchangeCalendarService', () => ({}));
jest.mock('../services/SharePointCalendarService', () => ({}));
jest.mock('../services/PlannerTaskService', () => ({}));
jest.mock('../services/TeamsShiftsService', () => ({}));
jest.mock('../services/UnifiedGroupCalendarService', () => ({}));
jest.mock('./UserSettingsPanel', () => ({ UserSettingsPanel: () => null }));
jest.mock('./views/DayView', () => ({ DayView: () => null }));
jest.mock('./views/WeekView', () => ({ WeekView: () => null }));
jest.mock('./views/MonthView', () => ({ MonthView: () => null }));
jest.mock('./views/SearchResultsView', () => ({ SearchResultsView: () => null }));
jest.mock('./CalendarToolbar', () => ({ CalendarToolbar: () => null }));
jest.mock('MyCalendarsWebPartStrings', () => new Proxy({ __esModule: true }, { get: (_target, key) => key === '__esModule' ? true : String(key) }));

import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { act } from 'react';
import { setIconOptions } from '@fluentui/react/lib/Styling';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import MyCalendars from './MyCalendars';
import type { IMyCalendarsProps } from './IMyCalendarsProps';
import { defaultCalendarSettings } from '../models/ICalendarSettings';

describe('toolbar preference rollback against current policy', () => {
  let host: HTMLDivElement;
  let root: Root;
  beforeEach(() => {
    host = document.createElement('div'); root = createRoot(host); document.body.appendChild(host);
    setIconOptions({ disableWarnings: true });
    Object.defineProperty(globalThis, 'ResizeObserver', { configurable: true, value: class { public observe(): void { /* no layout in jsdom */ } public disconnect(): void { /* no layout in jsdom */ } } });
    const loading = MyCalendars.prototype as unknown as { loadAppointments: () => Promise<void>; ensureVisibleRange: () => Promise<void> };
    jest.spyOn(loading, 'loadAppointments').mockResolvedValue(undefined);
    jest.spyOn(loading, 'ensureVisibleRange').mockResolvedValue(undefined);
    jest.spyOn(console, 'error').mockImplementation(() => undefined);
  });
  afterEach(() => { act(() => { root.unmount(); }); host.remove(); jest.restoreAllMocks(); });

  it.each([false, true])('uses the newly resolved confirmed view on failure (search active: %s)', async searchActive => {
    let rejectWrite!: (error: Error) => void;
    const pending = new Promise<void>((_resolve, reject) => { rejectWrite = reject; });
    const props: IMyCalendarsProps = {
      description: '', isDarkTheme: false, environmentMessage: '', hasTeamsContext: false, userDisplayName: '', webPartInstanceId: 'test',
      settings: { ...defaultCalendarSettings, defaultView: 'month' }, isSettingsWritePending: false,
      onSettingsChange: async () => undefined, onPreviewSettings: () => undefined, onCancelSettingsPreview: () => undefined,
      onResetSettings: async () => undefined, onDefaultViewChange: () => pending,
      context: { msGraphClientFactory: { getClient: async () => undefined }, pageContext: { web: { absoluteUrl: 'https://example.com/sites/calendar' } } } as unknown as WebPartContext
    };
    let calendar!: MyCalendars;
    await act(async () => { root.render(<MyCalendars {...props} ref={instance => { if (instance) calendar = instance; }} />); });
    let saving!: Promise<void>;
    act(() => { saving = (calendar as unknown as { handleViewChange: (view: 'week') => Promise<void> }).handleViewChange('week'); });
    expect(calendar.state.currentView).toBe('week');
    // The owner rolls back against an administrator default changed during this write.
    act(() => { root.render(<MyCalendars {...props} settings={{ ...props.settings, defaultView: 'day' }} />); });
    if (searchActive) act(() => { calendar.setState({ currentView: 'search', searchQuery: 'meeting', appliedSearchQuery: 'meeting' }); });
    expect((calendar as unknown as { settingsUiMounted: boolean }).settingsUiMounted).toBe(true);
    await act(async () => { rejectWrite(new Error('Write rejected')); await saving; });
    expect(calendar.state.previousView).toBe('day');
    expect(calendar.state.currentView).toBe(searchActive ? 'search' : 'day');
    expect(calendar.state.settingsWriteError).toBe('UserSettingsSaveErrorLabel');
    expect(host.querySelector('[role="alert"]')).not.toBeNull();
  });
});
