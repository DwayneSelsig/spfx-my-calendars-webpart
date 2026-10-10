jest.mock('@microsoft/sp-core-library', () => ({ Version: { parse: jest.fn() } }));
jest.mock('@microsoft/sp-webpart-base', () => ({ BaseClientSideWebPart: class { public properties = {}; public context = { pageContext: { aadInfo: {} } }; } }));
jest.mock('./components/MyCalendars', () => ({ __esModule: true, default: jest.fn() }));
jest.mock('./propertyPane/PropertyPaneAdminDefaultsManager', () => ({}));
jest.mock('./services/ExchangeCalendarService', () => ({}));
jest.mock('./services/AudienceService', () => ({}));
jest.mock('./services/SettingsStorageService', () => ({}));

import MyCalendarsWebPart from './MyCalendarsWebPart';
import { defaultAdminWebPartSettings, defaultCalendarSettings, defaultUserCalendarSettings, type IAdminWebPartSettings, type ICalendarSettings, type IUserCalendarSettings } from './models/ICalendarSettings';
import { loadAdminWebPartSettings, serializeAdminWebPartSettings } from './services/CalendarSettingsService';

interface Harness {
  properties: { adminSettings?: string };
  _adminSettings: IAdminWebPartSettings;
  _adminLoadNotice?: string;
  _userSettings: IUserCalendarSettings;
  _resolvedSettings: ICalendarSettings;
  _storageService: { saveUserSettings: jest.Mock; deleteUserSettings: jest.Mock };
  _userSavePending: boolean;
  _disposed: boolean;
  render: jest.Mock;
  rebuildResolvedSettings: jest.Mock;
  handleAdminSettingsSave: (draft: IAdminWebPartSettings, commit: (json: string | undefined) => void) => Promise<void>;
  handleUserSettingsPreview: (draft: ICalendarSettings) => void;
  handleCancelUserSettingsPreview: () => void;
  handleUserSettingsChange: (draft: ICalendarSettings) => Promise<void>;
  handleDefaultViewChange: (view: 'day' | 'week' | 'month') => Promise<void>;
  handleResetUserSettings: () => Promise<void>;
  onAfterDeserialize: (properties: Record<string, unknown>) => Record<string, unknown>;
}

describe('settings composition ownership', () => {
  let webpart: Harness;
  beforeAll(() => {
    if (!globalThis.structuredClone) Object.defineProperty(globalThis, 'structuredClone', { value: <T,>(value: T): T => (jest.requireActual('v8').deserialize(jest.requireActual('v8').serialize(value)) as T), configurable: true });
  });
  beforeEach(() => {
    webpart = new MyCalendarsWebPart() as unknown as Harness;
    webpart.render = jest.fn();
    webpart.rebuildResolvedSettings = jest.fn(async () => undefined);
    webpart._storageService = { saveUserSettings: jest.fn(async () => true), deleteUserSettings: jest.fn(async () => true) };
    jest.spyOn(console, 'error').mockImplementation(() => undefined);
  });
  afterEach(() => jest.restoreAllMocks());

  it('hands the full administrator configuration to exactly one SPFx writer', async () => {
    const draft = { ...structuredClone(defaultAdminWebPartSettings), showWeekends: false, organizationPrimaryColor: '#ff0000', plannerShowAllAssignedToMeOnly: true, teamsShiftsShowAllCalendars: false };
    const commit = jest.fn((json: string | undefined) => { webpart.properties.adminSettings = json; });
    await webpart.handleAdminSettingsSave(draft, commit);
    expect(commit).toHaveBeenCalledTimes(1);
    expect(commit).toHaveBeenCalledWith(serializeAdminWebPartSettings(draft));
    expect(loadAdminWebPartSettings({ current: webpart.properties.adminSettings }).settings).toEqual(draft);
    expect(Object.keys(webpart.properties)).toEqual(['adminSettings']);
  });

  it('does not perform an additional direct property write', async () => {
    const commit = jest.fn();
    await webpart.handleAdminSettingsSave(defaultAdminWebPartSettings, commit);
    expect(commit).toHaveBeenCalledTimes(1);
    expect(webpart.properties).toEqual({});
  });

  it('keeps accepted settings on transfer failure and restores a partial write through the same route', async () => {
    const previous = serializeAdminWebPartSettings(defaultAdminWebPartSettings);
    webpart.properties.adminSettings = previous;
    const commit = jest.fn((json: string | undefined) => {
      webpart.properties.adminSettings = json;
      if (json !== previous) throw new Error('Host rejected');
    });
    await expect(webpart.handleAdminSettingsSave({ ...defaultAdminWebPartSettings, defaultView: 'week' }, commit)).rejects.toThrow('Host rejected');
    expect(commit).toHaveBeenLastCalledWith(previous);
    expect(webpart._adminSettings.defaultView).toBe('month');
    expect(webpart.rebuildResolvedSettings).not.toHaveBeenCalled();
  });

  it('rejects concurrent administrator saves and classifies runtime failure after transfer separately', async () => {
    let finish!: () => void;
    webpart.rebuildResolvedSettings.mockImplementationOnce(() => new Promise<void>(resolve => { finish = resolve; }));
    const commit = jest.fn();
    const pending = webpart.handleAdminSettingsSave(defaultAdminWebPartSettings, commit);
    await expect(webpart.handleAdminSettingsSave(defaultAdminWebPartSettings, commit)).rejects.toThrow('already in progress');
    expect(commit).toHaveBeenCalledTimes(1);
    finish();
    await pending;
    webpart.rebuildResolvedSettings.mockRejectedValueOnce(new Error('Graph failed'));
    await expect(webpart.handleAdminSettingsSave({ ...defaultAdminWebPartSettings, defaultView: 'day' }, commit)).resolves.toBeUndefined();
    expect(webpart._adminSettings.defaultView).toBe('day');
    expect(console.error).toHaveBeenCalledWith(expect.stringContaining('were saved'), expect.any(Error));
  });

  it('projects only supported webpart properties without changing administrator JSON', () => {
    const input = { settings: 'legacy', adminSettings: 'current', obsoleteProperty: 'discard' };
    expect(webpart.onAfterDeserialize(input)).toEqual({ settings: 'legacy', adminSettings: 'current' });
    expect(input.obsoleteProperty).toBe('discard');
  });

  it('reports property restoration failure separately from the rejected transfer', async () => {
    webpart.properties.adminSettings = 'previous';
    const commit = jest.fn((json: string | undefined) => {
      if (json === 'previous') throw new Error('Restore rejected');
      webpart.properties.adminSettings = json;
      throw new Error('Transfer rejected');
    });
    await expect(webpart.handleAdminSettingsSave(defaultAdminWebPartSettings, commit)).rejects.toThrow('Transfer rejected');
    expect(commit).toHaveBeenCalledTimes(2);
    expect(webpart._adminLoadNotice).toContain('could not be restored');
    expect(console.error).toHaveBeenCalledWith(expect.stringContaining('Could not restore'), expect.any(Error));
  });

  it('previews and cancels personal settings without any storage or administrator writes', () => {
    const draft = { ...structuredClone(defaultCalendarSettings), showWeekends: false, userShowWeekends: false };
    webpart.handleUserSettingsPreview(draft);
    expect(webpart._resolvedSettings.showWeekends).toBe(false);
    expect(webpart._userSettings).toEqual(defaultUserCalendarSettings);
    expect(webpart._storageService?.saveUserSettings).not.toHaveBeenCalled();
    expect(webpart.properties).toEqual({});
    webpart.handleCancelUserSettingsPreview();
    expect(webpart._resolvedSettings.showWeekends).toBe(true);
  });

  it('awaits personal Save, blocks competing writes and accepts only the confirmed result', async () => {
    let finish!: (success: boolean) => void;
    webpart._storageService?.saveUserSettings.mockImplementationOnce(() => new Promise<boolean>(resolve => { finish = resolve; }));
    const pending = webpart.handleUserSettingsChange({ ...defaultCalendarSettings, showWeekends: false, userShowWeekends: false });
    expect(webpart._resolvedSettings.showWeekends).toBe(false);
    expect(webpart._userSettings.userShowWeekends).toBeUndefined();
    expect(webpart._userSavePending).toBe(true);
    await expect(webpart.handleResetUserSettings()).rejects.toThrow('already in progress');
    finish(true);
    await pending;
    expect(webpart._userSettings.userShowWeekends).toBe(false);
    expect(webpart._userSavePending).toBe(false);
    expect(webpart.properties).toEqual({});
  });

  it('rolls a failed optimistic toolbar preference back to the confirmed preference', async () => {
    webpart._userSettings = { ...defaultUserCalendarSettings, defaultView: 'day' };
    webpart._storageService?.saveUserSettings.mockResolvedValueOnce(false);
    await expect(webpart.handleDefaultViewChange('week')).rejects.toThrow('Failed to save');
    expect(webpart._resolvedSettings.defaultView).toBe('day');
    expect(webpart._userSettings.defaultView).toBe('day');
    expect(webpart._userSavePending).toBe(false);
  });

  it('restores confirmed personal settings after Save or Reset failure', async () => {
    webpart._userSettings = { ...defaultUserCalendarSettings, userShowWeekends: false };
    webpart._storageService?.saveUserSettings.mockResolvedValueOnce(false);
    await expect(webpart.handleUserSettingsChange(defaultCalendarSettings)).rejects.toThrow('Failed to save');
    expect(webpart._resolvedSettings.showWeekends).toBe(false);
    webpart._storageService?.deleteUserSettings.mockResolvedValueOnce(false);
    await expect(webpart.handleResetUserSettings()).rejects.toThrow('Failed to reset');
    expect(webpart._resolvedSettings.showWeekends).toBe(false);
    await webpart.handleResetUserSettings();
    expect(webpart._userSettings).toEqual(defaultUserCalendarSettings);
    expect(webpart._resolvedSettings.showWeekends).toBe(true);
  });

  it('does not report a successful personal write as failed when the final runtime render fails', async () => {
    webpart.render.mockImplementationOnce(() => undefined).mockImplementationOnce(() => { throw new Error('Render failed'); });
    await expect(webpart.handleDefaultViewChange('week')).resolves.toBeUndefined();
    expect(webpart._userSettings.defaultView).toBe('week');
    expect(console.error).toHaveBeenCalledWith(expect.stringContaining('storage completed'), expect.any(Error));
  });
});
