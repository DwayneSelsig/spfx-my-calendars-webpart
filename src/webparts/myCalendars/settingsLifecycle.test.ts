jest.mock('@microsoft/sp-core-library', () => ({ Version: { parse: jest.fn() } }));
jest.mock('@microsoft/sp-webpart-base', () => ({ BaseClientSideWebPart: class { public properties = {}; public context = { pageContext: { aadInfo: {} } }; } }));
jest.mock('./components/MyCalendars', () => ({ __esModule: true, default: jest.fn() }));
jest.mock('./propertyPane/PropertyPaneAdminDefaultsManager', () => ({}));
jest.mock('./services/ExchangeCalendarService', () => ({}));
jest.mock('./services/AudienceService', () => ({}));
jest.mock('./services/SettingsStorageService', () => ({}));

import MyCalendarsWebPart from './MyCalendarsWebPart';
import { defaultAdminWebPartSettings, defaultCalendarSettings, defaultUserCalendarSettings, type IAdminWebPartSettings, type ICalendarSettings, type IUserCalendarSettings } from './models/ICalendarSettings';
import { loadAdminWebPartSettings, normalizeUserCalendarSettings, serializeAdminWebPartSettings } from './services/CalendarSettingsService';

interface Harness {
  properties: { adminSettings?: string };
  _adminSettings: IAdminWebPartSettings;
  _adminLoadNotice?: string;
  _userSettings: IUserCalendarSettings;
  _audienceService: { getMatchingGroupIds: jest.Mock };
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

  it('restores a removed Exchange assignment through administrator Save and reload without a personal reset', async () => {
    const rebuild = (MyCalendarsWebPart.prototype as unknown as { rebuildResolvedSettings: () => Promise<void> }).rebuildResolvedSettings;
    const audience = { getMatchingGroupIds: jest.fn(async () => new Set(['target-group'])) };
    webpart._audienceService = audience;
    webpart.rebuildResolvedSettings = jest.fn(rebuild.bind(webpart));
    const optional = loadAdminWebPartSettings({ current: serializeAdminWebPartSettings({
      ...structuredClone(defaultAdminWebPartSettings),
      assignedSources: [{
        assignmentId: 'assignment-id', adminSourceId: 'source-id',
        source: { sourceType: 'exchange', exchangeMailbox: 'mailbox-id', exchangeCalendarId: 'calendar-id',
          name: 'Assigned calendar', color: '#0078d4', isEnabled: true },
        audienceGroups: [{ groupId: 'target-group', displayName: 'Target group' }],
        isMandatory: false, defaultEnabled: true
      }]
    }) }).settings;
    webpart._adminSettings = optional;
    await webpart.rebuildResolvedSettings();
    expect(webpart._resolvedSettings.sources).toHaveLength(1);

    await webpart.handleUserSettingsChange({ ...webpart._resolvedSettings, sources: [] });
    const removedFile = JSON.stringify(webpart._storageService.saveUserSettings.mock.calls[0][0]);
    expect(JSON.parse(removedFile).adminSourceOverridesById).toEqual({ 'source-id': { removed: true } });
    expect(webpart._resolvedSettings.sources).toEqual([]);
    webpart._storageService.saveUserSettings.mockClear();

    await webpart.handleAdminSettingsSave({ ...optional,
      assignedSources: optional.assignedSources.map(item => ({ ...item, isMandatory: true }))
    }, json => { webpart.properties.adminSettings = json; });
    expect(webpart._resolvedSettings.sources).toHaveLength(1);
    expect(webpart._resolvedSettings.sources[0]).toMatchObject({ id: 'source-id', isMandatory: true, isEnabled: true });
    expect(webpart._userSettings.adminSourceOverridesById).toEqual({});
    expect(webpart._storageService.saveUserSettings).not.toHaveBeenCalled();

    // Start another instance with the accepted administrator JSON and the old OneDrive file.
    const reloaded = new MyCalendarsWebPart() as unknown as Harness;
    reloaded.render = jest.fn();
    reloaded._adminSettings = loadAdminWebPartSettings({ current: webpart.properties.adminSettings }).settings;
    reloaded._userSettings = normalizeUserCalendarSettings(JSON.parse(removedFile))!;
    reloaded._audienceService = audience;
    reloaded._storageService = webpart._storageService;
    await rebuild.call(reloaded);
    expect(reloaded._resolvedSettings.sources).toHaveLength(1);
    expect(reloaded._resolvedSettings.sources[0]).toMatchObject({ id: 'source-id', isMandatory: true, isEnabled: true });
    expect(reloaded._storageService.saveUserSettings).not.toHaveBeenCalled();

    await reloaded.handleDefaultViewChange('week');
    expect(reloaded._storageService.saveUserSettings).toHaveBeenCalledTimes(1);
    expect(reloaded._storageService.saveUserSettings.mock.calls[0][0]).toMatchObject({
      defaultView: 'week', adminSourceOverridesById: {}
    });
  });

  it('rejects a callback that does not transfer the property instead of accepting an unsaved configuration', async () => {
    const commit = jest.fn();
    await expect(webpart.handleAdminSettingsSave({ ...defaultAdminWebPartSettings, defaultView: 'day' }, commit)).rejects.toThrow('SPFx did not accept');
    expect(commit).toHaveBeenCalledTimes(1);
    expect(webpart.properties).toEqual({});
    expect(webpart._adminSettings.defaultView).toBe('month');
    expect(webpart.rebuildResolvedSettings).not.toHaveBeenCalled();
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
    const commit = jest.fn((json: string | undefined) => { webpart.properties.adminSettings = json; });
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

  it('detects a mismatched transfer and a silently ignored restoration through the same callback', async () => {
    const previous = serializeAdminWebPartSettings(defaultAdminWebPartSettings);
    webpart.properties.adminSettings = previous;
    const commit = jest.fn((json: string | undefined) => {
      if (json !== previous) webpart.properties.adminSettings = 'unexpected host value';
    });
    await expect(webpart.handleAdminSettingsSave({ ...defaultAdminWebPartSettings, defaultView: 'day' }, commit)).rejects.toThrow('SPFx did not accept');
    expect(commit).toHaveBeenCalledTimes(2);
    expect(commit).toHaveBeenLastCalledWith(previous);
    expect(webpart._adminSettings.defaultView).toBe('month');
    expect(webpart._adminLoadNotice).toContain('could not be restored');
    expect(webpart.rebuildResolvedSettings).not.toHaveBeenCalled();
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
