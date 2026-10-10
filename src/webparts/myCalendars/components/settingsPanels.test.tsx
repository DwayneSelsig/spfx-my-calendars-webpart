/* eslint-disable @rushstack/pair-react-dom-render-unmount -- Every test unmounts its shared host in afterEach. */
jest.mock('@microsoft/sp-http', () => ({}));
jest.mock('../services/ExchangeCalendarService', () => ({ ExchangeCalendarService: jest.fn() }));
jest.mock('../services/SharePointCalendarService', () => ({ SharePointCalendarService: jest.fn() }));
jest.mock('../services/PlannerTaskService', () => ({ PlannerTaskService: jest.fn() }));
jest.mock('../services/UnifiedGroupCalendarService', () => ({ UnifiedGroupCalendarService: jest.fn() }));
jest.mock('../services/AudienceService', () => ({ AudienceService: jest.fn() }));
jest.mock('MyCalendarsWebPartStrings', () => new Proxy({ __esModule: true }, {
  get: (_target, key) => key === '__esModule' ? true : String(key)
}));

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import { setIconOptions } from '@fluentui/react/lib/Styling';
import { AdminDefaultsPanel, type IAdminDefaultsPanelProps } from './AdminDefaultsPanel';
import { UserSettingsPanel, type IUserSettingsPanelProps } from './UserSettingsPanel';
import { defaultAdminWebPartSettings, defaultCalendarSettings } from '../models/ICalendarSettings';

function deferred(): { promise: Promise<void>; resolve: () => void; reject: (error: Error) => void } {
  let complete!: () => void;
  let fail!: (error: Error) => void;
  const promise = new Promise<void>((resolve, reject) => { complete = resolve; fail = reject; });
  return { promise, resolve: complete, reject: fail };
}

describe('settings panels awaited lifecycle', () => {
  let host: HTMLDivElement;
  let errorLog: jest.SpyInstance;
  beforeAll(() => {
    setIconOptions({ disableWarnings: true });
    if (!globalThis.structuredClone) Object.defineProperty(globalThis, 'structuredClone', { value: <T,>(value: T): T => (jest.requireActual('v8').deserialize(jest.requireActual('v8').serialize(value)) as T), configurable: true });
  });
  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
    errorLog = jest.spyOn(console, 'error').mockImplementation(() => undefined);
    jest.spyOn(window, 'confirm').mockReturnValue(true);
  });
  afterEach(() => {
    act(() => { ReactDOM.unmountComponentAtNode(host); });
    host.remove();
    jest.restoreAllMocks();
  });
  function click(label: string): void {
    const button = Array.from(document.querySelectorAll('button')).find(item => item.textContent === label);
    if (!button) throw new Error('Missing button: ' + label);
    act(() => { Simulate.click(button); });
  }
  function admin(overrides: Partial<IAdminDefaultsPanelProps> = {}): { panel: AdminDefaultsPanel; props: IAdminDefaultsPanelProps } {
    const props = { isOpen: true, settings: structuredClone(defaultAdminWebPartSettings), onDismiss: jest.fn(), onSave: jest.fn(async () => undefined), ...overrides };
    let panel!: AdminDefaultsPanel;
    act(() => { ReactDOM.render(<AdminDefaultsPanel {...props} ref={instance => { if (instance) panel = instance; }} />, host); });
    return { panel, props };
  }
  function user(overrides: Partial<IUserSettingsPanelProps> = {}): { panel: UserSettingsPanel; props: IUserSettingsPanelProps } {
    const props = { isOpen: true, settings: structuredClone(defaultCalendarSettings), onDismiss: jest.fn(), onPreview: jest.fn(), onSave: jest.fn(async () => undefined), onReset: jest.fn(async () => undefined), ...overrides };
    let panel!: UserSettingsPanel;
    act(() => { ReactDOM.render(<UserSettingsPanel {...props} ref={instance => { if (instance) panel = instance; }} />, host); });
    return { panel, props };
  }

  it('blocks duplicate administrator Save and closes only after success', async () => {
    const pending = deferred();
    const { panel, props } = admin({ onSave: jest.fn(() => pending.promise) });
    click('SaveLabel');
    const save = panel as unknown as { handleSave: () => Promise<void>; handleDismiss: () => void };
    await save.handleSave();
    save.handleDismiss();
    expect(props.onSave).toHaveBeenCalledTimes(1);
    expect(props.onDismiss).not.toHaveBeenCalled();
    expect(panel.state.isSaving).toBe(true);
    expect(document.querySelector('fieldset')?.disabled).toBe(true);
    expect(Array.from(document.querySelectorAll('button')).find(item => item.textContent === 'SavingLabel')?.disabled).toBe(true);
    await act(async () => { pending.resolve(); await pending.promise; });
    expect(props.onDismiss).toHaveBeenCalledTimes(1);
    expect(panel.state.isSaving).toBe(false);
  });

  it('retains the complete administrator draft after failure and snapshots it independently', async () => {
    const pending = deferred();
    const settings = { ...structuredClone(defaultAdminWebPartSettings), organizationPrimaryColor: '#ff0000', plannerShowAllAssignedToMeOnly: true, teamsShiftsShowSourceLogo: false };
    const { panel, props } = admin({ settings, onSave: jest.fn(() => pending.promise) });
    act(() => { panel.setState(previous => ({ settings: { ...previous.settings, showWeekends: false } })); });
    click('SaveLabel');
    const snapshot = (props.onSave as jest.Mock).mock.calls[0][0];
    expect(snapshot).toEqual(panel.state.settings);
    expect(snapshot).not.toBe(panel.state.settings);
    expect(snapshot).toMatchObject({ organizationPrimaryColor: '#ff0000', plannerShowAllAssignedToMeOnly: true, teamsShiftsShowSourceLogo: false });
    await act(async () => { pending.reject(new Error('Rejected')); await pending.promise.catch(() => undefined); });
    expect(props.onDismiss).not.toHaveBeenCalled();
    expect(panel.state.settings.showWeekends).toBe(false);
    expect(panel.state.isSaving).toBe(false);
    expect(document.body.textContent).toContain('AdminSettingsSaveErrorLabel');
  });

  it('resets only the administrator draft and reloads latest settings on reopen', () => {
    const { panel, props } = admin({ settings: { ...structuredClone(defaultAdminWebPartSettings), defaultView: 'week' } });
    click('ResetDraftToDefaultsLabel');
    expect(panel.state.settings).toEqual(defaultAdminWebPartSettings);
    expect(props.onSave).not.toHaveBeenCalled();
    click('CancelLabel');
    expect(props.onSave).not.toHaveBeenCalled();
    act(() => { ReactDOM.render(<AdminDefaultsPanel {...props} isOpen={false} />, host); });
    const latest = { ...props.settings, defaultView: 'day' as const };
    act(() => { ReactDOM.render(<AdminDefaultsPanel {...props} settings={latest} isOpen />, host); });
    expect(panel.state.settings.defaultView).toBe('day');
    expect(panel.state.settings).not.toBe(latest);
  });

  it('keeps administrator edits when incoming props change while open', () => {
    const { panel, props } = admin();
    act(() => { panel.setState(previous => ({ settings: { ...previous.settings, showWeekends: false } })); });
    act(() => { ReactDOM.render(<AdminDefaultsPanel {...props} settings={{ ...props.settings, defaultView: 'day' }} />, host); });
    expect(panel.state.settings.showWeekends).toBe(false);
    expect(panel.state.settings.defaultView).toBe('month');
  });

  it('previews personal edits without saving and retains them on failed Save', async () => {
    const pending = deferred();
    const { panel, props } = user({ onSave: jest.fn(() => pending.promise) });
    act(() => { panel.setState(previous => ({ settings: { ...previous.settings, showWeekends: false, userShowWeekends: false } })); });
    expect(props.onPreview).toHaveBeenCalledWith(panel.state.settings);
    expect(props.onSave).not.toHaveBeenCalled();
    click('SaveLabel');
    expect(props.onDismiss).not.toHaveBeenCalled();
    await act(async () => { pending.reject(new Error('Storage unavailable')); await pending.promise.catch(() => undefined); });
    expect(panel.state.settings.showWeekends).toBe(false);
    expect(panel.state.isSaving).toBe(false);
    expect(props.onDismiss).not.toHaveBeenCalled();
    expect(document.body.textContent).toContain('UserSettingsSaveErrorLabel');
  });

  it('awaits personal Reset, preserves the draft on failure and permits retry', async () => {
    const pending = deferred();
    const { panel, props } = user({ onReset: jest.fn(() => pending.promise) });
    const draft = panel.state.settings;
    click('ResetToDefaultsLabel');
    expect(props.onDismiss).not.toHaveBeenCalled();
    await act(async () => { pending.reject(new Error('Delete rejected')); await pending.promise.catch(() => undefined); });
    expect(panel.state.settings).toBe(draft);
    expect(panel.state.isSaving).toBe(false);
    expect(document.body.textContent).toContain('UserSettingsResetErrorLabel');
  });

  it.each(['admin', 'user'])('blocks %s draft edits while Save is pending', async kind => {
    const pending = deferred();
    const { panel } = kind === 'admin' ? admin({ onSave: () => pending.promise }) : user({ onSave: () => pending.promise });
    click('SaveLabel');
    const originalDraft = panel.state.settings;
    act(() => {
      if (kind === 'admin') {
        (panel as unknown as { handleUpdateMailboxRule: (id: string, updates: { defaultEnabled: boolean }) => void }).handleUpdateMailboxRule('rule', { defaultEnabled: false });
      } else {
        (panel as unknown as { handleTogglePlannerShowAll: (checked: boolean) => void }).handleTogglePlannerShowAll(true);
      }
    });
    expect(panel.state.settings).toBe(originalDraft);
    await act(async () => { pending.resolve(); await pending.promise; });
  });

  it.each(['admin', 'user'])('ignores an awaited %s Save completion for a replaced edit session', async kind => {
    const pending = deferred();
    if (kind === 'admin') {
      const { props } = admin({ onSave: () => pending.promise });
      click('SaveLabel');
      act(() => { ReactDOM.render(<AdminDefaultsPanel {...props} isOpen={false} />, host); });
      act(() => { ReactDOM.render(<AdminDefaultsPanel {...props} isOpen />, host); });
      await act(async () => { pending.resolve(); await pending.promise; });
      expect(props.onDismiss).not.toHaveBeenCalled();
    } else {
      const { props } = user({ onSave: () => pending.promise });
      click('SaveLabel');
      act(() => { ReactDOM.render(<UserSettingsPanel {...props} isOpen={false} />, host); });
      act(() => { ReactDOM.render(<UserSettingsPanel {...props} isOpen />, host); });
      await act(async () => { pending.resolve(); await pending.promise; });
      expect(props.onDismiss).not.toHaveBeenCalled();
    }
  });

  it('ignores administrator group discovery after closing and reopening the draft', async () => {
    let complete!: (groups: Array<{ id: string; displayName: string }>) => void;
    const groups = new Promise<Array<{ id: string; displayName: string }>>(resolve => { complete = resolve; });
    const { panel, props } = admin();
    const loader = panel as unknown as {
      unifiedGroupService: { getUnifiedGroups: () => typeof groups; getJoinedTeamIds: () => Promise<Set<string>> };
      loadUnifiedGroups: () => Promise<void>;
    };
    loader.unifiedGroupService = { getUnifiedGroups: () => groups, getJoinedTeamIds: async () => new Set() };
    const loading = loader.loadUnifiedGroups();
    act(() => { ReactDOM.render(<AdminDefaultsPanel {...props} isOpen={false} />, host); });
    act(() => { ReactDOM.render(<AdminDefaultsPanel {...props} isOpen />, host); });
    await act(async () => { complete([{ id: 'old', displayName: 'Stale group' }]); await loading; });
    expect(panel.state.unifiedGroups).toEqual([]);
  });

  it('awaits personal Save, permits Cancel without writing and opens the latest accepted configuration', async () => {
    const pending = deferred();
    const { panel, props } = user({ onSave: jest.fn(() => pending.promise) });
    click('CancelLabel');
    expect(props.onSave).not.toHaveBeenCalled();
    act(() => { ReactDOM.render(<UserSettingsPanel {...props} isOpen={false} />, host); });
    act(() => { ReactDOM.render(<UserSettingsPanel {...props} settings={{ ...props.settings, showWeekends: false }} isOpen />, host); });
    expect(panel.state.settings.showWeekends).toBe(false);
    click('SaveLabel');
    expect(props.onDismiss).toHaveBeenCalledTimes(1);
    await act(async () => { pending.resolve(); await pending.promise; });
    expect(props.onDismiss).toHaveBeenCalledTimes(2);
  });

  it.each(['admin', 'user'])('ignores an awaited %s save completion after unmount', async kind => {
    const pending = deferred();
    const { panel, props } = kind === 'admin' ? admin({ onSave: () => pending.promise }) : user({ onSave: () => pending.promise });
    click('SaveLabel');
    act(() => { ReactDOM.unmountComponentAtNode(host); });
    const stateUpdate = jest.spyOn(panel, 'setState');
    await act(async () => { pending.resolve(); await pending.promise; });
    expect(stateUpdate).not.toHaveBeenCalled();
    expect(props.onDismiss).not.toHaveBeenCalled();
    expect(errorLog.mock.calls.some(call => String(call[0]).includes('unmounted'))).toBe(false);
  });
});
