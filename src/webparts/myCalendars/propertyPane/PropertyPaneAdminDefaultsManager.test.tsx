/* eslint-disable @rushstack/pair-react-dom-render-unmount -- afterEach unmounts the custom-field host. */
jest.mock('@microsoft/sp-property-pane', () => ({ PropertyPaneFieldType: { Custom: 1 } }));
jest.mock('../components/AdminDefaultsPanel', () => ({ AdminDefaultsPanel: jest.fn(() => null) }));
jest.mock('MyCalendarsWebPartStrings', () => new Proxy({ __esModule: true }, { get: (_target, key) => key === '__esModule' ? true : String(key) }));

import * as ReactDOM from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { AdminDefaultsPanel, type IAdminDefaultsPanelProps } from '../components/AdminDefaultsPanel';
import { defaultAdminWebPartSettings } from '../models/ICalendarSettings';
import { PropertyPaneAdminDefaultsManager } from './PropertyPaneAdminDefaultsManager';

describe('SPFx administrator defaults adapter', () => {
  let host: HTMLDivElement;
  beforeEach(() => { host = document.createElement('div'); document.body.appendChild(host); });
  afterEach(() => { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); jest.clearAllMocks(); });
  function panelProps(): IAdminDefaultsPanelProps {
    const calls = (AdminDefaultsPanel as unknown as jest.Mock).mock.calls;
    return calls[calls.length - 1][0];
  }
  function open(): void {
    act(() => { Simulate.click(host.querySelector('button') as HTMLButtonElement); });
  }

  it('provides one required framework writer, propagates errors and reads latest values without refresh', async () => {
    let settings = { ...defaultAdminWebPartSettings };
    const refresh = jest.fn();
    const context = { msGraphClientFactory: { getClient: async () => undefined }, pageContext: { cultureInfo: {} }, propertyPane: { refresh } } as unknown as WebPartContext;
    const field = PropertyPaneAdminDefaultsManager('adminSettings', {
      label: 'Defaults', context, getAdminSettings: () => settings, getAdminLoadNotice: () => undefined,
      onSave: async (draft, commit) => { commit('serialized'); settings = draft; }
    });
    const commit = jest.fn();
    await act(async () => { field.properties.onRender(host, undefined, commit); });
    open();
    expect(panelProps().isOpen).toBe(true);
    await act(async () => { await panelProps().onSave({ ...settings, defaultView: 'day' }); });
    expect(commit).toHaveBeenCalledTimes(1);
    expect(commit).toHaveBeenCalledWith('adminSettings', 'serialized', true);
    expect(panelProps().isOpen).toBe(true);
    act(() => { panelProps().onDismiss(); });
    open();
    expect(panelProps().settings.defaultView).toBe('day');
    expect(refresh).not.toHaveBeenCalled();
    commit.mockImplementationOnce(() => { throw new Error('Rejected'); });
    await act(async () => { await expect(panelProps().onSave(settings)).rejects.toThrow('Rejected'); });
    expect(panelProps().isOpen).toBe(true);
    expect(host.querySelector('button')?.disabled).toBe(false);
  });

  it('rejects missing SPFx callbacks instead of pretending a save succeeded', async () => {
    const onSave = jest.fn(async () => undefined);
    const context = { msGraphClientFactory: { getClient: async () => undefined }, pageContext: { cultureInfo: {} } } as unknown as WebPartContext;
    const field = PropertyPaneAdminDefaultsManager('adminSettings', { label: 'Defaults', context, getAdminSettings: () => defaultAdminWebPartSettings, getAdminLoadNotice: () => undefined, onSave });
    await act(async () => { field.properties.onRender(host); });
    open();
    await act(async () => { await expect(panelProps().onSave(defaultAdminWebPartSettings)).rejects.toThrow('unavailable'); });
    expect(onSave).not.toHaveBeenCalled();
    expect(panelProps().isOpen).toBe(true);
  });
});
