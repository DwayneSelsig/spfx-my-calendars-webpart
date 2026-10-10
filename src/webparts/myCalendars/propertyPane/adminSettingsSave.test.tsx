/* eslint-disable @rushstack/pair-react-dom-render-unmount -- afterEach disposes the custom field. */
jest.mock('@microsoft/sp-core-library', () => ({ Version: { parse: jest.fn() } }));
jest.mock('@microsoft/sp-webpart-base', () => ({ BaseClientSideWebPart: class { public properties = {}; } }));
jest.mock('@microsoft/sp-property-pane', () => ({ PropertyPaneFieldType: { Custom: 1 } }));
jest.mock('@microsoft/sp-http', () => ({}));
jest.mock('../components/MyCalendars', () => ({ __esModule: true, default: jest.fn() }));
jest.mock('../services/ExchangeCalendarService', () => ({ ExchangeCalendarService: jest.fn(() => ({ setGraphClient: jest.fn(),
  resolveMailbox: async () => ({ id: 'new-mailbox' }), getCalendars: async () => [{ id: 'new-calendar', name: 'New Exchange calendar', hexColor: '#0078d4' }] })) }));
jest.mock('../services/SharePointCalendarService', () => ({ SharePointCalendarService: jest.fn(() => ({ setGraphClient: jest.fn() })) }));
jest.mock('../services/PlannerTaskService', () => ({ PlannerTaskService: jest.fn(() => ({ setGraphClient: jest.fn() })) }));
jest.mock('../services/UnifiedGroupCalendarService', () => ({ UnifiedGroupCalendarService: jest.fn(() => ({ setGraphClient: jest.fn() })) }));
jest.mock('../services/AudienceService', () => ({}));
jest.mock('../services/SettingsStorageService', () => ({}));
jest.mock('MyCalendarsWebPartStrings', () => new Proxy({ __esModule: true }, { get: (_target, key) => key === '__esModule' ? true : String(key) }));

import { act, Simulate } from 'react-dom/test-utils';
import { setIconOptions } from '@fluentui/react/lib/Styling';
import type { IPropertyPaneCustomFieldProps, IPropertyPaneField } from '@microsoft/sp-property-pane';
import MyCalendarsWebPart from '../MyCalendarsWebPart';
import { defaultAdminWebPartSettings, type IAdminWebPartSettings } from '../models/ICalendarSettings';
import { loadAdminWebPartSettings, serializeAdminWebPartSettings } from '../services/CalendarSettingsService';

interface Harness {
  properties: { adminSettings?: string };
  context: unknown;
  _adminSettings: IAdminWebPartSettings;
  render: jest.Mock;
  rebuildResolvedSettings: jest.Mock;
  getPropertyPaneConfiguration: () => { pages: Array<{ groups: Array<{ groupFields: Array<IPropertyPaneField<IPropertyPaneCustomFieldProps>> }> }> };
}

describe('administrator editor to SPFx property hand-off', () => {
  let host: HTMLDivElement;
  let field: IPropertyPaneField<IPropertyPaneCustomFieldProps>;
  let webpart: Harness;
  beforeAll(() => {
    setIconOptions({ disableWarnings: true });
    if (!globalThis.structuredClone) Object.defineProperty(globalThis, 'structuredClone', { value: <T,>(value: T): T => (jest.requireActual('v8').deserialize(jest.requireActual('v8').serialize(value)) as T), configurable: true });
  });
  beforeEach(() => {
    host = document.createElement('div'); document.body.appendChild(host);
    webpart = new MyCalendarsWebPart() as unknown as Harness;
    webpart.context = { httpClient: {}, msGraphClientFactory: { getClient: async () => undefined }, pageContext: { cultureInfo: {} } };
    webpart.render = jest.fn();
    webpart.rebuildResolvedSettings = jest.fn(async () => undefined);
    webpart.properties.adminSettings = serializeAdminWebPartSettings({ ...structuredClone(defaultAdminWebPartSettings),
      assignedSources: [{ assignmentId: 'assignment', adminSourceId: 'source', source: { sourceType: 'exchange',
        exchangeMailbox: 'mailbox', exchangeCalendarId: 'calendar', name: 'Original calendar', color: '#0078d4', isEnabled: true },
        audienceGroups: [], isMandatory: false, defaultEnabled: true }] });
    webpart._adminSettings = loadAdminWebPartSettings({ current: webpart.properties.adminSettings }).settings;
    field = webpart.getPropertyPaneConfiguration().pages[0].groups[0].groupFields[0];
    jest.spyOn(console, 'error').mockImplementation(() => undefined);
  });
  afterEach(() => { act(() => { field.properties.onDispose?.(host); }); host.remove(); jest.restoreAllMocks(); });
  function button(label: string): HTMLButtonElement {
    const result = Array.from(document.querySelectorAll('button')).find(item => item.textContent === label || item.querySelector('.ms-Button-label')?.textContent === label || item.getAttribute('aria-label') === label);
    if (!result) throw new Error('Missing button: ' + label);
    return result;
  }
  function click(label: string): void { act(() => { Simulate.click(button(label)); }); }
  function expectClosed(): void {
    const panel = document.querySelector('.ms-Panel');
    // Fluent UI keeps a dismissed panel mounted during its exit animation.
    expect(!panel || panel.getAttribute('aria-hidden') === 'true').toBe(true);
  }
  function check(label: string): void {
    const element = Array.from(document.querySelectorAll('label')).find(item => item.textContent === label)!;
    const input = document.getElementById(element.htmlFor) as HTMLInputElement;
    act(() => { input.checked = true; Simulate.change(input); });
  }
  function edit(): void {
    click('ManageAdminDefaultsLabel'); click('EditLabel');
    const input = Array.from(document.querySelectorAll('input')).find(item => item.value === 'Original calendar')!;
    act(() => { input.value = 'Updated calendar'; Simulate.change(input); });
    const dropdown = Array.from(document.querySelectorAll('[role="combobox"]')).find(item => item.getAttribute('aria-labelledby')?.split(' ').some(id => document.getElementById(id)?.textContent === 'PolicyLabel'))!;
    act(() => { Simulate.click(dropdown); });
    const option = Array.from(document.querySelectorAll('[role="option"]')).find(item => item.textContent === 'MandatoryPolicyLabel')!;
    act(() => { Simulate.click(option); });
  }
  async function addExchange(): Promise<void> {
    click('ManageAdminDefaultsLabel');
    click('AddAdminDefaultLabel');
    await act(async () => { Simulate.click(button('OutlookCalendarLabel')); });
    check('EveryoneAudienceLabel');
    click('NextLabel');
    const mailbox = document.querySelector('input[placeholder="MailboxPlaceholder"]') as HTMLInputElement;
    act(() => { mailbox.value = 'mailbox@example.com'; Simulate.change(mailbox); });
    await act(async () => { Simulate.click(button('LoadCalendarsLabel')); });
    check('New Exchange calendar');
    click('AddItemLabel');
  }

  it('saves existing source and assignment edits once and reopens the accepted configuration after host rerender', async () => {
    const commit = jest.fn((property: string | undefined, value: string | undefined) => {
      expect(property).toBe('adminSettings');
      webpart.properties.adminSettings = value;
      // SPFx also re-renders custom controls when a field change is transferred.
      field = webpart.getPropertyPaneConfiguration().pages[0].groups[0].groupFields[0];
      field.properties.onRender(host, undefined, commit);
    });
    await act(async () => { field.properties.onRender(host, undefined, commit); });
    edit();
    await act(async () => { Simulate.click(button('SaveLabel')); });
    expect(commit).toHaveBeenCalledTimes(1);
    expectClosed();
    const stored = loadAdminWebPartSettings({ current: webpart.properties.adminSettings }).settings;
    expect(stored.assignedSources[0]).toMatchObject({ assignmentId: 'assignment', adminSourceId: 'source', isMandatory: true, source: { name: 'Updated calendar' } });
    click('ManageAdminDefaultsLabel');
    expect(document.body.textContent).toContain('Updated calendar');
    expect(document.body.textContent).toContain('MandatoryPolicyLabel');
  });

  it('adds a selected Exchange calendar through the mailbox wizard and keeps it after Save and reopen', async () => {
    webpart._adminSettings = structuredClone(defaultAdminWebPartSettings);
    webpart.properties.adminSettings = serializeAdminWebPartSettings(webpart._adminSettings);
    const commit = jest.fn((_property: string | undefined, value: string | undefined) => { webpart.properties.adminSettings = value; });
    await act(async () => { field.properties.onRender(host, undefined, commit); });
    await addExchange();
    expect(document.body.textContent).toContain('New Exchange calendar');
    await act(async () => { Simulate.click(button('SaveLabel')); });
    expect(commit).toHaveBeenCalledTimes(1);
    expectClosed();
    expect(loadAdminWebPartSettings({ current: webpart.properties.adminSettings }).settings.assignedSources[0]).toMatchObject({
      source: { exchangeMailbox: 'new-mailbox', exchangeCalendarId: 'new-calendar', name: 'New Exchange calendar' }
    });
    click('ManageAdminDefaultsLabel');
    expect(document.body.textContent).toContain('New Exchange calendar');
  });

  it('retains the added Exchange draft and keeps Save available when the host callback silently ignores the transfer', async () => {
    const previous = webpart.properties.adminSettings;
    await act(async () => { field.properties.onRender(host, undefined, jest.fn()); });
    await addExchange();
    await act(async () => { Simulate.click(button('SaveLabel')); });
    expect(webpart.properties.adminSettings).toBe(previous);
    expect(webpart._adminSettings.assignedSources).toHaveLength(1);
    expect(document.body.textContent).toContain('New Exchange calendar');
    expect(document.querySelector('[role="alert"]')).not.toBeNull();
    expect(button('SaveLabel').disabled).toBe(false);
    expect(webpart.rebuildResolvedSettings).not.toHaveBeenCalled();
  });
});
