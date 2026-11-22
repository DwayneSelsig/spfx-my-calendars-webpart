jest.mock('@microsoft/sp-core-library', () => ({ Version: { parse: jest.fn() } }));
jest.mock('@microsoft/sp-webpart-base', () => ({ BaseClientSideWebPart: class { public properties = {}; } }));
jest.mock('./components/MyCalendars', () => ({ __esModule: true, default: jest.fn() }));
jest.mock('./propertyPane/PropertyPaneAdminDefaultsManager', () => ({}));

import MyCalendarsWebPart from './MyCalendarsWebPart';
import { defaultAdminWebPartSettings, defaultAllowedOverrides, type IAdminWebPartSettings, type ICalendarSettings } from './models/ICalendarSettings';
import { serializeAdminWebPartSettings } from './services/CalendarSettingsService';
import type { GraphClient, GraphRequest } from './services/GraphRequestExecutor';
import type { GraphRequestScheduler } from './services/GraphRequestScheduler';

const current = '/me/drive/special/approot:/SPFx-My-Calendar-Webpart/user-calendar-settings.json:';
const legacy = '/me/drive/special/approot:/SPFx-My-Calendar-Webpart/calendar-settings.json:';
const folder = '/me/drive/special/approot/children';
const membership = '/me/checkMemberGroups';
const mailboxId = '12345678-1234-1234-1234-123456789012';
const mailbox = `/users/${mailboxId}`;
const calendars = `${mailbox}/calendars`;
const next = `https://graph.microsoft.com/v1.0${calendars}?$skiptoken=opaque`;

interface WebPartHarness {
  context: { pageContext: { aadInfo: unknown }; sdks: unknown; msGraphClientFactory: { getClient: () => Promise<GraphClient> } };
  properties: { adminSettings?: string };
  _scheduler: GraphRequestScheduler;
  _settingsGraph: GraphClient;
  _resolvedSettings: ICalendarSettings;
  render: jest.Mock;
  onInit: () => Promise<void>;
  onDispose: () => void;
  handleAdminSettingsSave: (settings: IAdminWebPartSettings, commit: (value: string | undefined) => void) => Promise<void>;
  handleRefreshAdminSources: () => Promise<ICalendarSettings>;
  handleDefaultViewChange: (view: 'week') => Promise<void>;
  handleResetUserSettings: () => Promise<void>;
}

interface ScheduledCall { name: string; high: number; normal: number }

function fixture(): { webPart: WebPartHarness; appSettings: IAdminWebPartSettings; calls: ScheduledCall[]; order: string[] } {
  const webPart = new MyCalendarsWebPart() as unknown as WebPartHarness;
  const appSettings: IAdminWebPartSettings = {
    ...defaultAdminWebPartSettings,
    maxConcurrentGraphRequests: 2,
    exchangeMailboxAssignments: [{ assignmentId: 'mailbox-rule', mailboxId, mailboxDisplayName: 'Shared mailbox', isMandatory: true, defaultEnabled: true,
      allowedOverrides: defaultAllowedOverrides, exceptions: [], audienceGroups: [{ groupId: 'audience', displayName: 'Audience' }] }]
  };
  const calls: ScheduledCall[] = [];
  const order: string[] = [];
  const graph: GraphClient = {
    api: path => {
      const execute = async (method: string): Promise<unknown> => {
        const { activeHigh, activeNormal } = webPart._scheduler.counts;
        const name = `${method} ${path}`;
        calls.push({ name, high: activeHigh, normal: activeNormal });
        order.push(name);
        if (name === `GET ${current}/content`) throw Object.assign(new Error('Current settings absent'), { statusCode: 404 });
        if (name === `GET ${legacy}/content`) return { defaultView: 'day', sources: [] };
        if (path === membership) return { value: ['audience'] };
        if (path === mailbox) return { id: mailboxId, userPrincipalName: 'mailbox@example.com' };
        if (path === calendars) return { value: [], '@odata.nextLink': next };
        return { value: [] };
      };
      const request: GraphRequest = {
        query: () => request, select: () => request, header: () => request, expand: () => request, top: () => request,
        get: () => execute('GET'), post: () => execute('POST'), put: () => execute('PUT'),
        patch: () => execute('PATCH'), delete: () => execute('DELETE')
      };
      return request;
    }
  };
  webPart.render = jest.fn();
  webPart.context = { pageContext: { aadInfo: {} }, sdks: {}, msGraphClientFactory: { getClient: async () => graph } };
  webPart.properties.adminSettings = serializeAdminWebPartSettings(appSettings);
  return { webPart, appSettings, calls, order };
}

describe('settings Graph priority in the web part', () => {
  beforeAll(() => {
    if (!globalThis.structuredClone) {
      Object.defineProperty(globalThis, 'structuredClone', {
        value: <T>(value: T): T => jest.requireActual('v8').deserialize(jest.requireActual('v8').serialize(value)) as T,
        configurable: true
      });
    }
  });

  beforeEach(() => sessionStorage.clear());

  it('runs OneDrive and assignment resolution at high priority while settings-panel calls stay normal', async () => {
    const { webPart, appSettings, calls } = fixture();
    await webPart.onInit();
    expect(webPart._resolvedSettings.defaultView).toBe('day');
    expect(calls.map(call => call.name)).toEqual([
      `GET ${current}/content`, `GET ${legacy}/content`, `POST ${folder}`, `PUT ${current}/content`,
      `POST ${membership}`, `GET ${mailbox}`, `GET ${calendars}`, `GET ${next}`
    ]);
    expect(calls.every(call => call.high === 1 && call.normal === 0)).toBe(true);

    for (const rebuild of [webPart.handleRefreshAdminSources, () => webPart.handleAdminSettingsSave(appSettings, value => {
      webPart.properties.adminSettings = value;
    })]) {
      sessionStorage.clear();
      calls.length = 0;
      await rebuild();
      expect(calls).toEqual([`POST ${membership}`, `GET ${mailbox}`, `GET ${calendars}`, `GET ${next}`]
        .map(name => ({ name, high: 1, normal: 0 })));
    }

    calls.length = 0;
    await webPart._settingsGraph.api('/groups').get();
    await webPart._settingsGraph.api('/$batch').post({ requests: [] });
    await webPart.handleResetUserSettings();
    expect(calls).toEqual([
      { name: 'GET /groups', high: 0, normal: 1 }, { name: 'POST /$batch', high: 0, normal: 1 },
      { name: `DELETE ${legacy}`, high: 1, normal: 0 }, { name: `DELETE ${current}`, high: 1, normal: 0 }
    ]);
    webPart.onDispose();
  });

  it('queues a OneDrive save behind an earlier visible-month request in the same high-priority queue', async () => {
    const { webPart, order } = fixture();
    await webPart.onInit();
    order.length = 0;
    let finish!: () => void;
    const gate = new Promise<void>(resolve => { finish = resolve; });
    const normal = webPart._scheduler.run(() => gate);
    const high = webPart._scheduler.run(() => gate, true);
    const month = webPart._scheduler.run(() => { order.push('visible month'); }, true);
    const save = webPart.handleDefaultViewChange('week');
    try {
      for (let index = 0; index < 10 && webPart._scheduler.counts.queuedHigh < 2; index++) await Promise.resolve();
      expect(webPart._scheduler.counts).toMatchObject({ queuedHigh: 2, queuedNormal: 0 });
    } finally {
      finish();
      await Promise.all([normal, high, month, save]);
      webPart.onDispose();
    }
    expect(order).toEqual(['visible month', `POST ${folder}`, `PUT ${current}/content`]);
  });

  it('cancels a waiting settings write when the web part is disposed', async () => {
    const { webPart, order } = fixture();
    await webPart.onInit();
    order.length = 0;
    let finish!: () => void;
    const gate = new Promise<void>(resolve => { finish = resolve; });
    const normal = webPart._scheduler.run(() => gate);
    const high = webPart._scheduler.run(() => gate, true);
    const logging = jest.spyOn(console, 'error').mockImplementation(() => undefined);
    const rejected = expect(webPart.handleDefaultViewChange('week')).rejects.toThrow('Failed to save personal settings');
    try {
      for (let index = 0; index < 10 && !webPart._scheduler.counts.queuedHigh; index++) await Promise.resolve();
      expect(webPart._scheduler.counts.queuedHigh).toBe(1);
      webPart.onDispose();
      await rejected;
      expect(webPart._scheduler.counts).toMatchObject({ activeHigh: 1, activeNormal: 1, queuedHigh: 0 });
      expect(order).toEqual([]);
    } finally {
      finish();
      await Promise.all([normal, high]);
      logging.mockRestore();
    }
  });
});
