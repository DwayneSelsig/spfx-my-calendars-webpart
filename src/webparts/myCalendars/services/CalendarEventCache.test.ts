import {
  CalendarEventCache,
  getCalendarEventCacheKey,
  normalizeCacheDuration,
  type ICalendarEventCacheConfiguration,
  type ICalendarEventCacheSegment
} from './CalendarEventCache';

const configuration: ICalendarEventCacheConfiguration = {
  tenantId: 'Tenant',
  userId: 'User',
  webPartInstanceId: 'WebPart',
  configSignature: 'signature'
};
const allowedMonthKeys = new Set(['2026-09']);

function createSegment(cachedAt: number, title: string = 'Cached event'): ICalendarEventCacheSegment {
  return {
    service: 'exchange',
    sourceId: 'exchange_calendar',
    monthKey: '2026-09',
    cachedAt,
    events: [{
      id: 'event-1',
      sourceId: 'exchange_calendar',
      title,
      start: '2026-09-10T08:00:00.000Z',
      end: '2026-09-10T09:00:00.000Z',
      sourceType: 'exchange'
    }]
  };
}

describe('CalendarEventCache', () => {
  beforeEach(() => window.localStorage.clear());
  afterEach(() => jest.restoreAllMocks());

  it('normalizes the administrator duration to whole minutes from 1 through 60', () => {
    expect(normalizeCacheDuration(1)).toBe(1);
    expect(normalizeCacheDuration(10)).toBe(10);
    expect(normalizeCacheDuration(60)).toBe(60);
    expect(normalizeCacheDuration(0)).toBe(1);
    expect(normalizeCacheDuration(61)).toBe(60);
    expect(normalizeCacheDuration(10.6)).toBe(11);
    expect(normalizeCacheDuration('invalid')).toBe(10);
  });

  it('marks a source/month segment stale when its configured duration is reached', () => {
    const cache = new CalendarEventCache();
    const cachedAt = Date.now() - 1000;
    expect(cache.replaceSegments(configuration, [createSegment(cachedAt)], allowedMonthKeys)).toBe(true);

    expect(cache.read(configuration, 10, allowedMonthKeys, cachedAt + 10 * 60 * 1000 - 1)?.segments[0].isStale).toBe(false);
    expect(cache.read(configuration, 10, allowedMonthKeys, cachedAt + 10 * 60 * 1000)?.segments[0].isStale).toBe(true);
  });

  it('replaces a successful source/month result, including with an empty result', () => {
    const cache = new CalendarEventCache();
    const cachedAt = Date.now() - 10_000;
    cache.replaceSegments(configuration, [createSegment(cachedAt)], allowedMonthKeys);

    cache.replaceSegments(configuration, [{
      ...createSegment(cachedAt + 1000),
      events: []
    }], allowedMonthKeys);

    expect(cache.read(configuration, 10, allowedMonthKeys, cachedAt + 2000)?.segments[0].events).toEqual([]);
    const persisted = window.localStorage.getItem(getCalendarEventCacheKey(configuration) as string);
    expect(persisted).toContain('"events":[]');
    expect(persisted).not.toContain('schemaVersion');
  });

  it('ignores a schema version in an existing entry and omits it after refresh', () => {
    const key = getCalendarEventCacheKey(configuration) as string;
    window.localStorage.setItem(key, JSON.stringify({
      schemaVersion: 1,
      tenantId: 'tenant', userId: 'user', webPartInstanceId: 'webpart',
      configSignature: 'signature', segments: [createSegment(Date.now())]
    }));
    const cache = new CalendarEventCache();

    expect(cache.read(configuration, 10, allowedMonthKeys)?.segments).toHaveLength(1);
    expect(cache.replaceSegments(configuration, [createSegment(Date.now(), 'Refreshed')], allowedMonthKeys)).toBe(true);
    expect(window.localStorage.getItem(key)).not.toContain('schemaVersion');
  });

  it('rejects invalid JSON and removes the complete entry', () => {
    const key = getCalendarEventCacheKey(configuration) as string;
    window.localStorage.setItem(key, '{invalid');

    expect(new CalendarEventCache().read(configuration, 10, allowedMonthKeys)).toBeUndefined();
    expect(window.localStorage.getItem(key)).toBeNull();
  });

  it.each([
    ['identity', { tenantId: 'other' }],
    ['signature', { configSignature: 'other' }],
    ['segments', { segments: [{ ...createSegment(Date.now()), cachedAt: -1 }] }],
    ['events', { segments: [{ ...createSegment(Date.now()), events: [{ ...createSegment(Date.now()).events[0], start: 'invalid' }] }] }]
  ])('rejects invalid %s data and removes the complete entry', (_name, override) => {
    const key = getCalendarEventCacheKey(configuration) as string;
    window.localStorage.setItem(key, JSON.stringify({
      tenantId: 'tenant', userId: 'user', webPartInstanceId: 'webpart',
      configSignature: 'signature', segments: [createSegment(Date.now())], ...override
    }));

    expect(new CalendarEventCache().read(configuration, 10, allowedMonthKeys)).toBeUndefined();
    expect(window.localStorage.getItem(key)).toBeNull();
  });

  it('recovers after localStorage reads become available again', () => {
    const cache = new CalendarEventCache();
    const key = getCalendarEventCacheKey(configuration) as string;
    const getItem = jest.spyOn(Storage.prototype, 'getItem').mockImplementation(() => { throw new Error('read failed'); });
    const removeItem = jest.spyOn(Storage.prototype, 'removeItem');

    expect(cache.read(configuration, 10, allowedMonthKeys)).toBeUndefined();
    expect(removeItem).toHaveBeenCalledWith(key);
    getItem.mockRestore();

    expect(cache.replaceSegments(configuration, [createSegment(Date.now())], allowedMonthKeys)).toBe(true);
    expect(cache.read(configuration, 10, allowedMonthKeys)?.segments).toHaveLength(1);
  });

  it('continues uncached when removal fails and can later write a fresh entry', () => {
    const cache = new CalendarEventCache();
    const key = getCalendarEventCacheKey(configuration) as string;
    window.localStorage.setItem(key, '{invalid');
    const removeItem = jest.spyOn(Storage.prototype, 'removeItem').mockImplementation(() => { throw new Error('remove failed'); });

    expect(cache.read(configuration, 10, allowedMonthKeys)).toBeUndefined();
    removeItem.mockRestore();

    expect(cache.replaceSegments(configuration, [createSegment(Date.now())], allowedMonthKeys)).toBe(true);
    expect(cache.read(configuration, 10, allowedMonthKeys)?.segments).toHaveLength(1);
  });

});
