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
    expect(window.localStorage.getItem(getCalendarEventCacheKey(configuration) as string)).toContain('"events":[]');
  });
});
