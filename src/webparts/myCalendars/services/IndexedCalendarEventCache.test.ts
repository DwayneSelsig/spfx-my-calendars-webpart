import { CalendarEventCache, getCalendarEventCacheKey, type ICalendarEventCacheSegment } from './CalendarEventCache';
import { IndexedCalendarEventCache } from './IndexedCalendarEventCache';

const config = { tenantId: 'tenant', userId: 'user', webPartInstanceId: 'calendar', configSignature: 'source configuration' };
const monthKey = '2026-11';

function storeLegacySegment(service: ICalendarEventCacheSegment['service']): ICalendarEventCacheSegment {
  const time = '2026-11-22T10:00:00Z';
  const segment: ICalendarEventCacheSegment = { service, sourceId: 'source', monthKey, cachedAt: Date.now() - 1000,
    events: [{ id: 'event', sourceId: 'source', title: 'Stored appointment', start: time, end: time }] };
  new CalendarEventCache().replaceSegments(config, [segment], new Set([monthKey]));
  return segment;
}

describe('IndexedDB fallback and legacy appointment data', () => {
  beforeEach(() => localStorage.clear());

  it('keeps legacy appointments usable in a browser without IndexedDB and distrusts previously unpaged results', async () => {
    const segment = storeLegacySegment('exchange');
    const cache = new IndexedCalendarEventCache(config, true);
    expect(await cache.read('source', monthKey)).toEqual({ events: segment.events, completedAt: undefined });
    expect(await cache.preview([monthKey])).toEqual([{ sourceId: 'source', range: monthKey, events: segment.events, completedAt: undefined }]);
    expect(localStorage.getItem(getCalendarEventCacheKey(config)!)).not.toBeNull();
    cache.close();
  });

  it('clears the legacy cache when administrator caching is disabled', async () => {
    storeLegacySegment('exchange');
    const disabled = new IndexedCalendarEventCache(config, false);
    expect(await disabled.read('source', monthKey)).toBeUndefined();
    expect(localStorage.getItem(getCalendarEventCacheKey(config)!)).toBeNull();
    disabled.close();
  });
});
