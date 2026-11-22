import type { ICalendarEvent } from '../models/ICalendarEvent';
import { CalendarRangeCoordinator, calendarMonth, calendarMonthKey, calendarMonthRange, COLLECTION_RANGE, initialCalendarMonths } from './CalendarRangeCoordinator';
import type { CalendarRangeSource } from './CalendarRangeCoordinator';
import type { GraphRequestScope } from './GraphRequestScheduler';
import type { AppointmentCache, CachedCalendarResult } from './IndexedCalendarEventCache';

const today = new Date(2026, 10, 22);
const current = calendarMonth(today);

function appointment(month: number, id: string): ICalendarEvent {
  const start = calendarMonthRange(month).start.toISOString();
  return { id, sourceId: 'calendar', sourceType: 'exchange', title: id, start, end: start };
}

async function advanceUntil(condition: () => boolean): Promise<void> {
  for (let step = 0; step < 500 && !condition(); step++) await Promise.resolve();
  expect(condition()).toBe(true);
}

function deferred(): { promise: Promise<void>; resolve: () => void; reject: (reason: Error) => void } {
  let succeed!: () => void;
  let fail!: (reason: Error) => void;
  const promise = new Promise<void>((resolve, reject) => { succeed = resolve; fail = reject; });
  return { promise, resolve: succeed, reject: fail };
}

describe('calendar month coordination', () => {
  it('starts both jobs independently and visits 27 distinct local months in near-to-far order', async () => {
    const loaded: number[] = [];
    const gates = new Map<number, ReturnType<typeof deferred>>();
    const coordinator = new CalendarRangeCoordinator({ today, cacheDurationMs: 600000, onChange: () => undefined });
    coordinator.setVisible([current], false);
    coordinator.register({ id: 'calendar', service: 'exchange', load: async (range, scope) => {
      const month = calendarMonth(range!.start);
      expect(scope.high).toBe(false);
      loaded.push(month);
      const gate = gates.get(month);
      if (gate) await gate.promise;
    } });
    expect(loaded).toEqual([current, current - 1]);
    await advanceUntil(() => !coordinator.status().background);
    const { future, past } = initialCalendarMonths(today);
    expect(loaded.filter(month => month >= current)).toEqual(future);
    expect(loaded.filter(month => month < current)).toEqual(past);
    expect(new Set(loaded).size).toBe(27);
    expect(calendarMonthKey(future[13])).toBe('2027-12');
    expect(calendarMonthKey(past[12])).toBe('2025-10');
    coordinator.dispose();
  });

  it('uses half-open local boundaries across leap days, DST and year rollover', () => {
    for (const date of [new Date(2024, 1, 29), new Date(2026, 2, 29), new Date(2026, 9, 25), new Date(2026, 11, 31)]) {
      const month = calendarMonth(date);
      const range = calendarMonthRange(month);
      expect(range.start.getDate()).toBe(1);
      expect(range.start.getHours()).toBe(0);
      expect(range.end.getTime()).toBe(calendarMonthRange(month + 1).start.getTime());
      expect(new Date(range.end.getTime() - 1).getMonth()).toBe(date.getMonth());
    }
  });

  it('promotes the same operation during navigation, demotes obsolete work and fills future and past gaps normally', async () => {
    const requests: Array<{ month: number; scope: GraphRequestScope }> = [];
    const blocked = deferred();
    const coordinator = new CalendarRangeCoordinator({ today, cacheDurationMs: 600000, onChange: () => undefined });
    coordinator.setVisible([current], false);
    coordinator.register({ id: 'calendar', service: 'exchange', load: async (range, scope, page) => {
      const month = calendarMonth(range!.start);
      requests.push({ month, scope });
      if (month === current + 17) await blocked.promise;
      page([appointment(month, `event ${month}`)]);
    } });
    coordinator.setVisible([current - 1], true);
    expect(requests.filter(item => item.month === current - 1)).toHaveLength(1);
    expect(requests[1].scope.high).toBe(true);
    coordinator.setVisible([current + 17, current + 18], true);
    const earlier = requests.find(item => item.month === current + 17)!;
    expect(earlier.scope.high).toBe(true);
    coordinator.setVisible([current - 16], true);
    expect(earlier.scope.high).toBe(false);
    expect(requests.find(item => item.month === current - 16)?.scope.high).toBe(true);
    blocked.resolve();
    await advanceUntil(() => !coordinator.status().background && requests.length === 35);
    for (const month of [current + 14, current + 15, current + 16, current - 14, current - 15]) {
      expect(requests.find(item => item.month === month)?.scope.high).toBe(false);
    }
    expect(coordinator.appointments()).toContainEqual(appointment(current + 17, `event ${current + 17}`));
    coordinator.setVisible([current + 17], true);
    expect(requests.filter(item => item.month === current + 17)).toHaveLength(1);
    coordinator.dispose();
  });

  it('reuses fresh cache entries and keeps successful pages before a later-page failure', async () => {
    const laterPage = deferred();
    const saved = new Map<string, CachedCalendarResult>();
    const onBegin = jest.fn();
    saved.set(calendarMonthKey(current - 1), { completedAt: 100, events: [] });
    const cache: AppointmentCache = {
      read: async (_id, range) => saved.get(range),
      begin: (_id, range) => {
        const result: CachedCalendarResult = { events: [] };
        saved.set(range, result);
        onBegin(range);
        return { append: page => { result.events.push(...page); }, complete: at => { result.completedAt = at; } };
      }
    };
    const event = appointment(current, 'first page');
    const calls: number[] = [];
    const coordinator = new CalendarRangeCoordinator({ today, cacheDurationMs: 600000, cache, now: () => 100, onChange: () => undefined });
    coordinator.setVisible([current], false);
    const source: CalendarRangeSource = { id: 'calendar', service: 'exchange', load: async (range, _scope, page) => {
      const month = calendarMonth(range!.start);
      calls.push(month);
      if (month !== current) return;
      if (calls.filter(value => value === current).length === 1) {
        page([event, event]);
        await laterPage.promise;
      }
    } };
    coordinator.register(source);
    await advanceUntil(() => coordinator.appointments().length > 0);
    expect(coordinator.appointments()).toEqual([event]);
    expect(coordinator.status().pending.has('exchange')).toBe(true);
    expect(calls).not.toContain(current - 1);
    expect(saved.get(calendarMonthKey(current))?.completedAt).toBeUndefined();
    laterPage.reject(new Error('Page failed'));
    await advanceUntil(() => coordinator.status().visibleFailures.has('exchange'));
    expect(coordinator.appointments()).toEqual([event]);
    coordinator.setVisible([current], true);
    await advanceUntil(() => saved.get(calendarMonthKey(current))?.completedAt === 100);
    expect(coordinator.appointments()).toEqual([]);
    expect(saved.get(calendarMonthKey(current))?.events).toEqual([]);
    expect(onBegin.mock.calls.filter(call => call[0] === calendarMonthKey(current))).toHaveLength(2);
    coordinator.dispose();
  });

  it('shares one collection across all months, persists out-of-range events and ignores old generation pages', async () => {
    const gate = deferred();
    const outOfRange = appointment(current + 25, 'distant');
    const append = jest.fn();
    const complete = jest.fn();
    const begin = jest.fn(() => ({ append, complete }));
    const coordinator = new CalendarRangeCoordinator({ today, cacheDurationMs: 600000, cache: { read: async () => undefined, begin },
      onChange: () => undefined });
    coordinator.setVisible([current], false);
    const load = jest.fn<ReturnType<CalendarRangeSource['load']>, Parameters<CalendarRangeSource['load']>>(async (_range, _scope, publish) => {
      publish([outOfRange]);
      await gate.promise;
      publish([appointment(current, 'late')]);
    });
    coordinator.register({ id: 'calendar', service: 'planner', collection: true, load });
    await advanceUntil(() => append.mock.calls.length > 0);
    expect(begin).toHaveBeenCalledWith('calendar', COLLECTION_RANGE);
    expect(append).toHaveBeenCalledWith([outOfRange]);
    expect(coordinator.appointments()).toEqual([]);
    coordinator.setVisible([current + 25], true);
    expect(coordinator.appointments()).toEqual([outOfRange]);
    expect(load).toHaveBeenCalledTimes(1);
    expect(load.mock.calls[0][1].high).toBe(true);
    coordinator.dispose();
    gate.resolve();
    await advanceUntil(() => !coordinator.status().background);
    expect(append).toHaveBeenCalledTimes(1);
    expect(complete).not.toHaveBeenCalled();
  });
});
