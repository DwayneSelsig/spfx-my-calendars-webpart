import type { ICalendarEvent } from '../models/ICalendarEvent';
import type { CalendarCacheServiceKey } from './CalendarEventCache';
import { GraphRequestScope } from './GraphRequestScheduler';
import type { AppointmentCache, CachedCalendarPreview, CachedCalendarResult } from './IndexedCalendarEventCache';

export interface CalendarDateRange { start: Date; end: Date }
export type CalendarService = Exclude<CalendarCacheServiceKey, 'ics'>;

export interface CalendarRangeSource {
  id: string;
  service: CalendarService;
  collection?: boolean;
  load: (range: CalendarDateRange | undefined, scope: GraphRequestScope, onPage: (events: ICalendarEvent[]) => void) => Promise<void>;
}

interface Work {
  scope: GraphRequestScope;
  month?: number;
  promise: Promise<void>;
  events: Map<string, ICalendarEvent>;
  completedAt?: number;
  failed?: boolean;
  pending: boolean;
}

interface SourceState {
  source: CalendarRangeSource;
  months: Map<number | undefined, Work>;
  inspectedPast: number;
  inspectedFuture: number;
}

type DiscoveryLoader = (scope: GraphRequestScope) => Promise<CalendarRangeSource[]>;
interface Discovery { service: CalendarService; load: DiscoveryLoader; scope?: GraphRequestScope; pending: boolean; completed: boolean; failed: boolean }

export const COLLECTION_RANGE = 'collection';

const automaticSourcePrefix: Partial<Record<CalendarService, string>> = {
  exchange: 'exchange_', planner: 'auto_planner_', unifiedGroup: 'auto_unifiedGroup_'
};

export function calendarMonth(date: Date): number { return date.getFullYear() * 12 + date.getMonth(); }

function monthStart(month: number): Date {
  const date = new Date(0);
  date.setHours(0, 0, 0, 0);
  date.setFullYear(Math.floor(month / 12), month % 12 < 0 ? month % 12 + 12 : month % 12, 1);
  return date;
}

export function calendarMonthRange(month: number): CalendarDateRange {
  return { start: monthStart(month), end: monthStart(month + 1) };
}

export function calendarMonthKey(month: number): string {
  const start = monthStart(month);
  return `${start.getFullYear()}-${start.getMonth() + 1}`;
}

export function initialCalendarMonths(today: Date): { future: number[]; past: number[] } {
  const current = calendarMonth(today);
  return {
    future: Array.from({ length: 14 }, (_, index) => current + index),
    past: Array.from({ length: 13 }, (_, index) => current - 1 - index)
  };
}

export function eventIntersectsRange(event: ICalendarEvent, range: CalendarDateRange): boolean {
  const start = Date.parse(event.start);
  const end = Math.max(Date.parse(event.end), start + 1);
  return start < range.end.getTime() && end > range.start.getTime();
}

interface CoordinatorOptions {
  today: Date;
  cacheDurationMs: number;
  cache?: AppointmentCache;
  force?: boolean;
  previous?: CalendarRangeCoordinator;
  onChange: () => void;
  now?: () => number;
}

export interface CalendarLoadStatus {
  pending: Set<CalendarService>;
  failures: Set<CalendarService>;
  visibleFailures: Set<CalendarService>;
  background: boolean;
}

/** Owns month and source lifetimes. React only consumes snapshots of events and status. */
export class CalendarRangeCoordinator {
  private readonly sources = new Map<string, SourceState>();
  private readonly discovery = new Map<CalendarService, Discovery>();
  private readonly scopes = new Set<GraphRequestScope>();
  private readonly initial: ReturnType<typeof initialCalendarMonths>;
  private readonly now: () => number;
  private readonly previous = new Map<string, Map<number | undefined, Map<string, ICalendarEvent>>>();
  private readonly interned = new Map<string, ICalendarEvent>();
  private readonly previews = new Map<string, CachedCalendarPreview>();
  private readonly expanding = new Set<SourceState>();
  private canShowSource = (_sourceId: string): boolean => true;
  private visible: number[] = [];
  private highPriority = false;
  private low: number;
  private high: number;
  private backgroundJobs = 0;
  private stopped = false;

  public constructor(private readonly options: CoordinatorOptions) {
    this.initial = initialCalendarMonths(options.today);
    this.low = this.initial.past[this.initial.past.length - 1];
    this.high = this.initial.future[this.initial.future.length - 1];
    this.now = options.now || Date.now;
    options.previous?.sources.forEach(state => {
      const months = new Map<number | undefined, Map<string, ICalendarEvent>>();
      state.months.forEach((work, month) => months.set(month, work.events));
      this.previous.set(state.source.id, months);
    });
    options.previous?.previous.forEach((months, source) => {
      if (!this.previous.has(source)) this.previous.set(source, new Map(months));
    });
  }

  public register(source: CalendarRangeSource): void {
    if (this.stopped || this.sources.has(source.id)) return;
    const state: SourceState = { source, months: new Map(),
      inspectedPast: this.initial.past[this.initial.past.length - 1],
      inspectedFuture: this.initial.future[this.initial.future.length - 1] };
    this.sources.set(source.id, state);
    // Start the two nearest months before scheduling the remaining visible months.
    this.backgroundJobs++;
    const future = this.walk(state, this.initial.future);
    const past = this.walk(state, this.initial.past);
    this.scheduleVisible(state);
    const finish = (): void => {
      this.backgroundJobs--;
      this.expand(state);
      this.changed();
    };
    Promise.all([future, past]).then(finish, finish).catch(() => undefined);
    this.changed();
  }

  public discover(service: CalendarService, load: DiscoveryLoader): void {
    const discovery: Discovery = { service, load, pending: false, completed: false, failed: false };
    this.discovery.set(service, discovery);
    this.startDiscovery(discovery);
  }

  private startDiscovery(discovery: Discovery): void {
    if (this.stopped || discovery.pending || discovery.completed) return;
    const scope = this.createScope(this.highPriority);
    discovery.scope = scope;
    discovery.pending = true;
    discovery.failed = false;
    this.changed();
    Promise.resolve().then(() => discovery.load(scope)).then(sources => {
      if (this.stopped) return;
      discovery.completed = true;
      sources.forEach(source => this.register(source));
      const prefix = automaticSourcePrefix[discovery.service];
      if (prefix) {
        const removed = (id: string): boolean => id.startsWith(prefix) && !this.sources.has(id);
        this.previous.forEach((_months, id) => { if (removed(id)) this.previous.delete(id); });
        this.previews.forEach((result, key) => { if (removed(result.sourceId)) this.previews.delete(key); });
      }
    }).catch(() => { discovery.failed = true; }).then(() => {
      discovery.pending = false;
      this.scopes.delete(scope);
      this.changed();
    }).catch(() => undefined);
  }

  public setVisible(months: number[], highPriority: boolean): void {
    if (this.stopped || !months.length) return;
    this.visible = months;
    this.highPriority = highPriority;
    this.low = Math.min(this.low, ...months);
    this.high = Math.max(this.high, ...months);
    this.discovery.forEach(discovery => {
      discovery.scope?.setHigh(highPriority);
      if (highPriority) this.startDiscovery(discovery);
    });
    this.sources.forEach(state => {
      state.months.forEach(work => work.scope.setHigh(this.isHigh(work.month)));
      this.scheduleVisible(state);
      this.expand(state);
    });
    this.changed();
  }

  public restrictExisting(canShowSource: (sourceId: string) => boolean): void {
    this.canShowSource = canShowSource;
    this.previous.forEach((_months, id) => { if (!canShowSource(id)) this.previous.delete(id); });
  }

  public async warmVisibleCache(): Promise<void> {
    if (this.options.force || !this.options.cache?.preview) return;
    const months = this.visible.map(calendarMonthKey);
    const results = await this.options.cache.preview([...months, COLLECTION_RANGE]);
    if (this.stopped) return;
    results.forEach(result => {
      if (!this.canShowSource(result.sourceId)) return;
      const state = this.sources.get(result.sourceId);
      if (!state && Array.from(this.discovery.values()).some(discovery => {
        const prefix = automaticSourcePrefix[discovery.service];
        return discovery.completed && prefix && result.sourceId.startsWith(prefix);
      })) return;
      const loaded = state && Array.from(state.months.values()).some(work =>
        work.completedAt !== undefined && (work.month === undefined || calendarMonthKey(work.month) === result.range));
      if (!loaded) this.previews.set(JSON.stringify([result.sourceId, result.range]), result);
    });
    this.changed();
  }

  private scheduleVisible(state: SourceState): void {
    this.visible.forEach(month => { this.ensure(state, month, this.highPriority).catch(() => undefined); });
  }

  private async walk(state: SourceState, months: number[]): Promise<void> {
    if (state.source.collection) {
      await this.ensure(state, months[0], false);
      return;
    }
    for (const month of months) {
      if (this.stopped) return;
      await this.ensure(state, month, false);
    }
  }

  private expand(state: SourceState): void {
    if (this.stopped || state.source.collection) return;
    const direction = async (side: 'inspectedPast' | 'inspectedFuture', step: -1 | 1): Promise<void> => {
      while (!this.stopped && (step < 0 ? state[side] > this.low : state[side] < this.high)) {
        const month = state[side] + step;
        state[side] = month;
        await this.ensure(state, month, false);
      }
    };
    if (this.expanding.has(state) || state.inspectedPast <= this.low && state.inspectedFuture >= this.high) return;
    this.expanding.add(state);
    const finish = (): void => {
      this.expanding.delete(state);
      if (this.stopped) return;
      if (state.inspectedPast > this.low || state.inspectedFuture < this.high) this.expand(state);
      this.changed();
    };
    Promise.all([direction('inspectedPast', -1), direction('inspectedFuture', 1)]).then(finish, finish).catch(() => undefined);
  }

  private isHigh(month?: number): boolean {
    return this.highPriority && (month === undefined || this.visible.indexOf(month) >= 0);
  }

  private createScope(high: boolean): GraphRequestScope {
    const scope = new GraphRequestScope();
    scope.setHigh(high);
    this.scopes.add(scope);
    return scope;
  }

  private ensure(state: SourceState, visibleMonth: number, retry: boolean): Promise<void> {
    if (this.stopped) return Promise.resolve();
    const month = state.source.collection ? undefined : visibleMonth;
    const previous = state.months.get(month);
    if (previous && (previous.pending || this.fresh(previous.completedAt) || !retry)) return previous.promise;

    const scope = this.createScope(this.isHigh(month));
    const work: Work = {
      scope, month, pending: true, promise: Promise.resolve(),
      events: previous?.events || this.previous.get(state.source.id)?.get(month) || new Map()
    };
    state.months.set(month, work);
    work.promise = this.retrieve(state.source, work, !!previous).catch(() => {
      work.failed = true;
    }).then(() => {
      work.pending = false;
      this.scopes.delete(scope);
      this.changed();
    });
    this.changed();
    return work.promise;
  }

  private async retrieve(source: CalendarRangeSource, work: Work, alreadyRead: boolean): Promise<void> {
    const key = work.month === undefined ? COLLECTION_RANGE : calendarMonthKey(work.month);
    if (!alreadyRead && !this.options.force && this.options.cache) {
      const cached = await this.options.cache.read(source.id, key);
      if (this.stopped) return;
      if (cached && this.hydrate(source.id, work, cached)) return;
    }
    if (this.stopped) return;
    const attempt = this.options.cache?.begin(source.id, key);
    const results = new Map<string, ICalendarEvent>();
    await source.load(work.month === undefined ? undefined : calendarMonthRange(work.month), work.scope, events => {
      if (this.stopped || !events.length) return;
      this.index(events, results);
      this.index(events, work.events);
      attempt?.append(events);
      this.changed();
    });
    if (this.stopped) return;
    const completedAt = this.now();
    this.finish(source.id, work, results, completedAt);
    attempt?.complete(completedAt);
  }

  private hydrate(sourceId: string, work: Work, cached: CachedCalendarResult): boolean {
    work.events = this.index(cached.events);
    if (this.fresh(cached.completedAt)) {
      work.completedAt = cached.completedAt;
      this.clearPreview(sourceId, work.month);
      return true;
    }
    this.changed();
    return false;
  }

  private finish(sourceId: string, work: Work, results: Map<string, ICalendarEvent>, completedAt: number): void {
    work.events = results;
    this.previous.get(sourceId)?.delete(work.month);
    this.clearPreview(sourceId, work.month);
    work.completedAt = completedAt;
  }

  private clearPreview(sourceId: string, month?: number): void {
    const range = month === undefined ? undefined : calendarMonthKey(month);
    this.previews.forEach((result, key) => {
      if (result.sourceId === sourceId && (range === undefined || result.range === range)) this.previews.delete(key);
    });
  }

  private index(events: ICalendarEvent[], target = new Map<string, ICalendarEvent>()): Map<string, ICalendarEvent> {
    for (const event of events) {
      const key = `${event.sourceId}:${event.id}`;
      const existing = this.interned.get(key);
      const canonical = existing && JSON.stringify(existing) === JSON.stringify(event) ? existing : event;
      this.interned.set(key, canonical);
      target.set(key, canonical);
    }
    return target;
  }

  private fresh(completedAt?: number): boolean {
    return completedAt !== undefined && this.now() - completedAt < this.options.cacheDurationMs;
  }

  private changed(): void { if (!this.stopped) this.options.onChange(); }

  public appointments(): ICalendarEvent[] {
    const events = new Map<string, ICalendarEvent>();
    const range = { start: monthStart(this.low), end: monthStart(this.high + 1) };
    this.previews.forEach(result => result.events.forEach(event => {
      if (eventIntersectsRange(event, range)) events.set(`${event.sourceId}:${event.id}`, event);
    }));
    this.previous.forEach(months => months.forEach(month => month.forEach((event, key) => events.set(key, event))));
    this.sources.forEach(state => state.months.forEach(work => work.events.forEach((event, key) => {
      if (!state.source.collection || eventIntersectsRange(event, range)) events.set(key, event);
    })));
    return Array.from(events.values());
  }

  public status(): CalendarLoadStatus {
    const pending = new Set<CalendarService>();
    const failures = new Set<CalendarService>();
    const visibleFailures = new Set<CalendarService>();
    let background = false;
    this.discovery.forEach((state, service) => {
      if (state.pending) pending.add(service);
      if (state.failed) { failures.add(service); visibleFailures.add(service); }
    });
    this.sources.forEach(state => {
      const service = state.source.service;
      state.months.forEach(work => {
        const visible = work.month === undefined || this.visible.indexOf(work.month) >= 0;
        if (work.pending) { background = true; if (visible) pending.add(service); }
        if (work.failed) { failures.add(service); if (visible) visibleFailures.add(service); }
      });
    });
    return { pending, failures, visibleFailures, background: background || pending.size > 0 || this.backgroundJobs > 0 || this.expanding.size > 0 };
  }

  public dispose(): void {
    this.stopped = true;
    this.scopes.forEach(scope => scope.cancel());
    this.scopes.clear();
  }
}
