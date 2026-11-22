import type { ICalendarEvent } from '../models/ICalendarEvent';
import { CalendarEventCache, getCalendarEventCacheKey, type ICalendarEventCacheConfiguration } from './CalendarEventCache';

const storeName = 'appointments';
const databaseName = 'my-calendars-appointment-cache';

interface CacheState { key: string[]; byRange: string[]; attempt: string; completedAt?: number }
interface CachedEvent { key: string[]; attempt: string; event: ICalendarEvent }
export interface CachedCalendarResult { events: ICalendarEvent[]; completedAt?: number }
export interface CachedCalendarPreview extends CachedCalendarResult { sourceId: string; range: string }

export interface CalendarCacheAttempt {
  append(events: ICalendarEvent[]): void;
  complete(completedAt: number): void;
}

export interface AppointmentCache {
  read(sourceId: string, range: string): Promise<CachedCalendarResult | undefined>;
  begin(sourceId: string, range: string): CalendarCacheAttempt;
  preview?(ranges: string[]): Promise<CachedCalendarPreview[]>;
}

function requestResult<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function transactionComplete(transaction: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onabort = () => reject(transaction.error || new Error('Appointment cache transaction aborted.'));
  });
}

function openDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(databaseName, 1);
    let blocked = false;
    request.onupgradeneeded = () => {
      request.result.createObjectStore(storeName, { keyPath: 'key' }).createIndex('byRange', 'byRange');
    };
    request.onerror = () => reject(request.error);
    request.onblocked = () => { blocked = true; reject(new Error('Appointment cache database is blocked.')); };
    request.onsuccess = () => {
      const database = request.result;
      database.onversionchange = () => database.close();
      if (blocked) database.close();
      else resolve(database);
    };
  });
}

function prefixRange(prefix: string[]): IDBKeyRange {
  return IDBKeyRange.bound(prefix, [...prefix, []]);
}

/** Each query has its own completion status and an incrementally updated set of events. */
export class IndexedCalendarEventCache implements AppointmentCache {
  private static readonly closing = new Map<string, Promise<void>>();
  private readonly legacy = new CalendarEventCache();
  private readonly namespace: string | undefined;
  private readonly ready: Promise<IDBDatabase | undefined>;
  private readonly fallback = new Map<string, CachedCalendarResult>();
  private writing: Promise<void> = Promise.resolve();
  private warned = false;
  private closed = false;
  private sequence = 0;

  public constructor(private readonly configuration: ICalendarEventCacheConfiguration, private readonly enabled: boolean) {
    this.namespace = getCalendarEventCacheKey(configuration);
    const closing = this.namespace && IndexedCalendarEventCache.closing.get(this.namespace);
    this.ready = Promise.resolve(closing).then(() => this.initialize()).catch(error => { this.warn(error); return undefined; });
  }

  private async initialize(): Promise<IDBDatabase | undefined> {
    if (!this.namespace) return undefined;
    const legacy = this.enabled ? this.legacy.readLegacySegments(this.configuration) : [];
    for (const segment of legacy) {
      if (segment.sourceId.startsWith('$service:')) continue;
      this.fallback.set(JSON.stringify([segment.sourceId, segment.monthKey]), {
        events: segment.events,
        completedAt: segment.service === 'teamsShifts' ? segment.cachedAt : undefined
      });
    }
    if (!this.enabled) this.legacy.remove(this.configuration);
    if (typeof indexedDB === 'undefined') return undefined;
    const database = await openDatabase();
    try {
      if (!this.enabled) {
        await this.write(database, store => { store.delete(prefixRange([this.namespace!])); });
        database.close();
        return undefined;
      }
      if (this.fallback.size) {
        await this.write(database, store => {
          this.fallback.forEach((result, entry) => {
            const [sourceId, month] = JSON.parse(entry) as [string, string];
            const key = this.key(sourceId, month);
            const existing = store.get(key);
            existing.onsuccess = () => {
              if (existing.result) return;
              const attempt = 'legacy';
              store.put(this.state(key, attempt, result.completedAt));
              this.writeEvents(store, key, attempt, result.events);
            };
          });
        });
        this.legacy.remove(this.configuration);
        this.fallback.clear();
      } else if (legacy.length) this.legacy.remove(this.configuration);
      return database;
    } catch (error) {
      database.close();
      throw error;
    }
  }

  private key(sourceId: string, range: string): string[] {
    return [this.namespace!, this.configuration.configSignature, sourceId, range];
  }

  private state(key: string[], attempt: string, completedAt?: number): CacheState {
    return { key, attempt, completedAt, byRange: [key[0], key[1], key[3]] };
  }

  public async preview(ranges: string[]): Promise<CachedCalendarPreview[]> {
    const results: CachedCalendarPreview[] = [];
    const database = await this.ready;
    if (!database || this.closed) {
      this.fallback.forEach((result, key) => {
        const [sourceId, range] = JSON.parse(key) as [string, string];
        if (ranges.indexOf(range) >= 0) results.push({ ...result, sourceId, range });
      });
      return results;
    }
    try {
      const index = database.transaction(storeName, 'readonly').objectStore(storeName).index('byRange');
      const matches = await Promise.all(ranges.map(range => requestResult<CacheState[]>(
        index.getAll([this.namespace!, this.configuration.configSignature, range])
      )));
      const states = matches.reduce<CacheState[]>((all, group) => all.concat(group), []);
      await Promise.all(states.map(async state => {
        const [, , sourceId, range] = state.key;
        const result = await this.read(sourceId, range);
        if (result) results.push({ ...result, sourceId, range });
      }));
    } catch (error) { this.warn(error); }
    return results;
  }

  public async read(sourceId: string, range: string): Promise<CachedCalendarResult | undefined> {
    const database = await this.ready;
    if (!database || this.closed) return this.fallback.get(JSON.stringify([sourceId, range]));
    const key = this.key(sourceId, range);
    try {
      const store = database.transaction(storeName, 'readonly').objectStore(storeName);
      const [state, events] = await Promise.all([
        requestResult<CacheState | undefined>(store.get(key)),
        requestResult<CachedEvent[]>(store.getAll(prefixRange([...key, 'event'])))
      ]);
      return state && { completedAt: state.completedAt, events: events.map(item => item.event) };
    } catch (error) { this.warn(error); return undefined; }
  }

  public begin(sourceId: string, range: string): CalendarCacheAttempt {
    const key = this.key(sourceId, range);
    const startedAt = performance.timeOrigin + performance.now();
    const attempt = `${startedAt}:${++this.sequence}:${Math.random()}`;
    let failed = false;
    const onFailure = (): void => { failed = true; };
    this.enqueue(store => {
      const request = store.get(key);
      request.onsuccess = () => {
        const existing = request.result as CacheState | undefined;
        const previousStart = Number(existing?.attempt.split(':')[0]);
        if (previousStart > startedAt) { onFailure(); return; }
        store.put(this.state(key, attempt));
      };
    }, onFailure);
    return {
      append: events => {
        if (events.length) this.updateAttempt(key, attempt, store => this.writeEvents(store, key, attempt, events), onFailure);
      },
      complete: completedAt => this.updateAttempt(key, attempt, store => {
        if (failed) return;
        const cursor = store.openCursor(prefixRange([...key, 'event']));
        cursor.onsuccess = () => {
          const record = cursor.result;
          if (!record) return;
          if ((record.value as CachedEvent).attempt !== attempt) record.delete();
          record.continue();
        };
        store.put(this.state(key, attempt, completedAt));
      }, onFailure)
    };
  }

  private updateAttempt(key: string[], attempt: string, update: (store: IDBObjectStore) => void, onFailure: () => void): void {
    this.enqueue(store => {
      const request = store.get(key);
      request.onsuccess = () => {
        if ((request.result as CacheState | undefined)?.attempt === attempt) update(store);
      };
    }, onFailure);
  }

  private writeEvents(store: IDBObjectStore, key: string[], attempt: string, events: ICalendarEvent[]): void {
    events.forEach(event => store.put({ key: [...key, 'event', event.id], attempt, event } as CachedEvent));
  }

  private enqueue(change: (store: IDBObjectStore) => void, onFailure: () => void): void {
    if (this.closed || !this.enabled) return;
    this.writing = this.writing.then(async () => {
      const database = await this.ready;
      if (database) await this.write(database, change);
    }).catch(error => { onFailure(); this.warn(error); });
  }

  private write(database: IDBDatabase, change: (store: IDBObjectStore) => void): Promise<void> {
    const transaction = database.transaction(storeName, 'readwrite');
    const completed = transactionComplete(transaction);
    try { change(transaction.objectStore(storeName)); }
    catch (error) { transaction.abort(); this.warn(error); }
    return completed;
  }

  public close(): void {
    if (this.closed) return;
    this.closed = true;
    const closing = this.writing.then(async () => (await this.ready)?.close()).catch(error => this.warn(error));
    const namespace = this.namespace;
    if (namespace) {
      IndexedCalendarEventCache.closing.set(namespace, closing);
      closing.then(() => {
        if (IndexedCalendarEventCache.closing.get(namespace) === closing) IndexedCalendarEventCache.closing.delete(namespace);
      }).catch(error => this.warn(error));
    }
  }

  private warn(error: unknown): void {
    if (this.warned) return;
    this.warned = true;
    console.warn('The browser appointment cache is unavailable for this operation.', error instanceof Error ? error.name : undefined);
  }
}
