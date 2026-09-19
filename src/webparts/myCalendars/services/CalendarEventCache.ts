import type { ICalendarEvent } from '../models/ICalendarEvent';

export type CalendarCacheServiceKey = 'exchange' | 'ics' | 'sharepoint' | 'planner' | 'teamsShifts' | 'unifiedGroup';

export interface ICalendarEventCacheConfiguration {
  tenantId?: string;
  userId?: string;
  webPartInstanceId?: string;
  configSignature: string;
}

export interface ICalendarEventCacheSegment {
  service: CalendarCacheServiceKey;
  sourceId: string;
  monthKey: string;
  cachedAt: number;
  events: ICalendarEvent[];
}

export interface ICalendarEventCacheReadResult {
  segments: Array<ICalendarEventCacheSegment & { isStale: boolean }>;
}

interface ICalendarEventCacheEntry {
  tenantId: string;
  userId: string;
  webPartInstanceId: string;
  configSignature: string;
  segments: ICalendarEventCacheSegment[];
}

export const DEFAULT_CACHE_DURATION_MINUTES = 10;
export const MIN_CACHE_DURATION_MINUTES = 1;
export const MAX_CACHE_DURATION_MINUTES = 60;

const serviceKeys: CalendarCacheServiceKey[] = ['exchange', 'ics', 'sharepoint', 'planner', 'teamsShifts', 'unifiedGroup'];

function normalizeIdentifier(value: unknown): string | undefined {
  const normalized = typeof value === 'string' ? value.trim().toLowerCase() : undefined;
  return normalized || undefined;
}

export function normalizeCacheDuration(value: unknown): number {
  const numericValue = typeof value === 'number'
    ? value
    : typeof value === 'string' && value.trim() !== '' ? Number(value) : Number.NaN;
  if (!Number.isFinite(numericValue)) return DEFAULT_CACHE_DURATION_MINUTES;
  return Math.min(MAX_CACHE_DURATION_MINUTES, Math.max(MIN_CACHE_DURATION_MINUTES, Math.round(numericValue)));
}

export function getCalendarEventCacheKey(configuration: ICalendarEventCacheConfiguration): string | undefined {
  const tenantId = normalizeIdentifier(configuration.tenantId);
  const userId = normalizeIdentifier(configuration.userId);
  const webPartInstanceId = normalizeIdentifier(configuration.webPartInstanceId);
  if (!tenantId || !userId || !webPartInstanceId) return undefined;
  return `myCalendars:${encodeURIComponent(tenantId)}:${encodeURIComponent(userId)}:${encodeURIComponent(webPartInstanceId)}`;
}

export class CalendarEventCache {
  public read(
    configuration: ICalendarEventCacheConfiguration,
    cacheDurationMinutes: unknown,
    allowedMonthKeys: Set<string>,
    now: number = Date.now()
  ): ICalendarEventCacheReadResult | undefined {
    const entry = this.readEntry(configuration);
    if (!entry) return undefined;
    const durationMs = normalizeCacheDuration(cacheDurationMinutes) * 60 * 1000;
    return {
      segments: entry.segments
        .filter(segment => allowedMonthKeys.has(segment.monthKey))
        .map(segment => ({ ...segment, isStale: now - segment.cachedAt >= durationMs }))
    };
  }

  public replaceSegments(
    configuration: ICalendarEventCacheConfiguration,
    replacements: ICalendarEventCacheSegment[],
    allowedMonthKeys: Set<string>
  ): boolean {
    const key = getCalendarEventCacheKey(configuration);
    const identity = this.getIdentity(configuration);
    const storage = this.getStorage();
    if (!key || !identity || !storage || !replacements.every(segment => this.isValidSegment(segment))) return false;

    const previous = this.readEntry(configuration);
    const byIdentity = new Map<string, ICalendarEventCacheSegment>();
    (previous?.segments || [])
      .filter(segment => allowedMonthKeys.has(segment.monthKey))
      .forEach(segment => byIdentity.set(this.getSegmentKey(segment), segment));
    replacements
      .filter(segment => allowedMonthKeys.has(segment.monthKey))
      .forEach(segment => byIdentity.set(this.getSegmentKey(segment), segment));

    const entry: ICalendarEventCacheEntry = {
      ...identity,
      configSignature: configuration.configSignature,
      segments: Array.from(byIdentity.values())
    };
    try {
      storage.setItem(key, JSON.stringify(entry));
      return true;
    } catch (error) {
      this.warn('Could not write the appointment cache; continuing without persistent caching.', error);
      return false;
    }
  }

  public removeSegments(
    configuration: ICalendarEventCacheConfiguration,
    service: CalendarCacheServiceKey,
    sourceId: string,
    monthKeys: Set<string>,
    allowedMonthKeys: Set<string>
  ): void {
    const key = getCalendarEventCacheKey(configuration);
    const storage = this.getStorage();
    const entry = this.readEntry(configuration);
    if (!key || !storage || !entry) return;
    entry.segments = entry.segments.filter(segment => allowedMonthKeys.has(segment.monthKey) &&
      !(segment.service === service && segment.sourceId === sourceId && monthKeys.has(segment.monthKey)));
    try {
      storage.setItem(key, JSON.stringify(entry));
    } catch (error) {
      this.warn('Could not invalidate appointment-cache state; continuing with the current browser entry.', error);
    }
  }

  public remove(configuration: ICalendarEventCacheConfiguration): void {
    const key = getCalendarEventCacheKey(configuration);
    const storage = this.getStorage();
    if (key && storage) this.removeStorageItem(storage, key);
  }

  private readEntry(configuration: ICalendarEventCacheConfiguration): ICalendarEventCacheEntry | undefined {
    const key = getCalendarEventCacheKey(configuration);
    const identity = this.getIdentity(configuration);
    const storage = this.getStorage();
    if (!key || !identity || !storage) return undefined;
    let raw: string | null;
    try {
      raw = storage.getItem(key);
    } catch (error) {
      this.warn('Could not read the appointment cache; continuing without cached appointments.', error);
      this.removeStorageItem(storage, key);
      return undefined;
    }
    if (!raw) return undefined;
    try {
      const parsed = JSON.parse(raw) as unknown;
      if (!this.isValidEntry(parsed, identity, configuration.configSignature)) {
        this.warn('The appointment cache was incompatible or invalid; removing it.');
        this.removeStorageItem(storage, key);
        return undefined;
      }
      return parsed;
    } catch (error) {
      this.warn('The appointment cache contained invalid JSON; removing it.', error);
      this.removeStorageItem(storage, key);
      return undefined;
    }
  }

  private getIdentity(configuration: ICalendarEventCacheConfiguration): Pick<ICalendarEventCacheEntry, 'tenantId' | 'userId' | 'webPartInstanceId'> | undefined {
    const tenantId = normalizeIdentifier(configuration.tenantId);
    const userId = normalizeIdentifier(configuration.userId);
    const webPartInstanceId = normalizeIdentifier(configuration.webPartInstanceId);
    return tenantId && userId && webPartInstanceId ? { tenantId, userId, webPartInstanceId } : undefined;
  }

  private isValidEntry(value: unknown, identity: Pick<ICalendarEventCacheEntry, 'tenantId' | 'userId' | 'webPartInstanceId'>, signature: string): value is ICalendarEventCacheEntry {
    if (!this.isRecord(value) || value.tenantId !== identity.tenantId || value.userId !== identity.userId ||
      value.webPartInstanceId !== identity.webPartInstanceId || value.configSignature !== signature ||
      !Array.isArray(value.segments)) return false;
    return value.segments.every(segment => this.isValidSegment(segment));
  }

  private isValidSegment(value: unknown): value is ICalendarEventCacheSegment {
    if (!this.isRecord(value) || serviceKeys.indexOf(value.service as CalendarCacheServiceKey) < 0 ||
      typeof value.sourceId !== 'string' || !value.sourceId || typeof value.monthKey !== 'string' ||
      typeof value.cachedAt !== 'number' || !Number.isFinite(value.cachedAt) || value.cachedAt < 0 || value.cachedAt > Date.now() ||
      !Array.isArray(value.events)) return false;
    return value.events.every(event => this.isValidEvent(event));
  }

  private isValidEvent(value: unknown): value is ICalendarEvent {
    if (!this.isRecord(value) || typeof value.id !== 'string' || typeof value.sourceId !== 'string' ||
      typeof value.title !== 'string' || typeof value.start !== 'string' || typeof value.end !== 'string' ||
      !Number.isFinite(Date.parse(value.start)) || !Number.isFinite(Date.parse(value.end)) ||
      Date.parse(value.end) < Date.parse(value.start) ||
      (value.sourceType !== undefined && serviceKeys.indexOf(value.sourceType as CalendarCacheServiceKey) < 0)) return false;
    const optionalStrings = ['sourceDisplayName', 'category', 'description', 'location', 'importance', 'sensitivity', 'type', 'joinUrl', 'webLink', 'color', 'colorHex', 'sourceIconName', 'imageUrl'];
    if (optionalStrings.some(key => value[key] !== undefined && typeof value[key] !== 'string')) return false;
    const optionalBooleans = ['isFullDay', 'isOrganizer', 'isOnlineMeeting', 'showSourceLogo', 'isDraft'];
    if (optionalBooleans.some(key => value[key] !== undefined && typeof value[key] !== 'boolean')) return false;
    if (value.percentComplete !== undefined && (typeof value.percentComplete !== 'number' || !Number.isFinite(value.percentComplete))) return false;
    if (value.organizer !== undefined && (!this.isRecord(value.organizer) ||
      (value.organizer.name !== undefined && typeof value.organizer.name !== 'string') ||
      (value.organizer.email !== undefined && typeof value.organizer.email !== 'string'))) return false;
    if (value.attendees !== undefined && (!Array.isArray(value.attendees) || value.attendees.some(attendee =>
      !this.isRecord(attendee) || ['id', 'name', 'email', 'role', 'imageUrl'].some(key =>
        attendee[key] !== undefined && typeof attendee[key] !== 'string')))) return false;
    return true;
  }

  private getSegmentKey(segment: Pick<ICalendarEventCacheSegment, 'service' | 'sourceId' | 'monthKey'>): string {
    return `${segment.service}:${segment.sourceId}:${segment.monthKey}`;
  }

  private getStorage(): Storage | undefined {
    try {
      return typeof window !== 'undefined' ? window.localStorage : undefined;
    } catch (error) {
      this.warn('Browser localStorage is unavailable; continuing without persistent caching.', error);
      return undefined;
    }
  }

  private removeStorageItem(storage: Storage, key: string): void {
    try { storage.removeItem(key); } catch (error) { this.warn('Could not remove an invalid appointment cache.', error); }
  }

  private isRecord(value: unknown): value is Record<string, unknown> {
    return typeof value === 'object' && value !== null;
  }

  private warn(message: string, error?: unknown): void {
    if (error === undefined) console.warn(`[My Calendars] ${message}`);
    else console.warn(`[My Calendars] ${message}`, error);
  }
}
