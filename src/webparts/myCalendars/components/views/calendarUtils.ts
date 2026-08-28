import type { ICalendarEvent } from '../../models/ICalendarEvent';

export const MINUTES_PER_DAY = 24 * 60;
export const PIXELS_PER_HOUR = 60;

export function minutesToTimelinePixels(minutes: number): number {
  return minutes * PIXELS_PER_HOUR / 60;
}

export interface ITimedEventSegment {
  event: ICalendarEvent;
  startMinutes: number;
  endMinutes: number;
  column: number;
  columnCount: number;
}

export function startOfLocalDay(value: Date): Date {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate());
}

export function addLocalDays(value: Date, days: number): Date {
  const result = startOfLocalDay(value);
  result.setDate(result.getDate() + days);
  return result;
}

export function localDateKey(value: Date): string {
  const year = value.getFullYear();
  const month = (`0${value.getMonth() + 1}`).slice(-2);
  const day = (`0${value.getDate()}`).slice(-2);
  return `${year}-${month}-${day}`;
}

export function parseLocalDateKey(value: string): Date | undefined {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) return undefined;
  const parsed = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  return Number.isNaN(parsed.getTime()) ? undefined : parsed;
}

export function getEventDates(event: ICalendarEvent): { start: Date; end: Date } | undefined {
  const start = new Date(event.start);
  const end = new Date(event.end);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return undefined;
  return { start, end: end > start ? end : new Date(start.getTime() + 60 * 1000) };
}

export function eventOccursOnDay(event: ICalendarEvent, day: Date): boolean {
  const dates = getEventDates(event);
  if (!dates) return false;
  const dayStart = startOfLocalDay(day);
  const dayEnd = addLocalDays(dayStart, 1);
  return dates.start < dayEnd && dates.end > dayStart;
}

export function eventsForDay(events: ICalendarEvent[], day: Date): ICalendarEvent[] {
  return events.filter(event => eventOccursOnDay(event, day));
}

export function buildTimedSegments(events: ICalendarEvent[], day: Date): ITimedEventSegment[] {
  const dayStart = startOfLocalDay(day);
  const dayEnd = addLocalDays(dayStart, 1);
  const segments = events
    .filter(event => !event.isFullDay)
    .map(event => {
      const dates = getEventDates(event);
      if (!dates || dates.start >= dayEnd || dates.end <= dayStart) return undefined;
      const visibleStart = dates.start < dayStart ? dayStart : dates.start;
      const visibleEnd = dates.end > dayEnd ? dayEnd : dates.end;
      return {
        event,
        startMinutes: Math.max(0, (visibleStart.getTime() - dayStart.getTime()) / 60000),
        endMinutes: Math.min(MINUTES_PER_DAY, (visibleEnd.getTime() - dayStart.getTime()) / 60000),
        column: 0,
        columnCount: 1
      } as ITimedEventSegment;
    })
    .filter((segment): segment is ITimedEventSegment => !!segment)
    .sort((a, b) => a.startMinutes - b.startMinutes || b.endMinutes - a.endMinutes || a.event.id.localeCompare(b.event.id));

  let group: ITimedEventSegment[] = [];
  let groupEnd = -1;
  const finishGroup = (): void => {
    if (!group.length) return;
    const columnEnds: number[] = [];
    group.forEach(segment => {
      let column = columnEnds.findIndex(end => end <= segment.startMinutes);
      if (column < 0) column = columnEnds.length;
      columnEnds[column] = Math.max(segment.endMinutes, segment.startMinutes + 24);
      segment.column = column;
    });
    const count = Math.max(1, columnEnds.length);
    group.forEach(segment => { segment.columnCount = count; });
    group = [];
  };

  segments.forEach(segment => {
    if (group.length && segment.startMinutes >= groupEnd) finishGroup();
    group.push(segment);
    groupEnd = Math.max(groupEnd, segment.endMinutes, segment.startMinutes + 24);
  });
  finishGroup();
  return segments;
}

export function getCalendarColor(event: ICalendarEvent): string {
  return event.colorHex || (typeof event.color === 'string' ? event.color : undefined) || '#0078d4';
}

export function isToday(day: Date): boolean {
  return localDateKey(day) === localDateKey(new Date());
}

export function safeOpen(url: string): void {
  const opened = window.open(url, '_blank', 'noopener,noreferrer');
  if (opened) opened.opener = null;
}
