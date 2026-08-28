import * as React from 'react';
import type { ICalendarEvent } from '../../models/ICalendarEvent';
import { EventDetailsDialog } from './EventDetailsDialog';
import { ALL_DAY_SECTION_HEIGHT, TimelineDay } from './TimelineDay';
import { addLocalDays, isToday, minutesToTimelinePixels } from './calendarUtils';

export interface IWeekViewProps {
  appointments: ICalendarEvent[];
  currentDate: Date;
  preferredStartMinutes: number;
  visibleHourCount: number;
  slotDurationMinutes: 15 | 30 | 60;
  showWeekends: boolean;
}

function getVisibleDays(start: Date, showWeekends: boolean): Date[] {
  if (showWeekends) return Array.from({ length: 7 }, (_, index) => addLocalDays(start, index));
  const result: Date[] = [];
  let offset = 0;
  while (result.length < 5) {
    const day = addLocalDays(start, offset++);
    if (day.getDay() !== 0 && day.getDay() !== 6) result.push(day);
  }
  return result;
}

export const WeekView: React.FC<IWeekViewProps> = ({ appointments, currentDate, preferredStartMinutes, visibleHourCount, slotDurationMinutes, showWeekends }) => {
  const days = React.useMemo(() => getVisibleDays(currentDate, showWeekends), [currentDate, showWeekends]);
  const scrollRefs = React.useRef<Array<HTMLDivElement | null>>([]);
  const syncingRef = React.useRef(false);
  const restoreFocusRef = React.useRef<HTMLElement>();
  const [selectedEvent, setSelectedEvent] = React.useState<ICalendarEvent>();

  React.useEffect(() => {
    const scrollTop = minutesToTimelinePixels(preferredStartMinutes);
    scrollRefs.current.forEach(element => { if (element) element.scrollTop = scrollTop; });
  }, [currentDate, preferredStartMinutes, showWeekends]);

  const synchronize = (sourceIndex: number): void => {
    if (syncingRef.current) return;
    const source = scrollRefs.current[sourceIndex];
    if (!source) return;
    syncingRef.current = true;
    scrollRefs.current.forEach((element, index) => {
      if (element && index !== sourceIndex && element.scrollTop !== source.scrollTop) element.scrollTop = source.scrollTop;
    });
    window.requestAnimationFrame(() => { syncingRef.current = false; });
  };

  const dismiss = (): void => {
    setSelectedEvent(undefined);
    window.setTimeout(() => restoreFocusRef.current?.focus(), 0);
  };

  return (
    <div style={{ overflowX: 'auto', overflowY: 'hidden', paddingBottom: 8 }}>
      <div style={{ display: 'flex', alignItems: 'stretch', gap: 8, minWidth: 'max-content' }}>
        {days.map((day, index) => (
          <section key={day.getTime()} style={{ width: 300, border: '1px solid var(--neutralLight, #edebe9)', borderRadius: 5, overflow: 'hidden', background: 'var(--white, #fff)' }}>
            <header style={{ padding: '9px 10px', borderBottom: '1px solid var(--neutralLight, #edebe9)', background: isToday(day) ? 'var(--themeLighterAlt, #eff6fc)' : 'var(--white, #fff)' }}>
              <div style={{ fontSize: 12, color: 'var(--neutralSecondary, #605e5c)', textTransform: 'capitalize' }}>{day.toLocaleDateString(undefined, { weekday: 'long' })}</div>
              <div style={{ fontSize: 18, fontWeight: 600 }}>{day.toLocaleDateString(undefined, { day: 'numeric', month: 'short' })}</div>
            </header>
            <div
              ref={element => { scrollRefs.current[index] = element; }}
              onScroll={() => synchronize(index)}
              style={{ height: Math.min(24, Math.max(1, visibleHourCount)) * 60 + ALL_DAY_SECTION_HEIGHT, overflowY: 'auto', overflowX: 'hidden' }}
            >
              <TimelineDay
                day={day}
                events={appointments}
                slotDurationMinutes={slotDurationMinutes}
                compact
                onSelectEvent={(event, focusElement) => { restoreFocusRef.current = focusElement; setSelectedEvent(event); }}
              />
            </div>
          </section>
        ))}
      </div>
      <EventDetailsDialog event={selectedEvent} onDismiss={dismiss} />
    </div>
  );
};
