import * as React from 'react';
import type { ICalendarEvent } from '../../models/ICalendarEvent';
import { EventDetailsDialog } from './EventDetailsDialog';
import { ALL_DAY_SECTION_HEIGHT, TimelineDay } from './TimelineDay';
import { minutesToTimelinePixels } from './calendarUtils';

export interface IDayViewProps {
  appointments: ICalendarEvent[];
  currentDate: Date;
  preferredStartMinutes: number;
  visibleHourCount: number;
  slotDurationMinutes: 15 | 30 | 60;
}

export const DayView: React.FC<IDayViewProps> = ({ appointments, currentDate, preferredStartMinutes, visibleHourCount, slotDurationMinutes }) => {
  const scrollRef = React.useRef<HTMLDivElement>(null);
  const restoreFocusRef = React.useRef<HTMLElement>();
  const [selectedEvent, setSelectedEvent] = React.useState<ICalendarEvent>();

  React.useEffect(() => {
    if (scrollRef.current) scrollRef.current.scrollTop = minutesToTimelinePixels(preferredStartMinutes);
  }, [currentDate, preferredStartMinutes]);

  const dismiss = React.useCallback((): void => {
    setSelectedEvent(undefined);
    window.setTimeout(() => restoreFocusRef.current?.focus(), 0);
  }, []);

  return (
    <div>
      <div ref={scrollRef} style={{ height: Math.min(24, Math.max(1, visibleHourCount)) * 60 + ALL_DAY_SECTION_HEIGHT, overflowY: 'auto', overflowX: 'auto', border: '1px solid var(--neutralLight, #edebe9)', borderRadius: 4 }}>
        <TimelineDay
          day={currentDate}
          events={appointments}
          slotDurationMinutes={slotDurationMinutes}
          onSelectEvent={(event, focusElement) => { restoreFocusRef.current = focusElement; setSelectedEvent(event); }}
        />
      </div>
      <EventDetailsDialog event={selectedEvent} onDismiss={dismiss} />
    </div>
  );
};
