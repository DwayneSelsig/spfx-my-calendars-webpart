import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type { ICalendarEvent } from '../../models/ICalendarEvent';
import { getSourceIconName } from '../../utils/sourceIconHelper';
import { EventDetailsDialog } from './EventDetailsDialog';
import { addLocalDays, eventsForDay, getCalendarColor, isToday, startOfLocalDay } from './calendarUtils';
import { formatCalendarDate, formatCalendarTime } from './calendarFormatting';
import { EventStatusGlyphs, getEventStatusAriaText } from './EventStatusGlyphs';
import { getEventStatusPresentation } from './eventStatusPresentation';

export interface IMonthViewProps {
  appointments: ICalendarEvent[];
  currentDate: Date;
  locale?: string;
}

export const MonthView: React.FC<IMonthViewProps> = ({ appointments, currentDate, locale }) => {
  const restoreFocusRef = React.useRef<HTMLElement>();
  const [selectedEvent, setSelectedEvent] = React.useState<ICalendarEvent>();
  const firstOfMonth = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
  const gridStart = addLocalDays(firstOfMonth, -firstOfMonth.getDay());
  const days = Array.from({ length: 42 }, (_, index) => addLocalDays(gridStart, index));

  const dismiss = (): void => {
    setSelectedEvent(undefined);
    window.setTimeout(() => restoreFocusRef.current?.focus(), 0);
  };

  return (
    <div style={{ width: '100%', overflowX: 'auto' }}>
      <div role="grid" aria-label={formatCalendarDate(currentDate, { month: 'long', year: 'numeric' }, locale)} style={{ display: 'grid', gridTemplateColumns: 'repeat(7, minmax(112px, 1fr))', minWidth: 784, borderTop: '1px solid var(--neutralLight, #edebe9)', borderLeft: '1px solid var(--neutralLight, #edebe9)' }}>
        {Array.from({ length: 7 }, (_, index) => addLocalDays(gridStart, index)).map(day => (
          <div key={`header-${day.getDay()}`} role="columnheader" style={{ padding: '7px 8px', fontSize: 12, fontWeight: 600, textAlign: 'center', borderRight: '1px solid var(--neutralLight, #edebe9)', borderBottom: '1px solid var(--neutralLight, #edebe9)', background: 'var(--neutralLighterAlt, #faf9f8)' }}>
            {formatCalendarDate(day, { weekday: 'short' }, locale)}
          </div>
        ))}
        {days.map(day => {
          const inMonth = day.getMonth() === currentDate.getMonth();
          const dayEvents = eventsForDay(appointments, day).sort((a, b) => {
            if (!!a.isFullDay !== !!b.isFullDay) return a.isFullDay ? -1 : 1;
            return new Date(a.start).getTime() - new Date(b.start).getTime() || a.id.localeCompare(b.id);
          });
          return (
            <div key={startOfLocalDay(day).getTime()} role="gridcell" style={{ minHeight: 116, height: '14vh', maxHeight: 180, padding: 5, boxSizing: 'border-box', borderRight: '1px solid var(--neutralLight, #edebe9)', borderBottom: '1px solid var(--neutralLight, #edebe9)', background: inMonth ? 'var(--white, #fff)' : 'var(--neutralLighterAlt, #faf9f8)', color: inMonth ? 'inherit' : 'var(--neutralTertiary, #a19f9d)', overflow: 'hidden' }}>
              <div style={{ height: 24, display: 'flex', justifyContent: 'flex-end' }}>
                <span style={{ width: 24, height: 24, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', borderRadius: '50%', background: isToday(day) ? 'var(--themePrimary, #0078d4)' : 'transparent', color: isToday(day) ? 'var(--white, #fff)' : 'inherit', fontWeight: isToday(day) ? 600 : 400 }}>{day.getDate()}</span>
              </div>
              <div style={{ height: 'calc(100% - 24px)', overflowY: 'auto', overflowX: 'hidden', display: 'flex', flexDirection: 'column', gap: 3 }}>
                {dayEvents.map(event => {
                  const color = getCalendarColor(event);
                  const statusPresentation = getEventStatusPresentation(event, color, 16);
                  return (
                    <button
                      key={`${event.sourceId}:${event.id}`}
                      type="button"
                      onClick={click => { restoreFocusRef.current = click.currentTarget; setSelectedEvent(event); }}
                      title={event.title}
                      aria-label={[event.title, getEventStatusAriaText(event)].filter(Boolean).join(', ')}
                      style={{ display: 'block', width: '100%', flex: '0 0 auto', border: 0, borderLeft: `3px solid ${color}`, borderRadius: 2, padding: '3px 5px', textAlign: 'left', ...statusPresentation.cardStyle, cursor: 'pointer', fontSize: 11, lineHeight: '16px', fontStyle: event.isDraft ? 'italic' : 'normal', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                    >
                      {event.showSourceLogo !== false && <Icon iconName={getSourceIconName(event.sourceType, event.sourceIconName)} style={{ marginRight: 4 }} />}
                      <EventStatusGlyphs event={event} color={color} />
                      {!event.isFullDay && `${formatCalendarTime(new Date(event.start), locale)} `}
                      <span style={statusPresentation.titleStyle}>{event.title}</span>
                    </button>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>
      <EventDetailsDialog event={selectedEvent} onDismiss={dismiss} locale={locale} />
    </div>
  );
};
