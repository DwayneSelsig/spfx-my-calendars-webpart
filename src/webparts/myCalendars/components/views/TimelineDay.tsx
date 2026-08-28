import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type { ICalendarEvent } from '../../models/ICalendarEvent';
import { getSourceIconName } from '../../utils/sourceIconHelper';
import { getCalendarLabels } from './calendarLabels';
import {
  buildTimedSegments,
  eventsForDay,
  getCalendarColor,
  isToday,
  MINUTES_PER_DAY,
  PIXELS_PER_HOUR
} from './calendarUtils';

export interface ITimelineDayProps {
  day: Date;
  events: ICalendarEvent[];
  slotDurationMinutes: 15 | 30 | 60;
  onSelectEvent: (event: ICalendarEvent, focusElement?: HTMLElement) => void;
  showCurrentTime?: boolean;
  compact?: boolean;
}

const gutterWidth = 50;
export const ALL_DAY_SECTION_HEIGHT = 56;

export const TimelineDay: React.FC<ITimelineDayProps> = ({
  day,
  events,
  slotDurationMinutes,
  onSelectEvent,
  showCurrentTime = true,
  compact = false
}) => {
  const labels = getCalendarLabels();
  const [now, setNow] = React.useState(() => new Date());
  React.useEffect(() => {
    if (!showCurrentTime || !isToday(day)) return undefined;
    const timer = window.setInterval(() => setNow(new Date()), 30000);
    return () => window.clearInterval(timer);
  }, [day, showCurrentTime]);
  const dayEvents = eventsForDay(events, day);
  const allDayEvents = dayEvents.filter(event => event.isFullDay);
  const segments = buildTimedSegments(dayEvents, day);
  const slots = Array.from({ length: MINUTES_PER_DAY / slotDurationMinutes }, (_, index) => index);
  const hours = Array.from({ length: 24 }, (_, index) => index);
  const nowMinutes = now.getHours() * 60 + now.getMinutes();

  return (
    <div style={{ minWidth: 0, background: 'var(--white, #fff)' }}>
      <div
        aria-label={labels.allDay}
        style={{
          height: ALL_DAY_SECTION_HEIGHT,
          boxSizing: 'border-box',
          overflowY: 'auto',
          borderBottom: '1px solid var(--neutralLight, #edebe9)',
          padding: '4px 6px',
          display: 'flex',
          flexWrap: 'wrap',
          gap: 4,
          alignItems: 'flex-start',
          position: 'sticky',
          top: 0,
          zIndex: 5,
          background: 'var(--white, #fff)'
        }}
      >
        {allDayEvents.length === 0 && <span style={{ fontSize: 11, color: 'var(--neutralTertiary, #a19f9d)' }}>{labels.allDay}</span>}
        {allDayEvents.map(event => (
          <button
            key={`${event.sourceId}:${event.id}`}
            type="button"
            onClick={click => onSelectEvent(event, click.currentTarget)}
            style={{
              border: 0,
              borderLeft: `3px solid ${getCalendarColor(event)}`,
              borderRadius: 2,
              background: `color-mix(in srgb, ${getCalendarColor(event)} 18%, var(--white, #fff))`,
              color: 'var(--neutralPrimary, #323130)',
              padding: '3px 6px',
              minWidth: 0,
              maxWidth: '100%',
              cursor: 'pointer',
              fontStyle: event.isDraft ? 'italic' : 'normal',
              overflow: 'hidden',
              textOverflow: 'ellipsis',
              whiteSpace: 'nowrap'
            }}
          >
            {event.showSourceLogo !== false && <Icon iconName={getSourceIconName(event.sourceType, event.sourceIconName)} style={{ marginRight: 4 }} />}
            {event.title}
          </button>
        ))}
      </div>
      <div style={{ position: 'relative', height: MINUTES_PER_DAY, minWidth: compact ? 270 : 500 }}>
        {slots.map(slot => {
          const minutes = slot * slotDurationMinutes;
          const isHour = minutes % 60 === 0;
          return (
            <div
              key={slot}
              aria-hidden="true"
              style={{
                position: 'absolute',
                top: minutes,
                left: gutterWidth,
                right: 0,
                borderTop: `1px ${isHour ? 'solid' : 'dashed'} ${isHour ? 'var(--neutralLight, #edebe9)' : 'var(--neutralLighter, #f3f2f1)'}`
              }}
            />
          );
        })}
        {hours.map(hour => (
          <div
            key={hour}
            style={{
              position: 'absolute',
              top: hour * PIXELS_PER_HOUR - 8,
              width: gutterWidth - 7,
              textAlign: 'right',
              fontSize: 11,
              color: 'var(--neutralSecondary, #605e5c)'
            }}
          >
            {new Date(2000, 0, 1, hour).toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' })}
          </div>
        ))}
        {segments.map(segment => {
          const color = getCalendarColor(segment.event);
          const columnFraction = segment.column / segment.columnCount;
          const leftOffset = gutterWidth + 3 - (gutterWidth + 4) * columnFraction;
          const widthOffset = (gutterWidth + 4) / segment.columnCount + 3;
          return (
            <button
              key={`${segment.event.sourceId}:${segment.event.id}:${segment.startMinutes}`}
              type="button"
              onClick={click => onSelectEvent(segment.event, click.currentTarget)}
              aria-label={`${segment.event.title}, ${new Date(segment.event.start).toLocaleTimeString()}`}
              style={{
                position: 'absolute',
                top: segment.startMinutes + 1,
                height: Math.max(24, segment.endMinutes - segment.startMinutes - 2),
                left: `calc(${columnFraction * 100}% + ${leftOffset}px)`,
                width: `calc(${100 / segment.columnCount}% - ${widthOffset}px)`,
                zIndex: 2,
                boxSizing: 'border-box',
                overflow: 'hidden',
                textAlign: 'left',
                border: 0,
                borderLeft: `4px solid ${color}`,
                borderRadius: 3,
                padding: compact ? '3px 5px' : '5px 7px',
                background: `color-mix(in srgb, ${color} 18%, var(--white, #fff))`,
                color: 'var(--neutralPrimary, #323130)',
                cursor: 'pointer',
                fontStyle: segment.event.isDraft ? 'italic' : 'normal'
              }}
            >
              <div style={{ fontWeight: 600, fontSize: compact ? 11 : 12, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                {segment.event.showSourceLogo !== false && <Icon iconName={getSourceIconName(segment.event.sourceType, segment.event.sourceIconName)} style={{ marginRight: 4 }} />}
                {segment.event.title}
              </div>
              {(segment.endMinutes - segment.startMinutes >= 32) && (
                <div style={{ fontSize: 10, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                  {segment.event.location || (segment.event.isOnlineMeeting ? 'Microsoft Teams' : '')}
                </div>
              )}
              {!compact && segment.event.organizer?.name && (segment.endMinutes - segment.startMinutes >= 52) && (
                <div style={{ fontSize: 10, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{segment.event.organizer.name}</div>
              )}
            </button>
          );
        })}
        {showCurrentTime && isToday(day) && (
          <div aria-label={now.toLocaleTimeString()} style={{ position: 'absolute', top: nowMinutes, left: gutterWidth - 4, right: 0, zIndex: 4, borderTop: '2px solid #d13438', pointerEvents: 'none' }}>
            <span style={{ position: 'absolute', width: 9, height: 9, borderRadius: '50%', background: '#d13438', left: -4, top: -5 }} />
          </div>
        )}
      </div>
    </div>
  );
};
