import * as React from 'react';
import { Calendar } from '@pnp/spfx-controls-react/lib/controls/calendar';
import { IEvent } from '@pnp/spfx-controls-react/lib/controls/calendar/models/IEvents';
import { ECalendarViews } from '@pnp/spfx-controls-react/lib/controls/calendar/models/ECalendarViews';
import { Icon } from '@fluentui/react/lib/Icon';
import { getSourceIcon } from '../../utils/sourceIconHelper';

export interface ICommunityCalendarViewProps {
  appointments: IEvent[];
  currentDate: Date;
  onDateChange: (date: Date) => void;
  isLoading: boolean;
  startHour: number;
  endHour: number;
  showWeekends: boolean;
  viewType: 'day' | 'week' | 'month';
}

export const CommunityCalendarView: React.FC<ICommunityCalendarViewProps> = (props) => {
  const { appointments, onDateChange, viewType } = props;

  // Minimal custom renderer — matches MonthView structure: flexDirection:row + alignItems:center
  // ensures icon and title stay on one line. minWidth:0 on span prevents column stretching.
  const renderEventWithCustomStyle = (event: IEvent): JSX.Element => {
    const hex = event.colorHex || '#0078d4';
    return (
      <div style={{
        display: 'flex',
        flexDirection: 'row',
        alignItems: 'center',
        flexShrink: 0,
        minWidth: 0,
        minHeight: 22,
        width: '100%',
        maxWidth: '100%',
        overflow: 'hidden',
        boxSizing: 'border-box',
        borderLeft: `3px solid ${hex}`,
        borderRadius: 2,
        marginBottom: 5,
        backgroundColor: `color-mix(in srgb, ${hex} 20%, transparent)`,
        fontSize: 11,
      }}>
        {event.showSourceLogo && event.sourceType && (
          <Icon
            iconName={event.sourceIconName ?? getSourceIcon(event.sourceType)}
            style={{ marginLeft: 3, marginRight: 4, flexShrink: 0, color: hex }}
          />
        )}
        <span style={{
          display: 'block',
          fontStyle: event.isDraft ? 'italic' : 'normal',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap',
          minWidth: 0,
          width: 0,
          maxWidth: '100%',
          flex: 1
        }}>
          {event.title}
        </span>
      </div>
    );
  };

  // Enrich events with custom renderers and ensure color is set
  const enrichedEvents = React.useMemo(() => {
    return appointments.map(apt => {
      const safeCategory = typeof apt.category === 'string' && apt.category.trim().length > 0
        ? apt.category
        : 'default';

      const safeEvent = {
        ...apt,
        title: apt.title || '(Untitled event)',
        category: safeCategory,
        // color must be truthy so the PnP control skips the category.toLowerCase() fallback path
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        color: (apt.colorHex || '#0078d4') as any
      } as IEvent;

      // cloneElement in the PnP month renderer merges styles.eventCard (provides border-radius,
      // overflow, cursor) but skips the color class and margin-top. We handle both here:
      // colorHex left-border as color indicator, margin-top:5px to match the default card spacing.
      return {
        ...safeEvent,
        onRenderInDayView: () => renderEventWithCustomStyle(safeEvent),
        onRenderInWeekView: () => renderEventWithCustomStyle(safeEvent),
        onRenderInMonthView: () => renderEventWithCustomStyle(safeEvent),
      } as IEvent;
    });
  }, [appointments]);

  const defaultView = React.useMemo(() => {
    switch (viewType) {
      case 'day':
        return ECalendarViews.Day;
      case 'week':
        return ECalendarViews.Week;
      default:
        return ECalendarViews.Month;
    }
  }, [viewType]);

  return (
    <div style={{ width: '100%', height: '100%' }}>
      <Calendar
        events={enrichedEvents}
        defaultView={defaultView}
        onDayChange={onDateChange}
        onWeekChange={onDateChange}
        onMonthChange={onDateChange}
      />
    </div>
  );
};

