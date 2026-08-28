import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { ICalendarViewProps } from './DayView';
import { getSourceIcon } from '../../utils/sourceIconHelper';
import { IEvent } from '@pnp/spfx-controls-react/lib/controls/calendar/models/IEvents';

const styles = mergeStyleSets({
  weekView: {
    width: '100%',
    height: '100%',
    padding: 16,
    display: 'flex',
    flexDirection: 'column',
    overflow: 'auto',
    background: 'var(--neutralLighterAlt, #faf9f8)',
    boxSizing: 'border-box'
  },
  weekGrid: {
    display: 'inline-flex',
    minWidth: '100%',
    border: '1px solid var(--neutralLight, #edebe9)',
    borderRadius: 10,
    background: 'var(--white, #ffffff)',
    overflowY: 'auto'
  },
  timeColumn: {
    width: 60,
    flexShrink: 0,
    background: 'var(--neutralLighterAlt, #faf9f8)',
    borderRight: '1px solid var(--neutralLight, #edebe9)',
    display: 'flex',
    flexDirection: 'column'
  },
  dayHeader: {
    padding: '10px 6px 12px',
    textAlign: 'center',
    borderBottom: '1px solid var(--neutralLight, #edebe9)',
    backgroundColor: 'var(--neutralLighterAlt, #faf9f8)',
    minHeight: 66,
    boxSizing: 'border-box',
    flexShrink: 0
  },
  allDayPlaceholder: {
    borderBottom: '1px solid var(--neutralLight, #edebe9)',
    height: 48,
    minHeight: 48,
    backgroundColor: 'var(--neutralLighterAlt, #faf9f8)',
    boxSizing: 'border-box'
  },
  timeColumnSlots: {
    display: 'grid',
    gridAutoRows: '60px',
    flex: 1
  },
  hourLabel: {
    padding: '6px 8px',
    fontSize: 11,
    color: 'var(--neutralTertiary, #a19f9d)',
    height: 60,
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'flex-end',
    boxSizing: 'border-box'
  },
  dayColumn: {
    flex: 1,
    minWidth: 140,
    borderLeft: '1px solid var(--neutralLight, #edebe9)',
    background: 'var(--white, #ffffff)',
    position: 'relative',
    display: 'flex',
    flexDirection: 'column'
  },
  weekendColumn: {
    background: 'var(--neutralLighter, #f3f2f1)'
  },
  todayColumn: {
    background: 'var(--themeLighter, #eef6fc)'
  },
  todayHeader: {
    color: 'var(--themePrimary, #0078d4)'
  },
  dayName: {
    fontSize: 11,
    color: 'var(--neutralSecondary, #605e5c)',
    textTransform: 'uppercase',
    letterSpacing: '0.08em'
  },
  dayNumber: {
    fontSize: 20,
    fontWeight: 700,
    color: 'var(--neutralPrimary, #323130)',
    marginTop: 4
  },
  allDaySection: {
    borderBottom: '1px solid var(--neutralLight, #edebe9)',
    padding: 8,
    minHeight: 48,
    backgroundColor: 'var(--neutralLighter, #f3f2f1)',
    boxSizing: 'border-box'
  },
  allDayAppointment: {
    fontSize: 11,
    padding: '4px 6px',
    marginBottom: 3,
    borderRadius: 3,
    borderLeft: '3px solid',
    cursor: 'pointer',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap'
  },
  daySlotsWrapper: {
    position: 'relative',
    flex: 1
  },
  hourBorders: {
    position: 'absolute',
    inset: 0,
    pointerEvents: 'none',
    zIndex: 1
  },
  hourBorderLine: {
    height: 60,
    borderTop: '1px solid var(--neutralLight, #edebe9)',
    width: '100%',
    boxSizing: 'border-box'
  },
  daySlots: {
    position: 'relative',
    display: 'grid',
    gridAutoRows: '60px'
  },
  hourSlot: {
    minHeight: 60
  },
  appointmentsLayer: {
    position: 'absolute',
    inset: 0,
    pointerEvents: 'none',
    zIndex: 2
  },
  appointment: {
    padding: '6px 8px',
    borderRadius: 6,
    borderLeft: '4px solid',
    cursor: 'pointer',
    position: 'relative',
    zIndex: 1,
    pointerEvents: 'auto',
    boxSizing: 'border-box',
    overflow: 'hidden'
  },
  appointmentTitle: {
    fontSize: 12,
    fontWeight: 600,
    color: 'var(--neutralPrimary, #323130)',
    marginBottom: 4,
    display: 'flex',
    alignItems: 'center',
    gap: 4,
    minWidth: 0
  },
  appointmentTime: {
    fontSize: 10,
    color: 'var(--neutralSecondary, #605e5c)',
    marginTop: 2,
    lineHeight: 1.2
  }
});

export const WeekView: React.FC<ICalendarViewProps> = (props) => {
  const { appointments, currentDate, startHour, endHour, showWeekends } = props;
  const weekGridRef = React.useRef<HTMLDivElement>(null);
  const [pixelsPerMinute, setPixelsPerMinute] = React.useState(1);

  const getWeekDays = (): Date[] => {
    const days: Date[] = [];
    const startOfWeek = new Date(currentDate);
    const day = startOfWeek.getDay();
    const diff = day === 0 ? -6 : 1 - day; // Monday = 1
    startOfWeek.setDate(startOfWeek.getDate() + diff);

    for (let i = 0; i < (showWeekends ? 7 : 5); i++) {
      const date = new Date(startOfWeek);
      date.setDate(startOfWeek.getDate() + i);
      days.push(date);
    }
    return days;
  };

  const weekDays = getWeekDays();
  const hours: number[] = [];
  for (let h = startHour; h < endHour; h++) {
    hours.push(h);
  }

  React.useLayoutEffect(() => {
    if (!weekGridRef.current) {
      return;
    }

    const timer = setTimeout(() => {
      if (!weekGridRef.current) {
        return;
      }

      const hourLines = weekGridRef.current.querySelectorAll(`.${styles.hourBorderLine}`) as NodeListOf<HTMLElement>;
      if (hourLines.length >= 2) {
        const firstTop = hourLines[0].getBoundingClientRect().top;
        const secondTop = hourLines[1].getBoundingClientRect().top;
        const hourHeight = secondTop - firstTop;
        setPixelsPerMinute(hourHeight / 60);
        return;
      }

      const firstHourLine = hourLines[0];
      const fallbackHeight = firstHourLine?.offsetHeight ?? 60;
      setPixelsPerMinute(fallbackHeight / 60);
    }, 0);

    return () => clearTimeout(timer);
  }, [startHour, endHour, showWeekends]);

  // Helper function to check if appointment is all-day (starts at 00:00 or spans entire day)
  const isAllDayAppointment = (apt: IEvent): boolean => {
    const start = new Date(apt.start);
    const end = new Date(apt.end);
    const durationHours = (end.getTime() - start.getTime()) / (1000 * 60 * 60);
    // Consider it all-day if it's 12+ hours or starts at midnight and longer than 4 hours
    return durationHours >= 12 || (start.getHours() === 0 && start.getMinutes() === 0 && durationHours > 4);
  };

  // Check if any day in the week has all-day appointments
  const hasAnyAllDayAppointments = weekDays.some(day => {
    const dayAppointments = appointments.filter(apt => {
      const aptDate = new Date(apt.start);
      return aptDate.getDate() === day.getDate() &&
             aptDate.getMonth() === day.getMonth() &&
             aptDate.getFullYear() === day.getFullYear();
    });
    return dayAppointments.some(isAllDayAppointment);
  });

  return (
    <div className={styles.weekView}>
      <div className={styles.weekGrid} ref={weekGridRef}>
        <div className={styles.timeColumn}>
          <div className={styles.dayHeader} />
          {/* Placeholder for all-day section to keep alignment */}
          {hasAnyAllDayAppointments && (
            <div className={styles.allDayPlaceholder} />
          )}
          <div className={styles.timeColumnSlots}>
            {hours.map(hour => {
              const displayHour = hour === 0 ? 0 : hour;
              const hourStr = displayHour.toString().length === 1 ? `0${displayHour}:00` : `${displayHour}:00`;
              return (
                <div key={hour} className={styles.hourLabel}>
                  {hourStr}
                </div>
              );
            })}
          </div>
        </div>
        {weekDays.map((day, dayIdx) => {
          const dayAppointments = appointments.filter(apt => {
            const aptDate = new Date(apt.start);
            return aptDate.getDate() === day.getDate() &&
                   aptDate.getMonth() === day.getMonth() &&
                   aptDate.getFullYear() === day.getFullYear();
          });
          
          const allDayAppointments = dayAppointments.filter(isAllDayAppointment);
          const timedAppointments = dayAppointments.filter(apt => !isAllDayAppointment(apt));
          
          const isWeekend = day.getDay() === 0 || day.getDay() === 6;
          const isToday = day.getDate() === currentDate.getDate() &&
            day.getMonth() === currentDate.getMonth() &&
            day.getFullYear() === currentDate.getFullYear();
          const dayColumnClassName = `${styles.dayColumn} ${isWeekend ? styles.weekendColumn : ''} ${isToday ? styles.todayColumn : ''}`;
          const dayHeaderClassName = `${styles.dayHeader} ${isToday ? styles.todayHeader : ''}`;

          return (
            <div key={dayIdx} className={dayColumnClassName}>
              <div className={dayHeaderClassName}>
                <div className={styles.dayName}>{day.toLocaleDateString(undefined, { weekday: 'short' })}</div>
                <div className={styles.dayNumber}>{day.getDate()}</div>
              </div>
              {/* Always render all-day section for consistent alignment */}
              {hasAnyAllDayAppointments && (
                <div className={styles.allDaySection}>
                  {allDayAppointments.map(apt => (
                    <div
                      key={apt.id}
                      className={styles.allDayAppointment}
                      style={{
                        backgroundColor: `color-mix(in srgb, ${apt.colorHex ?? '#0078d4'} 20%, transparent)`,
                        borderLeftColor: apt.colorHex ?? '#0078d4'
                      }}
                      title={apt.title}
                    >
                      <div className={styles.appointmentTitle}>
                        {apt.showSourceLogo && apt.sourceType && (
                          <Icon iconName={apt.sourceIconName ?? getSourceIcon(apt.sourceType)} style={{ marginRight: 4, fontSize: 12 }} />
                        )}
                        <span style={{ fontStyle: apt.isDraft ? 'italic' : 'normal' }}>{apt.title}</span>
                      </div>
                    </div>
                  ))}
                </div>
              )}
              <div className={styles.daySlotsWrapper}>
                {/* Absolute layer for hour borders */}
                <div className={styles.hourBorders}>
                  {hours.map(hour => (
                    <div key={`border-${hour}`} className={styles.hourBorderLine} />
                  ))}
                </div>
                {/* Grid for hour slots */}
                <div className={styles.daySlots}>
                  {hours.map(hour => (
                    <div key={hour} className={styles.hourSlot} />
                  ))}
                </div>
                {/* Appointments layer with absolute positioning */}
                <div className={styles.appointmentsLayer}>
                  {timedAppointments.map(apt => {
                    const aptStartDate = new Date(apt.start);
                    const aptEndDate = new Date(apt.end);
                    const aptStartHour = aptStartDate.getHours();
                    const aptStartMinutes = aptStartDate.getMinutes();
                    const aptEndHour = aptEndDate.getHours();
                    const aptEndMinutes = aptEndDate.getMinutes();
                    
                    // Calculate position from the start of visible hours (startHour prop)
                    const minutesFromViewStart = (aptStartHour - startHour) * 60 + aptStartMinutes;
                    const minutesFromViewEnd = (aptEndHour - startHour) * 60 + aptEndMinutes;
                    const durationMinutes = minutesFromViewEnd - minutesFromViewStart;
                    
                    // Position from the top of the day slots using measured hour height
                    const topPosition = minutesFromViewStart * pixelsPerMinute;
                    const height = durationMinutes * pixelsPerMinute;
                    
                    return (
                      <div
                        key={apt.id}
                        className={styles.appointment}
                        style={{
                          position: 'absolute',
                          top: `${topPosition}px`,
                          height: `${height}px`,
                          left: '4px',
                          right: '4px',
                          backgroundColor: `color-mix(in srgb, ${apt.colorHex ?? '#0078d4'} 20%, transparent)`,
                          borderLeftColor: apt.colorHex ?? '#0078d4'
                        }}
                        title={`${apt.title} (${aptStartDate.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })} - ${aptEndDate.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })})`}
                      >
                        <div className={styles.appointmentTitle}>
                          {apt.showSourceLogo && apt.sourceType && (
                            <Icon iconName={apt.sourceIconName ?? getSourceIcon(apt.sourceType)} style={{ marginRight: 4, fontSize: 12 }} />
                          )}
                          <span style={{ fontStyle: apt.isDraft ? 'italic' : 'normal' }}>{apt.title}</span>
                        </div>
                        <div className={styles.appointmentTime}>
                          {aptStartDate.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};

