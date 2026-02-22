import * as React from 'react';
import { Shimmer, ShimmerElementType } from '@fluentui/react/lib/Shimmer';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CalendarView.module.scss';
import { ICalendarViewProps } from './DayView';
import { getSourceIcon } from '../../utils/sourceIconHelper';
import { IAppointment } from '../../models/IAppointment';

export const WeekView: React.FC<ICalendarViewProps> = (props) => {
  const { appointments, currentDate, isLoading, startHour, endHour, showWeekends } = props;

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

  // Helper function to check if appointment is all-day (starts at 00:00 or spans entire day)
  const isAllDayAppointment = (apt: IAppointment): boolean => {
    if (!apt.endDate) return false;
    const start = apt.startDate; // Already a Date object
    const end = apt.endDate; // Already a Date object
    const durationHours = (end.getTime() - start.getTime()) / (1000 * 60 * 60);
    // Consider it all-day if it's 12+ hours or starts at midnight and longer than 4 hours
    return durationHours >= 12 || (start.getHours() === 0 && start.getMinutes() === 0 && durationHours > 4);
  };

  // Check if any day in the week has all-day appointments
  const hasAnyAllDayAppointments = weekDays.some(day => {
    const dayAppointments = appointments.filter(apt => {
      const aptDate = apt.startDate; // Already a Date object
      return aptDate.getDate() === day.getDate() &&
             aptDate.getMonth() === day.getMonth() &&
             aptDate.getFullYear() === day.getFullYear();
    });
    return dayAppointments.some(isAllDayAppointment);
  });

  if (isLoading) {
    return (
      <div className={styles.weekView}>
        {weekDays.map((day, idx) => (
          <div key={idx} className={styles.dayColumn}>
            <Shimmer className={styles.shimmerHeader} />
            {hours.map(hour => (
              <Shimmer key={hour} className={styles.shimmerLine} shimmerElements={[
                { type: ShimmerElementType.line, height: 60 }
              ]} />
            ))}
          </div>
        ))}
      </div>
    );
  }

  return (
    <div className={styles.weekView}>
      <div className={styles.weekGrid}>
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
            const aptDate = apt.startDate; // Already a Date object
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
                        backgroundColor: `color-mix(in srgb, ${apt.color} 20%, transparent)`,
                        borderLeftColor: apt.color
                      }}
                      title={apt.title}
                    >
                      <div className={styles.appointmentTitle}>
                        {apt.showSourceLogo && apt.sourceType && (
                          <Icon iconName={getSourceIcon(apt.sourceType)} style={{ marginRight: 4, fontSize: 12 }} />
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
                    const aptStartHour = apt.startDate.getHours();
                    const aptStartMinutes = apt.startDate.getMinutes();
                    const aptEndHour = apt.endDate ? apt.endDate.getHours() : aptStartHour + 1;
                    const aptEndMinutes = apt.endDate ? apt.endDate.getMinutes() : 0;
                    
                    // Calculate position from the start of visible hours (startHour prop)
                    const minutesFromViewStart = (aptStartHour - startHour) * 60 + aptStartMinutes;
                    const minutesFromViewEnd = (aptEndHour - startHour) * 60 + aptEndMinutes;
                    const durationMinutes = minutesFromViewEnd - minutesFromViewStart;
                    
                    // Position from the top of the day slots (60px per hour)
                    const topPosition = minutesFromViewStart * (60 / 60); // 60px per hour = 1px per minute
                    const height = durationMinutes * (60 / 60); // 60px per hour = 1px per minute
                    
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
                          backgroundColor: `color-mix(in srgb, ${apt.color} 20%, transparent)`,
                          borderLeftColor: apt.color
                        }}
                        title={`${apt.title} (${apt.startDate.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })} - ${apt.endDate?.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })})`}
                      >
                        <div className={styles.appointmentTitle}>
                          {apt.showSourceLogo && apt.sourceType && (
                            <Icon iconName={getSourceIcon(apt.sourceType)} style={{ marginRight: 4, fontSize: 12 }} />
                          )}
                          <span style={{ fontStyle: apt.isDraft ? 'italic' : 'normal' }}>{apt.title}</span>
                        </div>
                        <div className={styles.appointmentTime}>
                          {apt.startDate.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
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
