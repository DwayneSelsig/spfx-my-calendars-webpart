import * as React from 'react';
import { Shimmer, ShimmerElementType } from '@fluentui/react/lib/Shimmer';
import styles from './CalendarView.module.scss';
import { ICalendarViewProps } from './DayView';

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
  for (let h = startHour; h <= endHour; h++) {
    hours.push(h);
  }

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
          {hours.map(hour => (
            <div key={hour} className={styles.hourLabel}>
              {hour.toString().length === 1 ? `0${hour}:00` : `${hour}:00`}
            </div>
          ))}
        </div>
        {weekDays.map((day, dayIdx) => {
          const dayAppointments = appointments.filter(apt => {
            const aptDate = new Date(apt.startDate);
            return aptDate.getDate() === day.getDate() &&
                   aptDate.getMonth() === day.getMonth() &&
                   aptDate.getFullYear() === day.getFullYear();
          });
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
              <div className={styles.daySlots}>
                {hours.map(hour => (
                  <div key={hour} className={styles.hourSlot}>
                    {dayAppointments
                      .filter(apt => {
                        const aptHour = apt.startDate.getHours();
                        return aptHour === hour;
                      })
                      .map(apt => (
                        <div
                          key={apt.id}
                          className={styles.appointment}
                          style={{
                            backgroundColor: `color-mix(in srgb, ${apt.color} 20%, transparent)`,
                            borderLeftColor: apt.color
                          }}
                        >
                          <div className={styles.appointmentTitle}>{apt.title}</div>
                        </div>
                      ))}
                  </div>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};
