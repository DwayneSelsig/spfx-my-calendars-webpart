import * as React from 'react';
import { Shimmer, ShimmerElementType } from '@fluentui/react/lib/Shimmer';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CalendarView.module.scss';
import { ICalendarViewProps } from './DayView';
import { getSourceIcon } from '../../utils/sourceIconHelper';

export const MonthView: React.FC<ICalendarViewProps> = (props) => {
  const { appointments, currentDate, isLoading, onDateChange } = props;

  const getDaysInMonth = (): Date[] => {
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth();
    const firstDay = new Date(year, month, 1);
    
    const days: Date[] = [];
    const startDay = firstDay.getDay();
    const offset = startDay === 0 ? -6 : 1 - startDay;
    
    const startDate = new Date(firstDay);
    startDate.setDate(firstDay.getDate() + offset);
    
    for (let i = 0; i < 42; i++) {
      const date = new Date(startDate);
      date.setDate(startDate.getDate() + i);
      days.push(date);
    }
    
    return days;
  };

  const days = getDaysInMonth();
  const weekDays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

  if (isLoading) {
    return (
      <div className={styles.monthView}>
        {days.map((_, idx) => (
          <Shimmer key={idx} className={styles.shimmerDay} shimmerElements={[
            { type: ShimmerElementType.line, height: 100 }
          ]} />
        ))}
      </div>
    );
  }

  return (
    <div className={styles.monthView}>
      <div className={styles.monthGrid}>
        {weekDays.map(day => (
          <div key={day} className={styles.weekDayHeader}>{day}</div>
        ))}
        {days.map((day, idx) => {
          const isCurrentMonth = day.getMonth() === currentDate.getMonth();
          const dayAppointments = appointments.filter(apt => {
            const aptDate = new Date(apt.startDate);
            return aptDate.getDate() === day.getDate() &&
                   aptDate.getMonth() === day.getMonth() &&
                   aptDate.getFullYear() === day.getFullYear();
          });

          return (
            <div
              key={idx}
              className={`${styles.dayCell} ${!isCurrentMonth ? styles.otherMonth : ''}`}
              onClick={() => onDateChange(day)}
            >
              <div className={styles.dayNumber}>{day.getDate()}</div>
              <div className={styles.dayAppointments}>
                {dayAppointments.slice(0, 3).map(apt => (
                  <div
                    key={apt.id}
                    className={styles.monthAppointment}
                    style={{
                      backgroundColor: `color-mix(in srgb, ${apt.color} 20%, transparent)`,
                      borderLeftColor: apt.color
                    }}
                  >
                    {apt.showSourceLogo && apt.sourceType && (
                      <Icon iconName={getSourceIcon(apt.sourceType)} style={{ marginRight: 4, fontSize: 12 }} />
                    )}
                    <span style={{ fontStyle: apt.isDraft ? 'italic' : 'normal' }}>{apt.title}</span>
                  </div>
                ))}
                {dayAppointments.length > 3 && (
                  <div className={styles.moreAppointments}>
                    +{dayAppointments.length - 3} more
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};
