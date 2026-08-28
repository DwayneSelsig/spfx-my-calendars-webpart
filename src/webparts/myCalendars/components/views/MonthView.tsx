import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { ICalendarViewProps } from './DayView';
import { getSourceIcon } from '../../utils/sourceIconHelper';

const styles = mergeStyleSets({
  monthView: {
    width: '100%',
    height: '100%',
    padding: 16,
    display: 'flex',
    flexDirection: 'column',
    boxSizing: 'border-box'
  },
  monthGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(7, 1fr)',
    gap: 1,
    backgroundColor: 'var(--neutralLight, #edebe9)',
    border: '1px solid var(--neutralLight, #edebe9)'
  },
  weekDayHeader: {
    padding: '12px 8px',
    textAlign: 'center',
    backgroundColor: 'var(--neutralLighter, #f3f2f1)',
    fontWeight: 600,
    fontSize: 12,
    color: 'var(--neutralPrimary, #323130)',
    textTransform: 'uppercase'
  },
  dayCell: {
    backgroundColor: 'var(--white, #ffffff)',
    minHeight: 100,
    minWidth: 0,
    padding: 8,
    cursor: 'pointer',
    boxSizing: 'border-box'
  },
  otherMonth: {
    opacity: 0.5
  },
  dayNumber: {
    fontSize: 14,
    fontWeight: 600,
    marginBottom: 4,
    color: 'var(--neutralPrimary, #323130)'
  },
  dayAppointments: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: 2
  },
  monthAppointment: {
    fontSize: 11,
    padding: '2px 6px',
    borderRadius: 2,
    borderLeft: '3px solid',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap'
  },
  moreAppointments: {
    fontSize: 11,
    color: 'var(--themePrimary, #0078d4)',
    padding: '2px 6px',
    cursor: 'pointer'
  }
});

export const MonthView: React.FC<ICalendarViewProps> = (props) => {
  const { appointments, currentDate, onDateChange } = props;

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

  return (
    <div className={styles.monthView}>
      <div className={styles.monthGrid}>
        {weekDays.map(day => (
          <div key={day} className={styles.weekDayHeader}>{day}</div>
        ))}
        {days.map((day, idx) => {
          const isCurrentMonth = day.getMonth() === currentDate.getMonth();
          const dayAppointments = appointments.filter(apt => {
            const aptDate = new Date(apt.start);
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
                      backgroundColor: `color-mix(in srgb, ${apt.colorHex ?? '#0078d4'} 20%, transparent)`,
                      borderLeftColor: apt.colorHex ?? '#0078d4'
                    }}
                  >
                    {apt.showSourceLogo && apt.sourceType && (
                      <Icon iconName={apt.sourceIconName ?? getSourceIcon(apt.sourceType)} style={{ marginRight: 4, fontSize: 12 }} />
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
