import * as React from 'react';
import { IAppointment } from '../../models/IAppointment';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CalendarView.module.scss';
import { ICalendarViewProps } from './DayView';
import { getSourceIcon } from '../../utils/sourceIconHelper';

export const ScheduleView: React.FC<ICalendarViewProps> = (props) => {
  const { appointments, currentDate, isLoading } = props;

  const getDayAppointments = (): IAppointment[] => {
    return appointments
      .filter(apt => {
        const aptDate = new Date(apt.startDate);
        return aptDate.getDate() === currentDate.getDate() &&
               aptDate.getMonth() === currentDate.getMonth() &&
               aptDate.getFullYear() === currentDate.getFullYear();
      })
      .sort((a, b) => a.startDate.getTime() - b.startDate.getTime());
  };

  const dayAppointments = getDayAppointments();

  const getAppointmentDuration = (apt: IAppointment): string => {
    if (apt.isAllDay) {
      return 'All day';
    }
    const durationMs = apt.endDate.getTime() - apt.startDate.getTime();
    const hours = Math.floor(durationMs / (1000 * 60 * 60));
    const minutes = Math.floor((durationMs % (1000 * 60 * 60)) / (1000 * 60));
    if (hours > 0) {
      return `${hours} hr${hours > 1 ? 's' : ''}${minutes > 0 ? ` ${minutes} min` : ''}`;
    }
    return `${minutes} min`;
  };

  const dayName = currentDate.toLocaleDateString(undefined, { weekday: 'long' });
  const dateStr = currentDate.toLocaleDateString(undefined, { day: 'numeric', month: 'long' });

  return (
    <div className={styles.scheduleView}>
      <div className={styles.scheduleDay}>
        <div className={styles.scheduleDate}>
          <div className={styles.scheduleDayName}>
            {dayName}<span style={{ marginLeft: 8, fontSize: 12, fontWeight: 400 }}>• {dateStr}</span>
          </div>
        </div>
        <div className={styles.scheduleAppointments}>
          {dayAppointments.map(apt => (
            <div
              key={apt.id}
              className={styles.scheduleAppointment}
              style={{
                backgroundColor: `color-mix(in srgb, ${apt.color} 12%, transparent)`,
                borderLeftColor: apt.color
              }}
            >
              <div style={{ display: 'flex', gap: 12, alignItems: 'flex-start' }}>
                <div style={{ 
                  fontSize: 12, 
                  fontWeight: 600,
                  color: 'var(--neutralPrimary, #323130)',
                  whiteSpace: 'nowrap',
                  minWidth: 45
                }}>
                  {apt.isAllDay ? '09:00' : 
                    apt.startDate.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit', hour12: false })
                  }
                </div>
                <div style={{ flex: 1 }}>
                  <div className={styles.appointmentTitle}>
                    {apt.showSourceLogo && apt.sourceType && (
                      <Icon iconName={getSourceIcon(apt.sourceType)} style={{ marginRight: 4, fontSize: 12 }} />
                    )}
                    <span style={{ fontStyle: apt.isDraft ? 'italic' : 'normal' }}>{apt.title}</span>
                  </div>
                  <div className={styles.appointmentTime}>
                    {getAppointmentDuration(apt)}
                  </div>
                </div>
              </div>
              <div style={{ marginTop: 8, paddingLeft: 57 }}>
                {apt.location && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: 'var(--neutralSecondary, #605e5c)', marginBottom: 4 }}>
                    <span style={{ fontSize: 13 }}>📍</span>
                    {apt.location}
                  </div>
                )}
                {apt.organizer && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: 'var(--neutralSecondary, #605e5c)' }}>
                    <span style={{ fontSize: 13 }}>👤</span>
                    {apt.organizer}
                  </div>
                )}
              </div>
            </div>
          ))}
        </div>
      </div>
      {dayAppointments.length === 0 && !isLoading && (
        <div className={styles.noAppointments}>No appointments today</div>
      )}
    </div>
  );
};
