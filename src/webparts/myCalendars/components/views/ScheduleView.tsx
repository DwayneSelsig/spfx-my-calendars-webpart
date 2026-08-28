import * as React from 'react';
import type { ICalendarEvent as IEvent } from '../../models/ICalendarEvent';
import { Icon } from '@fluentui/react/lib/Icon';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { getSourceIconName } from '../../utils/sourceIconHelper';

interface IScheduleViewProps {
  appointments: IEvent[];
  currentDate: Date;
  isLoading: boolean;
}

const styles = mergeStyleSets({
  scheduleView: {
    width: '100%',
    height: '100%',
    padding: 16,
    display: 'flex',
    flexDirection: 'column',
    boxSizing: 'border-box',
    maxWidth: 900,
    margin: '0 auto'
  },
  scheduleDay: {
    marginBottom: 32
  },
  scheduleDate: {
    display: 'flex',
    alignItems: 'center',
    gap: 12,
    marginBottom: 16,
    paddingBottom: 12,
    borderBottom: '2px solid var(--neutralLight, #edebe9)'
  },
  scheduleDayName: {
    fontSize: 18,
    fontWeight: 600,
    color: 'var(--neutralPrimary, #323130)'
  },
  scheduleAppointments: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: 12
  },
  scheduleAppointment: {
    padding: '14px 16px',
    borderRadius: 6,
    borderLeft: '5px solid',
    backgroundColor: 'var(--white, #ffffff)',
    boxSizing: 'border-box'
  },
  appointmentTitle: {
    fontSize: 14,
    fontWeight: 600,
    color: 'var(--neutralPrimary, #323130)',
    marginBottom: 4,
    display: 'flex',
    alignItems: 'center',
    gap: 4,
    minWidth: 0
  },
  appointmentTime: {
    fontSize: 12,
    color: 'var(--neutralSecondary, #605e5c)',
    marginBottom: 4
  },
  noAppointments: {
    textAlign: 'center',
    padding: '40px 20px',
    color: 'var(--neutralSecondary, #605e5c)',
    fontSize: 14
  }
});

export const ScheduleView: React.FC<IScheduleViewProps> = (props) => {
  const { appointments, currentDate, isLoading } = props;

  const getDayAppointments = (): IEvent[] => {
    return appointments
      .filter(apt => {
        const aptDate = new Date(apt.start);
        return aptDate.getDate() === currentDate.getDate() &&
               aptDate.getMonth() === currentDate.getMonth() &&
               aptDate.getFullYear() === currentDate.getFullYear();
      })
      .sort((a, b) => new Date(a.start).getTime() - new Date(b.start).getTime());
  };

  const dayAppointments = getDayAppointments();

  const getAppointmentDuration = (apt: IEvent): string => {
    if (apt.isFullDay) {
      return 'All day';
    }
    const durationMs = new Date(apt.end).getTime() - new Date(apt.start).getTime();
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
                backgroundColor: `color-mix(in srgb, ${apt.colorHex ?? '#0078d4'} 12%, transparent)`,
                borderLeftColor: apt.colorHex ?? '#0078d4'
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
                  {apt.isFullDay ? '09:00' : 
                    new Date(apt.start).toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit', hour12: false })
                  }
                </div>
                <div style={{ flex: 1 }}>
                  <div className={styles.appointmentTitle}>
                    {apt.showSourceLogo !== false && (
                      <Icon iconName={getSourceIconName(apt.sourceType, apt.sourceIconName)} style={{ marginRight: 4, fontSize: 12 }} />
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

