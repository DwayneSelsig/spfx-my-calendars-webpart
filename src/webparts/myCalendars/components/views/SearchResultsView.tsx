import * as React from 'react';
import { IAppointment } from '../../models/IAppointment';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CalendarView.module.scss';
import { getSourceIcon } from '../../utils/sourceIconHelper';

export interface ISearchResultsViewProps {
  appointments: IAppointment[];
  isLoading: boolean;
  searchQuery: string;
}

interface IGroupedAppointments {
  date: Date;
  appointments: IAppointment[];
}

export const SearchResultsView: React.FC<ISearchResultsViewProps> = (props) => {
  const { appointments, isLoading, searchQuery } = props;

  // Helper function to pad numbers
  const padZero = (value: number): string => value < 10 ? '0' + value : String(value);

  // Group appointments by date
  const groupedByDate = React.useMemo(() => {
    const groups: { [key: string]: IGroupedAppointments } = {};

    appointments.forEach((apt: IAppointment) => {
      const aptDate = new Date(apt.startDate);
      const month = padZero(aptDate.getMonth() + 1);
      const day = padZero(aptDate.getDate());
      const dateKey = `${aptDate.getFullYear()}-${month}-${day}`;
      
      if (!groups[dateKey]) {
        groups[dateKey] = {
          date: new Date(aptDate.getFullYear(), aptDate.getMonth(), aptDate.getDate()),
          appointments: []
        };
      }
      
      groups[dateKey].appointments.push(apt);
    });

    // Sort by date and each group's appointments by start time
    return Object.keys(groups)
      .sort()
      .map((key: string) => groups[key])
      .map((group: IGroupedAppointments) => ({
        ...group,
        appointments: group.appointments.sort((a: IAppointment, b: IAppointment) => a.startDate.getTime() - b.startDate.getTime())
      }));
  }, [appointments]);

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

  if (groupedByDate.length === 0 && !isLoading) {
    return (
      <div className={styles.scheduleView}>
        <div className={styles.noAppointments}>
          {searchQuery ? `No results found for "${searchQuery}"` : 'No appointments'}
        </div>
      </div>
    );
  }

  return (
    <div className={styles.scheduleView}>
      {groupedByDate.map((group: IGroupedAppointments) => {
        const dayName = group.date.toLocaleDateString(undefined, { weekday: 'long' });
        const dateStr = group.date.toLocaleDateString(undefined, { day: 'numeric', month: 'long' });

        return (
          <div key={`${group.date.getFullYear()}-${group.date.getMonth()}-${group.date.getDate()}`} className={styles.scheduleDay}>
            <div className={styles.scheduleDate}>
              <div className={styles.scheduleDayName}>
                {dayName}<span style={{ marginLeft: 8, fontSize: 12, fontWeight: 400 }}>• {dateStr}</span>
              </div>
            </div>
            <div className={styles.scheduleAppointments}>
              {group.appointments.map((apt: IAppointment) => (
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
                          <Icon iconName={apt.sourceIconName ?? getSourceIcon(apt.sourceType)} style={{ marginRight: 4, fontSize: 12 }} />
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
        );
      })}
    </div>
  );
};
