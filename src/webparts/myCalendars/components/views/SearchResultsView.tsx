import * as React from 'react';
import { IEvent } from '@pnp/spfx-controls-react/lib/controls/calendar/models/IEvents';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CalendarView.module.scss';
import { getSourceIcon } from '../../utils/sourceIconHelper';

export interface ISearchResultsViewProps {
  appointments: IEvent[];
  isLoading: boolean;
  searchQuery: string;
}

interface IGroupedAppointments {
  date: Date;
  dayName: string;
  dateLabel: string;
  appointments: ISearchResultAppointment[];
}

interface ISearchResultAppointment {
  appointment: IEvent;
  startTimeLabel: string;
  durationLabel: string;
}

export const SearchResultsView: React.FC<ISearchResultsViewProps> = (props) => {
  const { appointments, isLoading, searchQuery } = props;

  const weekdayFormatter = React.useMemo(
    () => new Intl.DateTimeFormat(undefined, { weekday: 'long' }),
    []
  );
  const dateFormatter = React.useMemo(
    () => new Intl.DateTimeFormat(undefined, { day: 'numeric', month: 'long' }),
    []
  );
  const timeFormatter = React.useMemo(
    () => new Intl.DateTimeFormat(undefined, { hour: '2-digit', minute: '2-digit', hour12: false }),
    []
  );

  // Helper function to pad numbers
  const padZero = (value: number): string => value < 10 ? '0' + value : String(value);

  const getAppointmentDuration = React.useCallback((startDate: Date, endDate: Date, isFullDay: boolean): string => {
    if (isFullDay) {
      return 'All day';
    }

    const durationMs = endDate.getTime() - startDate.getTime();
    const hours = Math.floor(durationMs / (1000 * 60 * 60));
    const minutes = Math.floor((durationMs % (1000 * 60 * 60)) / (1000 * 60));
    if (hours > 0) {
      return `${hours} hr${hours > 1 ? 's' : ''}${minutes > 0 ? ` ${minutes} min` : ''}`;
    }

    return `${minutes} min`;
  }, []);

  // Group appointments by date and precompute display labels to keep render work minimal.
  const groupedByDate = React.useMemo(() => {
    const groups: { [key: string]: IGroupedAppointments } = {};
    const timeLabelCache = new Map<number, string>();

    appointments.forEach((apt: IEvent) => {
      const startDate = new Date(apt.start);
      const endDate = new Date(apt.end);
      const month = padZero(startDate.getMonth() + 1);
      const day = padZero(startDate.getDate());
      const dateKey = `${startDate.getFullYear()}-${month}-${day}`;

      if (!groups[dateKey]) {
        const groupDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
        groups[dateKey] = {
          date: groupDate,
          dayName: weekdayFormatter.format(groupDate),
          dateLabel: dateFormatter.format(groupDate),
          appointments: []
        };
      }

      let startTimeLabel = '09:00';
      if (!apt.isFullDay) {
        const startTimeKey = startDate.getTime();
        const cachedStartTime = timeLabelCache.get(startTimeKey);
        if (cachedStartTime) {
          startTimeLabel = cachedStartTime;
        } else {
          startTimeLabel = timeFormatter.format(startDate);
          timeLabelCache.set(startTimeKey, startTimeLabel);
        }
      }

      groups[dateKey].appointments.push({
        appointment: apt,
        startTimeLabel,
        durationLabel: getAppointmentDuration(startDate, endDate, apt.isFullDay ?? false)
      });
    });

    // Sort by date and each group's appointments by start time.
    return Object.keys(groups)
      .sort()
      .map((key: string) => groups[key])
      .map((group: IGroupedAppointments) => ({
        ...group,
        appointments: group.appointments.sort(
          (a: ISearchResultAppointment, b: ISearchResultAppointment) =>
            new Date(a.appointment.start).getTime() - new Date(b.appointment.start).getTime()
        )
      }));
  }, [appointments, dateFormatter, getAppointmentDuration, timeFormatter, weekdayFormatter]);

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
        return (
          <div key={`${group.date.getFullYear()}-${group.date.getMonth()}-${group.date.getDate()}`} className={styles.scheduleDay}>
            <div className={styles.scheduleDate}>
              <div className={styles.scheduleDayName}>
                {group.dayName}<span style={{ marginLeft: 8, fontSize: 12, fontWeight: 400 }}>• {group.dateLabel}</span>
              </div>
            </div>
            <div className={styles.scheduleAppointments}>
              {group.appointments.map((result: ISearchResultAppointment) => {
                const apt = result.appointment;
                return (
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
                      {result.startTimeLabel}
                    </div>
                    <div style={{ flex: 1 }}>
                      <div className={styles.appointmentTitle}>
                        {apt.showSourceLogo && apt.sourceType && (
                          <Icon iconName={apt.sourceIconName ?? getSourceIcon(apt.sourceType)} style={{ marginRight: 4, fontSize: 12 }} />
                        )}
                        <span style={{ fontStyle: apt.isDraft ? 'italic' : 'normal' }}>{apt.title}</span>
                      </div>
                      <div className={styles.appointmentTime}>
                        {result.durationLabel}
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
              );})}
            </div>
          </div>
        );
      })}
    </div>
  );
};

