import * as React from 'react';
import { IAppointment } from '../../models/IAppointment';
import { Shimmer, ShimmerElementType } from '@fluentui/react/lib/Shimmer';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CalendarView.module.scss';
import { getSourceIcon } from '../../utils/sourceIconHelper';

export interface ICalendarViewProps {
  appointments: IAppointment[];
  currentDate: Date;
  onDateChange: (date: Date) => void;
  isLoading: boolean;
  startHour: number;
  endHour: number;
  showWeekends: boolean;
  slotDuration: number;
}

export const DayView: React.FC<ICalendarViewProps> = (props) => {
  const { appointments, currentDate, isLoading, startHour } = props;
  const scrollContainerRef = React.useRef<HTMLDivElement>(null);

  // Render full day (0-24) to allow scrolling
  const hours: number[] = [];
  for (let h = 0; h < 24; h++) {
    hours.push(h);
  }

  // Scroll to show startHour after component renders
  React.useLayoutEffect(() => {
    if (scrollContainerRef.current) {
      // Use setTimeout to ensure DOM is ready
      const timer = setTimeout(() => {
        if (scrollContainerRef.current) {
          // Measure the actual height of one hour row
          const firstRow = scrollContainerRef.current.querySelector('[data-hour="0"]') as HTMLElement;
          let hourRowHeight = 60;
          
          if (firstRow) {
            hourRowHeight = firstRow.offsetHeight;
          }
          
          const scrollPosition = startHour * hourRowHeight;
          scrollContainerRef.current.scrollTop = scrollPosition;
        }
      }, 0);
      
      return () => clearTimeout(timer);
    }
  }, [startHour]);

  const dayAppointments = appointments.filter(apt => {
    const aptDate = new Date(apt.startDate);
    return aptDate.getDate() === currentDate.getDate() &&
           aptDate.getMonth() === currentDate.getMonth() &&
           aptDate.getFullYear() === currentDate.getFullYear();
  });

  if (isLoading) {
    return (
      <div className={styles.dayView}>
        <div className={styles.timeGridContainer} ref={scrollContainerRef}>
          {hours.map(hour => (
            <Shimmer key={hour} className={styles.shimmerLine} shimmerElements={[
              { type: ShimmerElementType.line, height: 60, width: '100%' }
            ]} />
          ))}
        </div>
      </div>
    );
  }

  return (
    <div className={styles.dayView}>
      <div className={styles.timeGridContainer} ref={scrollContainerRef}>
        <div className={styles.timeGrid}>
          {hours.map(hour => (
            <div key={hour} className={styles.hourRow} data-hour={hour}>
              <div className={styles.hourLabel}>
                {hour.toString().length === 1 ? `0${hour}:00` : `${hour}:00`}
              </div>
              <div className={styles.hourSlot} />
            </div>
          ))}
          {/* Appointments layer with absolute positioning */}
          <div className={styles.appointmentsLayerDay}>
            {dayAppointments.map(apt => {
              const aptStartDate = new Date(apt.startDate);
              const aptEndDate = new Date(apt.endDate);
              const aptStartHour = aptStartDate.getHours();
              const aptStartMinutes = aptStartDate.getMinutes();
              const aptEndHour = aptEndDate.getHours();
              const aptEndMinutes = aptEndDate.getMinutes();
              
              // Calculate position from midnight (hour 0)
              const minutesFromMidnight = aptStartHour * 60 + aptStartMinutes;
              const minutesFromMidnightEnd = aptEndHour * 60 + aptEndMinutes;
              const durationMinutes = minutesFromMidnightEnd - minutesFromMidnight;
              
              // Position from the top (61px per hour due to 1px border = 61/60 px per minute)
              const topPosition = minutesFromMidnight * (61 / 60);
              const height = durationMinutes * (61 / 60);
              
              return (
                <div
                  key={apt.id}
                  className={styles.appointment}
                  style={{
                    position: 'absolute',
                    top: `${topPosition}px`,
                    height: `${height}px`,
                    left: '88px', // Account for hourLabel width (60px) + padding + 20px
                    right: '8px',
                    backgroundColor: `color-mix(in srgb, ${apt.color} 20%, transparent)`,
                    borderLeftColor: apt.color,
                    margin: 0
                  }}
                  title={`${apt.title} (${aptStartDate.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })} - ${aptEndDate.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })})`}
                >
                  <div className={styles.appointmentTitle}>
                    {apt.showSourceLogo && apt.sourceType && (
                      <Icon iconName={getSourceIcon(apt.sourceType)} style={{ marginRight: 4, fontSize: 12 }} />
                    )}
                    {apt.title}
                  </div>
                  <div className={styles.appointmentTime}>
                    {aptStartDate.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                  </div>
                  {apt.location && <div className={styles.appointmentLocation}>{apt.location}</div>}
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
};
