import * as React from 'react';
import { IAppointment } from '../../models/IAppointment';
import { Shimmer, ShimmerElementType } from '@fluentui/react/lib/Shimmer';
import styles from './CalendarView.module.scss';

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
  const { appointments, currentDate, isLoading, startHour, endHour } = props;
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
              <div className={styles.hourSlot}>
                {dayAppointments
                  .filter(apt => {
                    const aptHour = new Date(apt.startDate).getHours();
                    const aptEndHour = new Date(apt.endDate).getHours();
                    return aptHour === hour || (aptHour < hour && aptEndHour > hour);
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
                      {apt.location && <div className={styles.appointmentLocation}>{apt.location}</div>}
                    </div>
                  ))}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};
