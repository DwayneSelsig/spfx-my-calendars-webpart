import * as strings from 'MyCalendarsWebPartStrings';
import { formatLocalizedString } from '../../utils/localization';

export function formatAppointmentDuration(startDate: Date, endDate: Date, isFullDay: boolean): string {
  if (isFullDay) return strings.AllDayLabel;

  const durationMs = endDate.getTime() - startDate.getTime();
  const hours = Math.floor(durationMs / (1000 * 60 * 60));
  const minutes = Math.floor((durationMs % (1000 * 60 * 60)) / (1000 * 60));

  if (hours > 0) {
    const hourLabel = formatLocalizedString(strings.DurationHoursLabel, hours);
    return minutes > 0
      ? formatLocalizedString(strings.DurationHoursMinutesLabel, hours, minutes)
      : hourLabel;
  }

  return formatLocalizedString(strings.DurationMinutesLabel, minutes);
}
