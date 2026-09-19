import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type { ICalendarEvent } from '../../models/ICalendarEvent';
import { getAvailabilityLabel, getCalendarLabels, getResponseLabel } from './calendarLabels';
import { getEventStatusPresentation } from './eventStatusPresentation';

export interface IEventStatusGlyphsProps {
  event: ICalendarEvent;
  color: string;
}

export function getEventStatusAriaText(event: ICalendarEvent): string {
  const labels = getCalendarLabels();
  return [
    event.showAs ? `${labels.calendarAvailability}: ${getAvailabilityLabel(event.showAs)}` : undefined,
    event.responseStatus ? `${labels.yourResponse}: ${getResponseLabel(event.responseStatus)}` : undefined
  ].filter(Boolean).join(', ');
}

export const EventStatusGlyphs: React.FC<IEventStatusGlyphsProps> = ({ event, color }) => {
  const labels = getCalendarLabels();
  const presentation = getEventStatusPresentation(event, color);
  return (
    <>
      {event.showAs && presentation.availabilityIconName && (
        <span role="img" aria-label={`${labels.calendarAvailability}: ${getAvailabilityLabel(event.showAs)}`} title={`${labels.calendarAvailability}: ${getAvailabilityLabel(event.showAs)}`} style={{ marginRight: 4 }}>
          <Icon iconName={presentation.availabilityIconName} aria-hidden="true" />
        </span>
      )}
      {event.responseStatus && presentation.responseIconName && (
        <span role="img" aria-label={`${labels.yourResponse}: ${getResponseLabel(event.responseStatus)}`} title={`${labels.yourResponse}: ${getResponseLabel(event.responseStatus)}`} style={{ marginRight: 4 }}>
          <Icon iconName={presentation.responseIconName} aria-hidden="true" />
        </span>
      )}
    </>
  );
};
