import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { Icon } from '@fluentui/react/lib/Icon';
import type { ICalendarEvent } from '../../models/ICalendarEvent';
import { getSourceIconName, getSourceTypeDisplayName } from '../../utils/sourceIconHelper';
import { getCalendarLabels } from './calendarLabels';
import { getCalendarColor, safeOpen } from './calendarUtils';

export interface IEventDetailsDialogProps {
  event?: ICalendarEvent;
  onDismiss: () => void;
}

export const EventDetailsDialog: React.FC<IEventDetailsDialogProps> = ({ event, onDismiss }) => {
  const labels = getCalendarLabels();
  if (!event) return null;
  const start = new Date(event.start);
  const end = new Date(event.end);
  const dateTime = event.isFullDay
    ? `${start.toLocaleDateString()} · ${labels.allDay}`
    : `${start.toLocaleString()} – ${end.toLocaleString()}`;
  const organizer = event.organizer?.name || event.organizer?.email;
  const attendees = (event.attendees || []).map(attendee => attendee.name || attendee.email).filter(Boolean).join(', ');
  const sourceTypeDisplayName = getSourceTypeDisplayName(event.sourceType, event.sourceIconName);
  const sourceDisplayName = event.sourceDisplayName && event.sourceDisplayName !== sourceTypeDisplayName
    ? `${event.sourceDisplayName} · ${sourceTypeDisplayName}`
    : sourceTypeDisplayName;

  return (
    <Dialog
      hidden={false}
      onDismiss={onDismiss}
      minWidth={360}
      maxWidth={560}
      dialogContentProps={{ type: DialogType.normal, title: event.title, subText: dateTime, closeButtonAriaLabel: labels.close }}
      modalProps={{ isBlocking: false, styles: { main: { borderTop: `4px solid ${getCalendarColor(event)}` } } }}
    >
      <div style={{ display: 'grid', rowGap: 10 }}>
        {event.location && <div><Icon iconName="POI" /> <strong>{labels.location}:</strong> {event.location}</div>}
        {organizer && <div><Icon iconName="Contact" /> <strong>{labels.organizer}:</strong> {organizer}</div>}
        {attendees && <div><Icon iconName="People" /> <strong>{labels.attendees}:</strong> {attendees}</div>}
        {event.description && <div><strong>{labels.description}</strong><div style={{ whiteSpace: 'pre-wrap', marginTop: 4 }}>{event.description}</div></div>}
        {event.sourceType && (
          <div>
            {event.showSourceLogo !== false && <Icon iconName={getSourceIconName(event.sourceType, event.sourceIconName)} style={{ marginRight: 4 }} />}
            <strong>{labels.source}:</strong> {sourceDisplayName}
          </div>
        )}
      </div>
      <DialogFooter>
        {event.joinUrl && <PrimaryButton text={labels.join} iconProps={{ iconName: 'Video' }} onClick={() => safeOpen(event.joinUrl as string)} />}
        {event.webLink && <DefaultButton text={labels.open} iconProps={{ iconName: 'OpenInNewWindow' }} onClick={() => safeOpen(event.webLink as string)} />}
        <DefaultButton text={labels.close} onClick={onDismiss} />
      </DialogFooter>
    </Dialog>
  );
};
