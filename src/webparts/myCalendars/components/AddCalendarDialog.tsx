import * as React from 'react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { Label } from '@fluentui/react/lib/Label';
import { CalendarSourceType } from '../models/ICalendarSettings';

export interface IAddCalendarDialogProps {
  isOpen: boolean;
  onDismiss: () => void;
  onSelectType: (type: CalendarSourceType) => void;
}

export class AddCalendarDialog extends React.Component<IAddCalendarDialogProps> {
  public render(): React.ReactElement {
    const { isOpen, onDismiss, onSelectType } = this.props;

    const handleSelectType = (type: CalendarSourceType): void => {
      onSelectType(type);
      // Don't close - parent will show next step
    };

    return (
      <Dialog
        hidden={!isOpen}
        onDismiss={onDismiss}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Add Calendar',
          closeButtonAriaLabel: 'Close'
        }}
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <Label>Choose the type of calendar to add:</Label>
          
          <Stack.Item>
            <PrimaryButton
              text="SharePoint Calendar"
              description="Select a calendar from a SharePoint site"
              onClick={() => handleSelectType('sharepoint')}
              style={{ width: '100%', height: 'auto', padding: '16px', textAlign: 'left' }}
            />
          </Stack.Item>

          <Stack.Item>
            <PrimaryButton
              text="Exchange Calendar"
              description="Add a calendar from your mailbox or another mailbox"
              onClick={() => handleSelectType('exchange')}
              style={{ width: '100%', height: 'auto', padding: '16px', textAlign: 'left' }}
            />
          </Stack.Item>

          <Stack.Item>
            <PrimaryButton
              text="Online Calendar (ICS)"
              description="Add a calendar from an ICS URL or paste content"
              onClick={() => handleSelectType('ics')}
              style={{ width: '100%', height: 'auto', padding: '16px', textAlign: 'left' }}
            />
          </Stack.Item>
        </Stack>

        <DialogFooter>
          <DefaultButton onClick={onDismiss} text="Cancel" />
        </DialogFooter>
      </Dialog>
    );
  }
}
