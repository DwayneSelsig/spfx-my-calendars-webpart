import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {
  type IPropertyPaneCustomFieldProps,
  type IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Label } from '@fluentui/react/lib/Label';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Stack } from '@fluentui/react/lib/Stack';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { IAdminWebPartSettings } from '../models/ICalendarSettings';
import { AdminDefaultsPanel } from '../components/AdminDefaultsPanel';
import * as strings from 'MyCalendarsWebPartStrings';
import { formatLocalizedString } from '../utils/localization';
import { getCalendarLabels } from '../components/views/calendarLabels';

export interface IPropertyPaneAdminDefaultsManagerProps {
  label: string;
  getAdminSettings: () => IAdminWebPartSettings;
  getAdminLoadNotice: () => string | undefined;
  context: WebPartContext;
  onSave: (settings: IAdminWebPartSettings, commitProperty: (serialized: string | undefined) => void) => Promise<void>;
}

interface IAdminDefaultsManagerControlProps extends IPropertyPaneAdminDefaultsManagerProps {
  targetProperty: string;
  changeCallback?: Parameters<IPropertyPaneCustomFieldProps['onRender']>[2];
}

interface IAdminDefaultsManagerControlState {
  isPanelOpen: boolean;
  isSaving: boolean;
  graphClient: MSGraphClientV3 | undefined;
}

class AdminDefaultsManagerControl extends React.Component<IAdminDefaultsManagerControlProps, IAdminDefaultsManagerControlState> {
  constructor(props: IAdminDefaultsManagerControlProps) {
    super(props);
    this.state = {
      isPanelOpen: false,
      isSaving: false,
      graphClient: undefined
    };
  }

  private mounted = false;
  private savePending = false;

  public componentWillUnmount(): void { this.mounted = false; }

  public componentDidMount(): void {
    this.mounted = true;
    this.props.context.msGraphClientFactory.getClient('3')
      .then(client => { if (this.mounted) this.setState({ graphClient: client }); })
      .catch(error => console.error('Failed to create graph client for admin property pane:', error));
  }

  private handleSave = async (settings: IAdminWebPartSettings): Promise<void> => {
    if (this.savePending) throw new Error('An administrator save is already in progress.');
    this.savePending = true;
    this.setState({ isSaving: true });
    try {
      const changeCallback = this.props.changeCallback;
      if (!changeCallback) throw new Error('The SPFx property pane change callback is unavailable.');
      await this.props.onSave(settings, serialized => changeCallback(this.props.targetProperty, serialized, true));
    } finally {
      this.savePending = false;
      if (this.mounted) this.setState({ isSaving: false });
    }
  };

  public render(): React.ReactElement {
    const { label, context } = this.props;
    const adminSettings = this.props.getAdminSettings();
    const adminLoadNotice = this.props.getAdminLoadNotice();
    const { isPanelOpen, graphClient, isSaving } = this.state;
    const calendarLabels = getCalendarLabels();
    const defaultViewLabel = adminSettings.defaultView === 'day'
      ? calendarLabels.day
      : adminSettings.defaultView === 'week' ? calendarLabels.week : calendarLabels.month;

    return (
      <div style={{ marginTop: 12 }}>
        <Stack tokens={{ childrenGap: 8 }}>
          <Label>{label}</Label>
          {adminLoadNotice && (
            <MessageBar messageBarType={MessageBarType.warning}>
              {adminLoadNotice}
            </MessageBar>
          )}
          <div style={{ fontSize: 12, color: '#605e5c' }}>
            {formatLocalizedString(strings.AdminPropertyPaneSummaryLabel, adminSettings.assignedSources.length + (adminSettings.exchangeMailboxAssignments || []).length, adminSettings.icsCatalog.length, defaultViewLabel)}
          </div>
          <DefaultButton
            text={isSaving ? strings.SavingLabel : strings.ManageAdminDefaultsLabel}
            onClick={() => this.setState({ isPanelOpen: true })}
            disabled={isSaving}
          />
        </Stack>

        <AdminDefaultsPanel
          isOpen={isPanelOpen}
          onDismiss={() => { if (!this.savePending) this.setState({ isPanelOpen: false }); }}
          settings={adminSettings}
          onSave={this.handleSave}
          httpClient={context.httpClient}
          graphClient={graphClient}
          loadNotice={adminLoadNotice}
          locale={context.pageContext.cultureInfo.currentCultureName}
        />
      </div>
    );
  }
}

export function PropertyPaneAdminDefaultsManager(
  targetProperty: string,
  properties: IPropertyPaneAdminDefaultsManagerProps
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty,
    properties: {
      key: targetProperty,
      onRender: (elem, _context, changeCallback): void => {
        ReactDOM.render(
          <AdminDefaultsManagerControl
            {...properties}
            targetProperty={targetProperty}
            changeCallback={changeCallback}
          />,
          elem
        );
      },
      onDispose: (elem: HTMLElement): void => {
        ReactDOM.unmountComponentAtNode(elem);
      }
    }
  };
}
