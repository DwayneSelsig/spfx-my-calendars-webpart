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
import {
  notifyAdminSettingsPropertyChanges,
  type AdminSettingsPropertyChangeNotifier,
  type PropertyPaneChangeCallback
} from '../services/AdminSettingsPropertyPersistence';
import { AdminSettingsPanel } from '../components/AdminSettingsPanel';

export interface IPropertyPaneAdminCalendarManagerProps {
  label: string;
  adminSettings: IAdminWebPartSettings;
  backupTargetProperty: string;
  adminLoadNotice?: string;
  context: WebPartContext;
  onSave: (settings: IAdminWebPartSettings, notifyPropertyChange: AdminSettingsPropertyChangeNotifier) => Promise<void> | void;
}

interface IAdminCalendarManagerControlProps extends IPropertyPaneAdminCalendarManagerProps {
  targetProperty: string;
  changeCallback?: PropertyPaneChangeCallback;
}

interface IAdminCalendarManagerControlState {
  isPanelOpen: boolean;
  isSaving: boolean;
  graphClient: MSGraphClientV3 | undefined;
}

class AdminCalendarManagerControl extends React.Component<IAdminCalendarManagerControlProps, IAdminCalendarManagerControlState> {
  constructor(props: IAdminCalendarManagerControlProps) {
    super(props);
    this.state = {
      isPanelOpen: false,
      isSaving: false,
      graphClient: undefined
    };
  }

  public componentDidMount(): void {
    this.props.context.msGraphClientFactory.getClient('3')
      .then(client => this.setState({ graphClient: client }))
      .catch(error => console.error('Failed to create graph client for admin property pane:', error));
  }

  private handleSave = async (settings: IAdminWebPartSettings): Promise<void> => {
    this.setState({ isSaving: true });
    try {
      if (!this.props.changeCallback) {
        throw new Error('The SPFx property pane change callback is unavailable.');
      }

      const changeCallback = this.props.changeCallback;
      await this.props.onSave(settings, serialized => {
        notifyAdminSettingsPropertyChanges(
          changeCallback,
          this.props.targetProperty,
          this.props.backupTargetProperty,
          serialized
        );
      });
      this.setState({ isPanelOpen: false, isSaving: false });
    } catch (error) {
      console.error('Failed to persist admin settings:', error);
      this.setState({ isSaving: false });
    }
  };

  public render(): React.ReactElement {
    const { adminSettings, adminLoadNotice, label, context } = this.props;
    const { isPanelOpen, graphClient, isSaving } = this.state;

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
            {adminSettings.assignedSources.length} admin default calendar(s), {adminSettings.icsCatalog.length} admin ICS catalog item(s), default view: {adminSettings.defaultView}
          </div>
          <DefaultButton
            text={isSaving ? 'Saving...' : 'Manage Admin Defaults'}
            onClick={() => this.setState({ isPanelOpen: true })}
            disabled={isSaving}
          />
        </Stack>

        <AdminSettingsPanel
          isOpen={isPanelOpen}
          onDismiss={() => this.setState({ isPanelOpen: false })}
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

export function PropertyPaneAdminCalendarManager(
  targetProperty: string,
  properties: IPropertyPaneAdminCalendarManagerProps
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty,
    properties: {
      key: targetProperty,
      onRender: (elem: HTMLElement, _context?: unknown, changeCallback?: PropertyPaneChangeCallback): void => {
        ReactDOM.render(
          <AdminCalendarManagerControl
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
