import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { IPropertyPaneField, PropertyPaneFieldType } from '@microsoft/sp-property-pane';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Label } from '@fluentui/react/lib/Label';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Stack } from '@fluentui/react/lib/Stack';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { IAdminWebPartSettings } from '../models/ICalendarSettings';
import { AdminSettingsPanel } from '../components/AdminSettingsPanel';

export interface IPropertyPaneAdminCalendarManagerProps {
  label: string;
  adminSettings: IAdminWebPartSettings;
  adminLoadNotice?: string;
  context: WebPartContext;
  onSave: (settings: IAdminWebPartSettings) => Promise<void> | void;
}

interface IAdminCalendarManagerControlState {
  isPanelOpen: boolean;
  isSaving: boolean;
  graphClient: MSGraphClientV3 | undefined;
}

class AdminCalendarManagerControl extends React.Component<IPropertyPaneAdminCalendarManagerProps, IAdminCalendarManagerControlState> {
  constructor(props: IPropertyPaneAdminCalendarManagerProps) {
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
      await this.props.onSave(settings);
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
        />
      </div>
    );
  }
}

class PropertyPaneAdminCalendarManagerInternal implements IPropertyPaneField<IPropertyPaneAdminCalendarManagerProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneAdminCalendarManagerProps;
  private elem: HTMLElement | undefined;

  constructor(targetProperty: string, properties: IPropertyPaneAdminCalendarManagerProps) {
    this.targetProperty = targetProperty;
    this.properties = properties;
  }

  public render(): void {
    if (!this.elem) {
      return;
    }

    this.onRender(this.elem);
  }

  private onRender(elem: HTMLElement): void {
    this.elem = elem;
    ReactDOM.render(<AdminCalendarManagerControl {...this.properties} />, elem);
  }

  public onDispose(elem: HTMLElement): void {
    if (elem) {
      ReactDOM.unmountComponentAtNode(elem);
    }
  }
}

export function PropertyPaneAdminCalendarManager(
  targetProperty: string,
  properties: IPropertyPaneAdminCalendarManagerProps
): IPropertyPaneField<IPropertyPaneAdminCalendarManagerProps> {
  return new PropertyPaneAdminCalendarManagerInternal(targetProperty, properties);
}
