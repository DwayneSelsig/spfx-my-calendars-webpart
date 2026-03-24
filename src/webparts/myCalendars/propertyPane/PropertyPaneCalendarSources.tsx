import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { IPropertyPaneField, PropertyPaneFieldType } from '@microsoft/sp-property-pane';
import { ICalendarSource } from '../models/ICalendarSettings';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { ColorPicker } from '@fluentui/react/lib/ColorPicker';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { Stack } from '@fluentui/react/lib/Stack';
import { Label } from '@fluentui/react/lib/Label';

export interface IPropertyPaneCalendarSourcesProps {
  label: string;
  sources: ICalendarSource[];
  onChanged: (sources: ICalendarSource[]) => void;
}

interface ICalendarSourcesControlState {
  sources: ICalendarSource[];
  editingId: string | undefined;
}

class CalendarSourcesControl extends React.Component<IPropertyPaneCalendarSourcesProps, ICalendarSourcesControlState> {
  constructor(props: IPropertyPaneCalendarSourcesProps) {
    super(props);
    this.state = {
      sources: props.sources || [],
      editingId: undefined
    };
  }

  private generateId(): string {
    return `source_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  private handleAddSource = (): void => {
    const newSource: ICalendarSource = {
      id: this.generateId(),
      sourceType: 'ics',
      name: 'New Calendar',
      color: '#0078d4',
      isEnabled: true
    };

    const sources = [...this.state.sources, newSource];
    this.setState({ sources, editingId: newSource.id });
    this.props.onChanged(sources);
  };

  private handleUpdateSource = (id: string, updates: Partial<ICalendarSource>): void => {
    const sources = this.state.sources.map(s => 
      s.id === id ? { ...s, ...updates } : s
    );
    this.setState({ sources });
    this.props.onChanged(sources);
  };

  private handleDeleteSource = (id: string): void => {
    const sources = this.state.sources.filter(s => s.id !== id);
    this.setState({ sources });
    this.props.onChanged(sources);
  };

  private toggleEdit = (id: string | undefined): void => {
    this.setState({ editingId: id });
  };

  public render(): React.ReactElement {
    const { label } = this.props;
    const { sources, editingId } = this.state;

    return (
      <div style={{ marginTop: 10 }}>
        <Label>{label}</Label>
        <Stack tokens={{ childrenGap: 10 }}>
          {sources.map(source => (
            <div key={source.id} style={{
              border: '1px solid #edebe9',
              borderRadius: 4,
              padding: 12,
              backgroundColor: editingId === source.id ? '#f3f2f1' : 'transparent'
            }}>
              {editingId === source.id ? (
                <Stack tokens={{ childrenGap: 8 }}>
                  <TextField
                    label="Name"
                    value={source.name}
                    onChange={(_, value) => this.handleUpdateSource(source.id, { name: value || '' })}
                  />
                  <div>
                    <Label>Color</Label>
                    <ColorPicker
                      color={source.color}
                      onChange={(_, color) => this.handleUpdateSource(source.id, { color: `#${color.hex}` })}
                      alphaType="none"
                    />
                  </div>
                  <Toggle
                    label="Enabled"
                    checked={source.isEnabled}
                    onChange={(_, checked) => this.handleUpdateSource(source.id, { isEnabled: !!checked })}
                  />
                  <Stack horizontal tokens={{ childrenGap: 8 }}>
                    <PrimaryButton text="Done" onClick={() => this.toggleEdit(undefined)} />
                    <DefaultButton text="Delete" onClick={() => this.handleDeleteSource(source.id)} />
                  </Stack>
                </Stack>
              ) : (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <div style={{
                    width: 16,
                    height: 16,
                    backgroundColor: source.color,
                    borderRadius: 2,
                    flexShrink: 0
                  }} />
                  <div style={{ flex: 1 }}>
                    <strong>{source.name}</strong>
                  </div>
                  <Toggle
                    checked={source.isEnabled}
                    onChange={(_, checked) => this.handleUpdateSource(source.id, { isEnabled: !!checked })}
                  />
                  <IconButton
                    iconProps={{ iconName: 'Edit' }}
                    onClick={() => this.toggleEdit(source.id)}
                  />
                </Stack>
              )}
            </div>
          ))}
          <PrimaryButton text="Add Calendar Source" onClick={this.handleAddSource} />
        </Stack>
      </div>
    );
  }
}

class PropertyPaneCalendarSourcesInternal implements IPropertyPaneField<IPropertyPaneCalendarSourcesProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneCalendarSourcesProps;
  private elem: HTMLElement | undefined;

  constructor(targetProperty: string, properties: IPropertyPaneCalendarSourcesProps) {
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
    if (!elem) {
      return;
    }
    
    this.elem = elem;
    ReactDOM.render(<CalendarSourcesControl {...this.properties} />, elem);
  }

  public onDispose(elem: HTMLElement): void {
    if (elem) {
      ReactDOM.unmountComponentAtNode(elem);
    }
  }
}

export function PropertyPaneCalendarSources(targetProperty: string, properties: IPropertyPaneCalendarSourcesProps): IPropertyPaneField<IPropertyPaneCalendarSourcesProps> {
  return new PropertyPaneCalendarSourcesInternal(targetProperty, properties);
}
