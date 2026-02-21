import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { ColorPicker } from '@fluentui/react/lib/ColorPicker';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { Stack } from '@fluentui/react/lib/Stack';
import { Label } from '@fluentui/react/lib/Label';
import { Slider } from '@fluentui/react/lib/Slider';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Icon } from '@fluentui/react/lib/Icon';
import { HttpClient } from '@microsoft/sp-http';
import { ICalendarSource, ICalendarSettings, CalendarSourceType } from '../models/ICalendarSettings';
import * as strings from 'MyCalendarsWebPartStrings';

// MSGraphClientV3 type - using any since @microsoft/sp-client-preview is not available
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type MSGraphClientV3 = any;
import { ExchangeCalendarService, IExchangeCalendar } from '../services/ExchangeCalendarService';
import { SharePointCalendarService, ISharePointSite, ISharePointList } from '../services/SharePointCalendarService';

export interface ISettingsPanelProps {
  isOpen: boolean;
  onDismiss: () => void;
  settings: ICalendarSettings;
  onSave: (settings: ICalendarSettings) => void;
  onReset?: () => void;
  httpClient?: HttpClient;
  graphClient?: MSGraphClientV3;
}

interface ISettingsPanelState {
  settings: ICalendarSettings;
  editingSourceId: string | undefined;
  showAddDialog: boolean;
  // Exchange calendars (auto-loaded)
  userExchangeCalendars: IExchangeCalendar[];
  userExchangeCalendarsLoading: boolean;
  // Add calendar flow state
  addingCalendarType: CalendarSourceType | undefined;
  addingCalendarStep: 'initial' | 'sharepoint-site' | 'sharepoint-list' | 'sharepoint-fields' | 'exchange-calendar' | 'exchange-mailbox' | 'ics';
  // SharePoint flow
  spSites: ISharePointSite[];
  spSitesLoading: boolean;
  spSiteFilter: string;
  spCurrentPage: number;
  spSelectedSite: ISharePointSite | undefined;
  spLists: ISharePointList[];
  spListsLoading: boolean;
  spSelectedList: ISharePointList | undefined;
  // Exchange flow
  exchangeCalendars: IExchangeCalendar[];
  exchangeCalendarsLoading: boolean;
  exchangeMailbox: string;
  exchangeMailboxResolved: boolean;
  exchangeSelectedCalendarId: string | undefined;
  // SharePoint field mapping
  spAvailableFields: IDropdownOption[];
  spFieldMapping: {
    titleField?: string;
    startDateField?: string;
    endDateField?: string;
    descriptionField?: string;
    locationField?: string;
    allDayField?: string;
  };
  // ICS flow
  icsUrl: string;
  // Color for new calendar
  newCalendarColor: string;
  newCalendarName: string;
}

interface IGraphColumn {
  name?: string;
  displayName?: string;
  columnGroup?: string;
}

export class SettingsPanel extends React.Component<ISettingsPanelProps, ISettingsPanelState> {
  private exchangeService: ExchangeCalendarService | null = null;
  private sharePointService: SharePointCalendarService | null = null;
  private readonly SITES_PER_PAGE = 20;

  constructor(props: ISettingsPanelProps) {
    super(props);
    
    if (props.httpClient) {
      this.exchangeService = new ExchangeCalendarService(props.httpClient, props.graphClient);
      this.sharePointService = new SharePointCalendarService(props.httpClient, props.graphClient);
      
      // If graphClient is provided, set it on the services
      if (props.graphClient) {
        this.exchangeService.setGraphClient(props.graphClient);
        this.sharePointService.setGraphClient(props.graphClient);
      }
    }

    this.state = {
      settings: JSON.parse(JSON.stringify(props.settings)),
      editingSourceId: undefined,
      showAddDialog: false,
      // User Exchange calendars
      userExchangeCalendars: [],
      userExchangeCalendarsLoading: false,
      addingCalendarType: undefined,
      addingCalendarStep: 'initial',
      // SharePoint
      spSites: [],
      spSitesLoading: false,
      spSiteFilter: '',
      spCurrentPage: 0,
      spSelectedSite: undefined,
      spLists: [],
      spListsLoading: false,
      spSelectedList: undefined,
      // Exchange
      exchangeCalendars: [],
      exchangeCalendarsLoading: false,
      exchangeMailbox: '',
      exchangeMailboxResolved: false,
      exchangeSelectedCalendarId: undefined,
      // SharePoint field mapping
      spAvailableFields: [],
      spFieldMapping: {},
      // ICS
      icsUrl: '',
      // New calendar
      newCalendarColor: props.settings.organizationPrimaryColor || '#0078d4',
      newCalendarName: ''
    };
  }

  public componentDidMount(): void {
    // Initialize GraphClient async from the promise
    if (this.props.graphClient) {
      Promise.resolve(this.props.graphClient).then((client) => {
        if (client) {
          if (this.sharePointService) {
            this.sharePointService.setGraphClient(client);
          }
          if (this.exchangeService) {
            this.exchangeService.setGraphClient(client);
          }
          // Load user's Exchange calendars automatically
          this.loadUserExchangeCalendars().catch(err => console.error('Failed to load Exchange calendars:', err));
        }
      }).catch((err) => console.error('Failed to get GraphClient:', err));
    }
  }

  private loadUserExchangeCalendars = async (): Promise<void> => {
    if (!this.exchangeService) {
      return;
    }

    this.setState({ userExchangeCalendarsLoading: true });
    try {
      const calendars = await this.exchangeService.getCalendars();
      this.setState({ 
        userExchangeCalendars: calendars,
        userExchangeCalendarsLoading: false
      });
    } catch (error) {
      console.error('Error loading user Exchange calendars:', error);
      this.setState({ 
        userExchangeCalendars: [],
        userExchangeCalendarsLoading: false
      });
    }
  };

  public componentDidUpdate(prevProps: ISettingsPanelProps): void {
    if (prevProps.isOpen !== this.props.isOpen && this.props.isOpen) {
      this.setState({
        settings: JSON.parse(JSON.stringify(this.props.settings)),
        editingSourceId: undefined,
        showAddDialog: false,
        addingCalendarType: undefined,
        addingCalendarStep: 'initial',
        spCurrentPage: 0
      });
      // Reload Exchange calendars when panel opens
      this.loadUserExchangeCalendars().catch(err => console.error('Failed to reload Exchange calendars:', err));
    }
  }

  private generateId(): string {
    return `source_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  private handleOpenAddDialog = (): void => {
    this.setState({ showAddDialog: true });
  };

  private handleCloseAddDialog = (): void => {
    this.setState({
      showAddDialog: false,
      addingCalendarType: undefined,
      addingCalendarStep: 'initial',
      spSites: [],
      spSiteFilter: '',
      spCurrentPage: 0,
      spSelectedSite: undefined,
      spLists: [],
      spSelectedList: undefined,
      exchangeCalendars: [],
      exchangeMailbox: '',
      exchangeMailboxResolved: false,
      exchangeSelectedCalendarId: undefined,
      icsUrl: '',
      newCalendarColor: this.props.settings.organizationPrimaryColor || '#0078d4',
      newCalendarName: ''
    });
  };

  private handleSelectAddType = async (type: CalendarSourceType): Promise<void> => {
    if (type === 'sharepoint') {
      this.setState({ addingCalendarType: type, addingCalendarStep: 'sharepoint-site', spSitesLoading: true });
      const sites = await this.sharePointService?.getAccessibleSites() || [];
      this.setState({ spSites: sites, spSitesLoading: false });
    } else if (type === 'exchange') {
      this.setState({ addingCalendarType: type, addingCalendarStep: 'exchange-mailbox' });
    } else if (type === 'ics') {
      this.setState({ addingCalendarType: type, addingCalendarStep: 'ics' });
    }
  };

  private handleBackToTypeSelection = (): void => {
    this.setState({
      addingCalendarStep: 'initial',
      spSites: [],
      spSitesLoading: false,
      spSiteFilter: '',
      spCurrentPage: 0,
      spSelectedSite: undefined,
      spLists: [],
      spListsLoading: false,
      spSelectedList: undefined,
      exchangeCalendars: [],
      exchangeCalendarsLoading: false,
      exchangeMailbox: '',
      exchangeMailboxResolved: false,
      exchangeSelectedCalendarId: undefined,
      icsUrl: '',
      newCalendarColor: this.props.settings.organizationPrimaryColor || '#0078d4',
      newCalendarName: ''
    });
  };

  private handleBackOneStep = (): void => {
    const { addingCalendarStep, addingCalendarType } = this.state;

    if (addingCalendarType === 'sharepoint') {
      if (addingCalendarStep === 'sharepoint-fields') {
        this.setState({ spSelectedList: undefined, addingCalendarStep: 'sharepoint-list' });
      } else if (addingCalendarStep === 'sharepoint-list') {
        this.setState({ spSelectedSite: undefined, spLists: [], addingCalendarStep: 'sharepoint-site' });
      } else if (addingCalendarStep === 'sharepoint-site') {
        this.handleBackToTypeSelection();
      }
    } else if (addingCalendarType === 'exchange') {
      if (this.state.exchangeSelectedCalendarId) {
        this.setState({ exchangeSelectedCalendarId: undefined });
      } else {
        this.handleBackToTypeSelection();
      }
    } else if (addingCalendarType === 'ics') {
      this.handleBackToTypeSelection();
    }
  };

  private renderNavigationHeader = (): React.ReactElement | null => {
    const { addingCalendarStep, addingCalendarType } = this.state;

    // Don't show navigation on initial step
    if (addingCalendarStep === 'initial') {
      return null;
    }

    const showBackButton = addingCalendarType && (
      (addingCalendarType === 'sharepoint' && (this.state.spSelectedSite || this.state.spSelectedList)) ||
      (addingCalendarType === 'exchange' && this.state.exchangeSelectedCalendarId)
    );

    return (
      <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 16 }}>
        {showBackButton && (
          <IconButton
            iconProps={{ iconName: 'Back' }}
            title="Back"
            ariaLabel="Back"
            onClick={this.handleBackOneStep}
            styles={{ root: { height: 32 } }}
          />
        )}
        <IconButton
          iconProps={{ iconName: 'Home' }}
          title="Home"
          ariaLabel="Home"
          onClick={this.handleBackToTypeSelection}
          styles={{ root: { height: 32 } }}
        />
      </Stack>
    );
  };

  // SharePoint flow
  private handleSharePointFilterChange = (value?: string): void => {
    this.setState({ spSiteFilter: value || '', spCurrentPage: 0 });
  };

  private handleSharePointSearch = async (): Promise<void> => {
    this.setState({ spSitesLoading: true, spCurrentPage: 0 });
    const sites = await this.sharePointService?.searchSites(this.state.spSiteFilter) || [];
    this.setState({ spSites: sites, spSitesLoading: false });
  };

  private handleSelectSharePointSite = async (site: ISharePointSite): Promise<void> => {
    this.setState({ spSelectedSite: site, spListsLoading: true, spLists: [] });
    const lists = await this.sharePointService?.getCalendarLists(site.id) || [];
    this.setState({ spLists: lists, spListsLoading: false, addingCalendarStep: 'sharepoint-list' });
  };

  private handleSelectSharePointList = (list: ISharePointList): void => {
    // Use organization primary color for first SharePoint calendar, or let user choose
    this.setState({
      spSelectedList: list,
      newCalendarName: list.name,
      newCalendarColor: this.props.settings.organizationPrimaryColor || '#0078d4'
    }, () => {
      // Fetch available fields from first item
      void this.fetchSharePointListFields(list);
    });
  };

  private fetchSharePointListFields = async (list: ISharePointList): Promise<void> => {
    const { spSelectedSite } = this.state;
    if (!spSelectedSite || !this.sharePointService) {
      return;
    }

    try {
      // Fetch one item to get available fields
      const graphClient = this.props.graphClient;
      if (!graphClient) return;

      const client = await Promise.resolve(graphClient);
      if (!client) return;

      const columnsData = await client
        .api(`/sites/${spSelectedSite.id}/lists/${list.id}/columns`)
        .query({ $select: 'name,displayName,columnGroup' })
        .get();

      const rawOptions: IDropdownOption[] = (columnsData.value || [])
        .filter((column: IGraphColumn) => column.name && !column.name.startsWith('_') && column.columnGroup !== '_Hidden')
        .map((column: IGraphColumn) => ({
          key: column.name as string,
          text: column.displayName || (column.name as string)
        }));

      const dedupedOptions = new Map<string, IDropdownOption>();
      rawOptions.forEach(option => {
        const key = String(option.text || option.key);
        const existing = dedupedOptions.get(key);
        const candidateKey = String(option.key);
        if (!existing || candidateKey.length < String(existing.key).length) {
          dedupedOptions.set(key, option);
        }
      });

      let fieldOptions: IDropdownOption[] = Array.from(dedupedOptions.values());

      if (fieldOptions.length === 0) {
        const itemsData = await client
          .api(`/sites/${spSelectedSite.id}/lists/${list.id}/items`)
          .expand('fields')
          .get();

        if (itemsData.value && itemsData.value.length > 0) {
          const fields = Object.keys(itemsData.value[0].fields || {});
          fieldOptions = fields
            .filter(fieldName => !fieldName.startsWith('_'))
            .map(fieldName => ({ key: fieldName, text: fieldName }));
        }
      }

      if (fieldOptions.length > 0) {
        const candidateLists = {
          title: this.getFieldCandidates('title'),
          start: this.getFieldCandidates('start'),
          end: this.getFieldCandidates('end')
        };

        const guessedMapping = { ...this.state.spFieldMapping };
        if (!guessedMapping.titleField) {
          guessedMapping.titleField = this.findBestMatchingFieldKey(fieldOptions, candidateLists.title);
        }
        if (!guessedMapping.startDateField) {
          guessedMapping.startDateField = this.findBestMatchingFieldKey(fieldOptions, candidateLists.start);
        }
        if (!guessedMapping.endDateField) {
          guessedMapping.endDateField = this.findBestMatchingFieldKey(fieldOptions, candidateLists.end);
        }

        this.setState({ 
          spAvailableFields: fieldOptions,
          addingCalendarStep: 'sharepoint-fields',
          spFieldMapping: guessedMapping
        });
      }
    } catch (error) {
      console.error('Failed to fetch SharePoint list fields:', error);
      // Fallback to next step anyway
      this.setState({ addingCalendarStep: 'sharepoint-fields' });
    }
  };

  private parseFieldCandidates(value: string | undefined): string[] {
    if (!value) {
      return [];
    }
    return value
      .split(';')
      .map(candidate => candidate.trim())
      .filter(Boolean);
  }

  private getFieldCandidates(field: 'title' | 'start' | 'end'): string[] {
    const defaults: Record<'title' | 'start' | 'end', string[]> = {
      title: ['Title', 'Subject', 'Event Title'],
      start: ['Start Time', 'Start', 'Start Date', 'StartDate', 'StartDateTime', 'EventDate', 'Starttijd', 'Begindatum', 'Begin', 'Startdatum'],
      end: ['End Time', 'End', 'End Date', 'EndDate', 'EndDateTime', 'Eindtijd', 'Einddatum', 'Einde']
    };

    const localized = this.parseFieldCandidates(
      field === 'title' ? strings.FieldTitleCandidates
        : field === 'start' ? strings.FieldStartCandidates
          : strings.FieldEndCandidates
    );

    const merged = [...localized, ...defaults[field]];
    const seen = new Set<string>();
    return merged.filter(candidate => {
      const normalized = this.normalizeFieldCandidate(candidate);
      if (!normalized || seen.has(normalized)) {
        return false;
      }
      seen.add(normalized);
      return true;
    });
  }

  private normalizeFieldCandidate(value: string): string {
    return value.toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  private findBestMatchingFieldKey(options: IDropdownOption[], candidates: string[]): string | undefined {
    const normalizedCandidates = candidates.map(candidate => this.normalizeFieldCandidate(candidate));
    const normalizedOptions = options.map(option => ({
      option,
      text: this.normalizeFieldCandidate(String(option.text || '')),
      key: this.normalizeFieldCandidate(String(option.key || ''))
    }));

    for (const candidate of normalizedCandidates) {
      if (!candidate) {
        continue;
      }
      const exactMatch = normalizedOptions.find(entry => entry.text === candidate || entry.key === candidate);
      if (exactMatch) {
        return exactMatch.option.key as string;
      }
    }

    for (const candidate of normalizedCandidates) {
      if (!candidate) {
        continue;
      }
      const partialMatch = normalizedOptions.find(entry => entry.text.includes(candidate) || entry.key.includes(candidate));
      if (partialMatch) {
        return partialMatch.option.key as string;
      }
    }

    return undefined;
  }

  private handleConfirmSharePointCalendar = (): void => {
    if (!this.state.spSelectedSite || !this.state.spSelectedList) {
      return;
    }

    const newSource: ICalendarSource = {
      id: this.generateId(),
      sourceType: 'sharepoint',
      name: this.state.newCalendarName,
      color: this.state.newCalendarColor,
      isEnabled: true,
      sharePointSiteId: this.state.spSelectedSite.id,
      sharePointListId: this.state.spSelectedList.id,
      sharePointFieldMapping: this.state.spFieldMapping
    };

    const settings = {
      ...this.state.settings,
      sources: [...this.state.settings.sources, newSource]
    };

    this.setState({ settings }, () => this.handleCloseAddDialog());
  };

  // Exchange flow
  private handleExchangeMailboxChange = (value?: string): void => {
    this.setState({ exchangeMailbox: value || '' });
  };

  private handleExchangeLookupMailbox = async (): Promise<void> => {
    if (!this.state.exchangeMailbox.trim()) {
      return;
    }

    this.setState({ exchangeCalendarsLoading: true });
    const resolved = await this.exchangeService?.resolveMailbox(this.state.exchangeMailbox);
    
    if (resolved) {
      const calendars = await this.exchangeService?.getCalendars(this.state.exchangeMailbox) || [];
      this.setState({
        exchangeCalendars: calendars,
        exchangeCalendarsLoading: false,
        exchangeMailboxResolved: true,
        addingCalendarStep: 'exchange-calendar'
      });
    } else {
      this.setState({
        exchangeCalendarsLoading: false,
        exchangeMailboxResolved: false
      });
      alert('Mailbox not found or not accessible');
    }
  };

  private handleSelectExchangeCalendar = (calendar: IExchangeCalendar): void => {
    this.setState({
      exchangeSelectedCalendarId: calendar.id,
      newCalendarName: calendar.name,
      newCalendarColor: calendar.hexColor // Use the hex color from the calendar
    });
  };

  private handleConfirmExchangeCalendar = (): void => {
    const newSource: ICalendarSource = {
      id: this.generateId(),
      sourceType: 'exchange',
      name: this.state.newCalendarName,
      color: this.state.newCalendarColor,
      isEnabled: true,
      exchangeMailbox: this.state.exchangeMailbox || undefined,
      exchangeCalendarId: this.state.exchangeSelectedCalendarId || 'calendar'
    };

    const settings = {
      ...this.state.settings,
      sources: [...this.state.settings.sources, newSource]
    };

    this.setState({ settings }, () => this.handleCloseAddDialog());
  };

  // ICS flow
  private handleGenerateOutlookLink = (): void => {
    const { icsUrl, newCalendarName } = this.state;

    if (!icsUrl.trim() || !newCalendarName.trim()) {
      return;
    }

    // Encode the URL and name for the Outlook calendar add URL
    const encodedIcsUrl = encodeURIComponent(icsUrl.trim());
    const encodedName = encodeURIComponent(newCalendarName.trim());

    // Generate and open the Outlook calendar add link
    const outlookLink = `https://outlook.office.com/calendar/addcalendar?url=${encodedIcsUrl}&name=${encodedName}`;
    window.open(outlookLink, '_blank', 'noopener,noreferrer');
  };

  private handleOpenOutlookCalendar = (): void => {
    this.handleGenerateOutlookLink();
    this.handleCloseAddDialog();
  };

  private handleUpdateSource = (id: string, updates: Partial<ICalendarSource>): void => {
    const settings = {
      ...this.state.settings,
      sources: this.state.settings.sources.map(s => 
        s.id === id ? { ...s, ...updates } : s
      )
    };
    this.setState({ settings });
  };

  private handleToggleExchangeCalendar = (calendarId: string, isEnabled: boolean): void => {
    const settings = {
      ...this.state.settings,
      exchangeCalendarStates: {
        ...(this.state.settings.exchangeCalendarStates || {}),
        [calendarId]: isEnabled
      }
    };
    this.setState({ settings });
  };

  private isExchangeCalendarEnabled = (calendarId: string): boolean => {
    const states = this.state.settings.exchangeCalendarStates || {};
    // If not set, default to enabled
    return states[calendarId] !== false;
  };

  private handleDeleteSource = (id: string): void => {
    const settings = {
      ...this.state.settings,
      sources: this.state.settings.sources.filter(s => s.id !== id)
    };
    this.setState({ settings });
  };

  private toggleEdit = (id: string | undefined): void => {
    this.setState({ editingSourceId: id });
  };

  private handleSave = (): void => {
    this.props.onSave(this.state.settings);
    this.props.onDismiss();
  };

  private handleReset = (): void => {
    if (confirm('Are you sure you want to reset all settings to the defaults? This action cannot be undone.')) {
      if (this.props.onReset) {
        this.props.onReset();
      }
      this.props.onDismiss();
    }
  };

  private onRenderFooterContent = (): React.ReactElement => {
    return (
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton onClick={this.handleSave} text="Save" />
        <DefaultButton onClick={this.props.onDismiss} text="Cancel" />
        <DefaultButton 
          onClick={this.handleReset} 
          text="Reset to Defaults" 
          title="Reset your calendar settings to administrator defaults"
        />
      </Stack>
    );
  };

  private renderSharePointFlow = (): React.ReactElement => {
    const { spSites, spSitesLoading, spSiteFilter, spSelectedSite, spLists, spListsLoading, spSelectedList, spAvailableFields, spFieldMapping, addingCalendarStep } = this.state;

    // Field mapping step
    if (spSelectedList && addingCalendarStep === 'sharepoint-fields') {
      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>Map SharePoint fields to calendar fields</Label>
          <Dropdown
            label="Title/Subject field"
            options={spAvailableFields}
            selectedKey={spFieldMapping.titleField || ''}
            onChange={(_, option) => this.setState({ 
              spFieldMapping: { ...spFieldMapping, titleField: option?.key as string } 
            })}
          />
          <Dropdown
            label="Start Date field"
            options={spAvailableFields}
            selectedKey={spFieldMapping.startDateField || ''}
            onChange={(_, option) => this.setState({ 
              spFieldMapping: { ...spFieldMapping, startDateField: option?.key as string } 
            })}
          />
          <Dropdown
            label="End Date field"
            options={spAvailableFields}
            selectedKey={spFieldMapping.endDateField || ''}
            onChange={(_, option) => this.setState({ 
              spFieldMapping: { ...spFieldMapping, endDateField: option?.key as string } 
            })}
          />
          <Dropdown
            label="Location field (optional)"
            options={[{ key: '', text: '(none)' }, ...spAvailableFields]}
            selectedKey={spFieldMapping.locationField || ''}
            onChange={(_, option) => this.setState({ 
              spFieldMapping: { ...spFieldMapping, locationField: option?.key as string } 
            })}
          />
          <Dropdown
            label="Description field (optional)"
            options={[{ key: '', text: '(none)' }, ...spAvailableFields]}
            selectedKey={spFieldMapping.descriptionField || ''}
            onChange={(_, option) => this.setState({ 
              spFieldMapping: { ...spFieldMapping, descriptionField: option?.key as string } 
            })}
          />
          <TextField
            label="Calendar Name"
            value={this.state.newCalendarName}
            onChange={(_, value) => this.setState({ newCalendarName: value || '' })}
          />
          <div>
            <Label>Color</Label>
            <ColorPicker
              color={this.state.newCalendarColor}
              onChange={(_, color) => this.setState({ newCalendarColor: `#${color.hex}` })}
              alphaType="none"
            />
          </div>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Add Calendar" onClick={this.handleConfirmSharePointCalendar} />
          </Stack>
        </Stack>
      );
    }

    if (spSelectedList) {
      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>Selected List: {spSelectedList.name}</Label>
          <Label style={{ color: '#605e5c' }}>Configuring field mapping...</Label>
        </Stack>
      );
    }

    if (spSelectedSite) {
      if (spListsLoading) {
        return <Spinner size={SpinnerSize.medium} label="Loading lists..." />;
      }

      if (spLists.length === 0) {
        return (
          <Stack tokens={{ childrenGap: 12 }}>
            <Label>No calendar lists found in this site</Label>
          </Stack>
        );
      }

      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>Select a calendar list from {spSelectedSite.name}:</Label>
          <Stack tokens={{ childrenGap: 8 }}>
            {spLists.map(list => (
              <DefaultButton
                key={list.id}
                text={list.name}
                onClick={() => this.handleSelectSharePointList(list)}
                style={{ textAlign: 'left', height: 'auto', padding: '8px' }}
              />
            ))}
          </Stack>
        </Stack>
      );
    }

    if (spSitesLoading) {
      return <Spinner size={SpinnerSize.medium} label="Loading sites..." />;
    }

    // Pagination logic
    const filteredSites = spSites;
    const totalPages = Math.ceil(filteredSites.length / this.SITES_PER_PAGE);
    const startIndex = this.state.spCurrentPage * this.SITES_PER_PAGE;
    const endIndex = startIndex + this.SITES_PER_PAGE;
    const sitesOnPage = filteredSites.slice(startIndex, endIndex);

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>Search and select a SharePoint site:</Label>
        <TextField
          placeholder="Type to filter sites..."
          value={spSiteFilter}
          onChange={(_, value) => this.handleSharePointFilterChange(value)}
        />
        <PrimaryButton text="Search" onClick={this.handleSharePointSearch} />
        
        {spSites.length === 0 && !spSiteFilter && (
          <Label style={{ color: '#605e5c', fontStyle: 'italic' }}>
            No sites found. Try searching for a specific site name.
          </Label>
        )}

        {spSites.length > 0 && (
          <>
            <Stack tokens={{ childrenGap: 8 }}>
              {sitesOnPage.map(site => (
                <Stack
                  key={site.id}
                  horizontal
                  verticalAlign="center"
                  tokens={{ childrenGap: 8 }}
                  onClick={() => this.handleSelectSharePointSite(site)}
                  style={{
                    border: '1px solid #edebe9',
                    borderRadius: 4,
                    padding: '8px 12px',
                    backgroundColor: '#f3f2f1',
                    cursor: 'pointer',
                    transition: 'background-color 0.2s'
                  }}
                  onMouseEnter={(e) => e.currentTarget.style.backgroundColor = '#e1dfdd'}
                  onMouseLeave={(e) => e.currentTarget.style.backgroundColor = '#f3f2f1'}
                >
                  <div style={{ flex: 1 }}>
                    <strong>{site.name}</strong>
                    <div style={{ fontSize: 12, color: '#605e5c' }}>{site.url}</div>
                  </div>
                  <IconButton
                    iconProps={{ iconName: 'ChevronRight' }}
                    title="Select this site"
                    styles={{ root: { pointerEvents: 'none' } }}
                  />
                </Stack>
              ))}
            </Stack>

            {/* Pagination Controls */}
            {totalPages > 1 && (
              <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="center">
                <DefaultButton
                  text="Previous"
                  disabled={this.state.spCurrentPage === 0}
                  onClick={() => this.setState({ spCurrentPage: this.state.spCurrentPage - 1 })}
                />
                <Label style={{ margin: 0 }}>
                  Page {this.state.spCurrentPage + 1} of {totalPages}
                </Label>
                <DefaultButton
                  text="Next"
                  disabled={this.state.spCurrentPage >= totalPages - 1}
                  onClick={() => this.setState({ spCurrentPage: this.state.spCurrentPage + 1 })}
                />
              </Stack>
            )}

            <Label style={{ fontSize: 12, color: '#605e5c' }}>
              Showing {sitesOnPage.length} of {filteredSites.length} sites
            </Label>
          </>
        )}
      </Stack>
    );
  };

  private renderExchangeFlow = (): React.ReactElement => {
    const { exchangeMailbox, exchangeMailboxResolved, exchangeCalendars, exchangeCalendarsLoading, exchangeSelectedCalendarId } = this.state;

    if (exchangeSelectedCalendarId) {
      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <TextField
            label="Calendar Name"
            value={this.state.newCalendarName}
            onChange={(_, value) => this.setState({ newCalendarName: value || '' })}
          />
          <div>
            <Label>Color</Label>
            <ColorPicker
              color={this.state.newCalendarColor}
              onChange={(_, color) => this.setState({ newCalendarColor: `#${color.hex}` })}
              alphaType="none"
            />
          </div>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Add Calendar" onClick={this.handleConfirmExchangeCalendar} />
          </Stack>
        </Stack>
      );
    }

    if (exchangeCalendarsLoading) {
      return <Spinner size={SpinnerSize.medium} label="Loading calendars..." />;
    }

    if (exchangeMailboxResolved && exchangeCalendars.length > 0) {
      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>Select a calendar from {exchangeMailbox || 'your mailbox'}:</Label>
          <Stack tokens={{ childrenGap: 8 }}>
            {exchangeCalendars.map(cal => (
              <Stack
                key={cal.id}
                horizontal
                verticalAlign="center"
                tokens={{ childrenGap: 8 }}
                style={{
                  border: '1px solid #edebe9',
                  borderRadius: 4,
                  padding: '8px 12px',
                  backgroundColor: '#f3f2f1',
                  cursor: 'pointer'
                }}
                onClick={() => this.handleSelectExchangeCalendar(cal)}
              >
                <div style={{
                  width: 16,
                  height: 16,
                  backgroundColor: cal.hexColor,
                  borderRadius: 2,
                  flexShrink: 0
                }} />
                <div style={{ flex: 1 }}>
                  <strong>{cal.name}</strong>
                  {cal.isDefaultCalendar && (
                    <span style={{ fontSize: 11, color: '#605e5c', marginLeft: 8 }}>(Default)</span>
                  )}
                </div>
              </Stack>
            ))}
          </Stack>
          
          {/* Option to search another mailbox */}
          <div style={{ borderTop: '1px solid #edebe9', paddingTop: 12, marginTop: 8 }}>
            <Label>Or enter another mailbox email:</Label>
            <TextField
              placeholder="user@example.com"
              value={exchangeMailbox}
              onChange={(_, value) => this.handleExchangeMailboxChange(value)}
            />
            <div style={{ marginTop: 8 }}>
              <PrimaryButton text="Load Other Mailbox" onClick={this.handleExchangeLookupMailbox} />
            </div>
          </div>
        </Stack>
      );
    }

    // Default: Show instructions to open Outlook
    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <Label>To add an Outlook calendar from a shared mailbox or another user:</Label>
        <div style={{
          padding: '16px',
          backgroundColor: '#f3f2f1',
          borderRadius: 4,
          border: '1px solid #edebe9'
        }}>
          <ol style={{ margin: 0, paddingLeft: 20 }}>
            <li style={{ marginBottom: 8 }}>
              Open Outlook (web) by clicking the button below
            </li>
            <li style={{ marginBottom: 8 }}>
              In Outlook, click on the left sidebar button:
              <div style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: 4,
                padding: '4px 8px',
                backgroundColor: '#fff',
                border: '1px solid #ccc',
                borderRadius: 2,
                marginLeft: 8,
                fontSize: 12
              }}>
                <span style={{ fontSize: 14 }}>👥</span>
                <strong>Add from directory</strong>
              </div>
            </li>
            <li>
              Select the calendar you want to add and it will automatically appear in this webpart
            </li>
          </ol>
        </div>
        <PrimaryButton
          text="Open Outlook (web)"
          iconProps={{ iconName: 'OpenInNewWindow' }}
          onClick={() => {
            window.open('https://outlook.cloud.microsoft/calendar/addcalendar', '_blank', 'noopener,noreferrer');
            this.handleCloseAddDialog();
          }}
        />
      </Stack>
    );
  };

  private renderIcsFlow = (): React.ReactElement => {
    const { icsUrl } = this.state;
    const hasValidInput = icsUrl.trim() && this.state.newCalendarName.trim();

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <TextField
          label="Calendar Name"
          value={this.state.newCalendarName}
          onChange={(_, value) => this.setState({ newCalendarName: value || '' })}
          placeholder="e.g. Office Holidays"
        />
        <TextField
          label="ICS URL"
          value={icsUrl}
          onChange={(_, value) => this.setState({ icsUrl: value || '' })}
          placeholder="https://www.officeholidays.com/ics-all/netherlands"
        />
        
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton 
            text="Open in Outlook" 
            onClick={this.handleOpenOutlookCalendar}
            disabled={!hasValidInput}
            iconProps={{ iconName: 'OpenInNewWindow' }}
          />
          <DefaultButton text="Cancel" onClick={this.handleCloseAddDialog} />
        </Stack>
      </Stack>
    );
  };

  private renderAddCalendarFlow = (): React.ReactElement => {
    const { addingCalendarStep, addingCalendarType } = this.state;

    // Show type selection first
    if (addingCalendarStep === 'initial') {
      return (
        <Stack tokens={{ childrenGap: 16 }}>
          <Label>Select the type of calendar to add:</Label>
          <Stack tokens={{ childrenGap: 12 }}>
            <PrimaryButton
              text="SharePoint Calendar"
              secondaryText="Select a calendar from a SharePoint site"
              iconProps={{ iconName: 'SharepointLogo' }}
              onClick={() => this.handleSelectAddType('sharepoint')}
              style={{ textAlign: 'left', height: 'auto', padding: '12px' }}
            />
            <PrimaryButton
              text="Outlook Calendar"
              secondaryText="Manage Outlook calendars (opens in new window)"
              iconProps={{ iconName: 'OutlookLogo' }}
              onClick={() => this.handleSelectAddType('exchange')}
              style={{ textAlign: 'left', height: 'auto', padding: '12px' }}
            />
            <PrimaryButton
              text="Internet Calendar"
              secondaryText="Add calendar from URL or paste ICS content"
              iconProps={{ iconName: 'World' }}
              onClick={() => this.handleSelectAddType('ics')}
              style={{ textAlign: 'left', height: 'auto', padding: '12px' }}
            />
          </Stack>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton text="Cancel" onClick={this.handleCloseAddDialog} />
          </Stack>
        </Stack>
      );
    }

    // Show flow for selected type
    if (!addingCalendarType) {
      return <div>Unknown calendar type</div>;
    }

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        {this.renderNavigationHeader()}
        {addingCalendarType === 'sharepoint' && this.renderSharePointFlow()}
        {addingCalendarType === 'exchange' && this.renderExchangeFlow()}
        {addingCalendarType === 'ics' && this.renderIcsFlow()}
      </Stack>
    );
  };

  private renderExchangeCalendarItem = (calendar: IExchangeCalendar): React.ReactElement => {
    const isEnabled = this.isExchangeCalendarEnabled(calendar.id);

    return (
      <div key={calendar.id} style={{
        border: '1px solid #edebe9',
        borderRadius: 4,
        padding: 12
      }}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <div style={{
            width: 16,
            height: 16,
            backgroundColor: calendar.hexColor,
            borderRadius: 2,
            flexShrink: 0
          }} />
          <div style={{ flex: 1 }}>
            <strong>{calendar.name}</strong>
            {calendar.isDefaultCalendar && (
              <span style={{ fontSize: 11, color: '#605e5c', marginLeft: 8 }}>(Default)</span>
            )}
          </div>
          <Toggle
            checked={isEnabled}
            onChange={(_, checked) => this.handleToggleExchangeCalendar(calendar.id, !!checked)}
          />
        </Stack>
      </div>
    );
  };

  private renderCalendarSource = (source: ICalendarSource): React.ReactElement => {
    const { editingSourceId } = this.state;

    return (
      <div key={source.id} style={{
        border: '1px solid #edebe9',
        borderRadius: 4,
        padding: 12,
        backgroundColor: editingSourceId === source.id ? '#f3f2f1' : 'transparent'
      }}>
        {editingSourceId === source.id && source.sourceType === 'ics' ? (
          <Stack tokens={{ childrenGap: 8 }}>
            <TextField
              label="Name"
              value={source.name}
              onChange={(_, value) => this.handleUpdateSource(source.id, { name: value || '' })}
            />
            <Label>ICS Source (choose one):</Label>
            <TextField
              label="ICS URL"
              value={source.url || ''}
              onChange={(_, value) => this.handleUpdateSource(source.id, { url: value || '' })}
              multiline
              rows={2}
              placeholder="https://example.com/calendar.ics"
            />
            <div style={{ fontSize: 12, color: '#605e5c', margin: '8px 0' }}>OR</div>
            <TextField
              label="Paste ICS Content"
              value={source.rawContent || ''}
              onChange={(_, value) => this.handleUpdateSource(source.id, { rawContent: value || '' })}
              multiline
              rows={5}
              placeholder="BEGIN:VCALENDAR&#10;VERSION:2.0&#10;..."
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
        ) : editingSourceId === source.id ? (
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
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              onClick={() => this.toggleEdit(source.id)}
            />
            <Toggle
              checked={source.isEnabled}
              onChange={(_, checked) => this.handleUpdateSource(source.id, { isEnabled: !!checked })}
            />
          </Stack>
        )}
      </div>
    );
  };

  public render(): React.ReactElement {
    const { isOpen, onDismiss } = this.props;
    const { settings, showAddDialog, userExchangeCalendars, userExchangeCalendarsLoading } = this.state;

    return (
      <Panel
        isOpen={isOpen}
        onDismiss={onDismiss}
        type={PanelType.medium}
        headerText={showAddDialog ? "Add Calendar" : "Calendar Settings"}
        onRenderFooterContent={!showAddDialog ? this.onRenderFooterContent : undefined}
        isFooterAtBottom={true}
      >
        {showAddDialog ? (
          this.renderAddCalendarFlow()
        ) : (
          <Stack tokens={{ childrenGap: 16 }}>
            {/* Add Calendar Button at Top */}
            <div>
              <PrimaryButton 
                text="Add Calendar" 
                iconProps={{ iconName: 'Add' }}
                onClick={this.handleOpenAddDialog} 
              />
            </div>

            {/* Calendar Sources Section */}
            <div>
              <Stack tokens={{ childrenGap: 20 }}>
                {/* Exchange Calendars (Auto-loaded from current user) */}
                {userExchangeCalendarsLoading && (
                  <Spinner size={SpinnerSize.small} label="Loading your calendars..." />
                )}
                {!userExchangeCalendarsLoading && userExchangeCalendars.length > 0 && (
                  <div>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ marginBottom: 8 }}>
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ margin: 0 }}>
                        <Icon iconName="OutlookLogo" style={{ fontSize: 16 }} />&nbsp;
                        <Label style={{ fontWeight: 600, fontSize: 14, margin: 0 }}>Outlook</Label>
                      </Stack>
                      <DefaultButton
                        text="Manage"
                        iconProps={{ iconName: 'OpenInNewWindow' }}
                        onClick={() => window.open('https://outlook.cloud.microsoft/calendar/', '_blank', 'noopener,noreferrer')}
                        styles={{ root: { height: 24, minWidth: 0, padding: '0 8px' } }}
                      />
                    </Stack>
                    <Stack tokens={{ childrenGap: 10 }}>
                      {userExchangeCalendars.map(calendar => (
                        this.renderExchangeCalendarItem(calendar)
                      ))}
                    </Stack>
                  </div>
                )}

                {/* SharePoint Calendars */}
                {settings.sources.filter(s => s.sourceType === 'sharepoint').length > 0 && (
                  <div>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ marginBottom: 8 }}>
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ margin: 0 }}>
                        <Icon iconName="SharepointLogo" style={{ fontSize: 16 }} />&nbsp;
                        <Label style={{ fontWeight: 600, fontSize: 14, margin: 0 }}>SharePoint</Label>
                      </Stack>
                    </Stack>
                    <Stack tokens={{ childrenGap: 10 }}>
                      {settings.sources.filter(s => s.sourceType === 'sharepoint').map(source => (
                        this.renderCalendarSource(source)
                      ))}
                    </Stack>
                  </div>
                )}

                {/* Web (ICS) Calendars */}
                {settings.sources.filter(s => s.sourceType === 'ics').length > 0 && (
                  <div>
                    <Label style={{ fontWeight: 600, fontSize: 14, marginBottom: 8 }}>Web</Label>
                    <Stack tokens={{ childrenGap: 10 }}>
                      {settings.sources.filter(s => s.sourceType === 'ics').map(source => (
                        this.renderCalendarSource(source)
                      ))}
                    </Stack>
                  </div>
                )}

                {/* Show message if no calendars */}
                {!userExchangeCalendarsLoading && userExchangeCalendars.length === 0 && settings.sources.length === 0 && (
                  <Label style={{ color: '#605e5c', fontStyle: 'italic' }}>
                    No calendars found. Click &quot;Add Calendar&quot; to add SharePoint or Web calendars.
                  </Label>
                )}
              </Stack>
            </div>
          </Stack>
        )}
      </Panel>
    );
  }


}
