import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { ColorPicker } from '@fluentui/react/lib/ColorPicker';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { Stack } from '@fluentui/react/lib/Stack';
import { Label } from '@fluentui/react/lib/Label';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Icon } from '@fluentui/react/lib/Icon';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { HttpClient, type MSGraphClientV3 } from '@microsoft/sp-http';
import { ICalendarSource, ICalendarSettings, CalendarSourceType } from '../models/ICalendarSettings';
import * as strings from 'MyCalendarsWebPartStrings';
import { ExchangeCalendarService, IExchangeCalendar } from '../services/ExchangeCalendarService';
import { SharePointCalendarService, ISharePointSite, ISharePointList } from '../services/SharePointCalendarService';
import { PlannerTaskService, IPlannerPlan } from '../services/PlannerTaskService';
import { UnifiedGroupCalendarService, IUnifiedGroupItem } from '../services/UnifiedGroupCalendarService';

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
  addingCalendarStep: 'initial' | 'sharepoint-site' | 'sharepoint-list' | 'sharepoint-fields' | 'exchange-calendar' | 'exchange-mailbox' | 'ics' | 'planner-plan' | 'planner-options' | 'teams-shifts' | 'unified-group-select';
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
  // Planner flow
  plannerPlans: IPlannerPlan[];
  plannerPlansLoading: boolean;
  plannerSelectedPlanId: string | undefined;
  plannerAssignedToMeOnly: boolean;
  plannerShowCompleted: boolean;
  plannerShowLogo: boolean;
  teamsShiftsShowLogo: boolean;
  // M365 Groups/Teams flow
  unifiedGroups: IUnifiedGroupItem[];
  unifiedGroupsLoading: boolean;
  unifiedGroupsSelection: Record<string, boolean>;
  // Color for new calendar
  newCalendarColor: string;
  newCalendarName: string;
  // Accordion expanded state per section
  expandedSections: Record<string, boolean>;
}

interface IGraphColumn {
  name?: string;
  displayName?: string;
  columnGroup?: string;
}

export class SettingsPanel extends React.Component<ISettingsPanelProps, ISettingsPanelState> {
  private exchangeService: ExchangeCalendarService | null = null;
  private sharePointService: SharePointCalendarService | null = null;
  private plannerService: PlannerTaskService | null = null;
  private unifiedGroupService: UnifiedGroupCalendarService | null = null;
  private readonly SITES_PER_PAGE = 20;

  constructor(props: ISettingsPanelProps) {
    super(props);

    if (props.httpClient) {
      this.exchangeService = new ExchangeCalendarService(props.httpClient, props.graphClient);
      this.sharePointService = new SharePointCalendarService(props.httpClient, props.graphClient);
      this.plannerService = new PlannerTaskService(props.httpClient, props.graphClient);
      this.unifiedGroupService = new UnifiedGroupCalendarService(props.httpClient, props.graphClient);

      // If graphClient is provided, set it on the services
      if (props.graphClient) {
        this.exchangeService.setGraphClient(props.graphClient);
        this.sharePointService.setGraphClient(props.graphClient);
        this.plannerService.setGraphClient(props.graphClient);
        this.unifiedGroupService.setGraphClient(props.graphClient);
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
      // Planner
      plannerPlans: [],
      plannerPlansLoading: false,
      plannerSelectedPlanId: undefined,
      plannerAssignedToMeOnly: false,
      plannerShowCompleted: true,
      plannerShowLogo: true,
      teamsShiftsShowLogo: true,
      unifiedGroups: [],
      unifiedGroupsLoading: false,
      unifiedGroupsSelection: {},
      // New calendar
      newCalendarColor: props.settings.organizationPrimaryColor || '#0078d4',
      newCalendarName: '',
      expandedSections: {
        outlook: true,
        sharepoint: false,
        planner: false,
        unifiedGroup: false,
        teamsShifts: false
      }
    };
  }

  public componentDidMount(): void {
    // Initialize GraphClient if available
    if (this.props.graphClient) {
      this.initializeGraphClient(this.props.graphClient);
    }
  }

  public componentDidUpdate(prevProps: ISettingsPanelProps): void {
    // If graphClient becomes available, initialize it
    if (this.props.graphClient && !prevProps.graphClient) {
      this.initializeGraphClient(this.props.graphClient);
    }

    // Reset state when panel is opened
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

  private initializeGraphClient(client: MSGraphClientV3): void {
    if (this.sharePointService) {
      this.sharePointService.setGraphClient(client);
    }
    if (this.exchangeService) {
      this.exchangeService.setGraphClient(client);
    }
    if (this.plannerService) {
      this.plannerService.setGraphClient(client);
    }
    if (this.unifiedGroupService) {
      this.unifiedGroupService.setGraphClient(client);
    }
    // Load user's Exchange calendars automatically
    this.loadUserExchangeCalendars().catch(err => console.error('Failed to load Exchange calendars:', err));
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
      teamsShiftsShowLogo: true,
      unifiedGroups: [],
      unifiedGroupsLoading: false,
      unifiedGroupsSelection: {},
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
    } else if (type === 'planner') {
      this.setState({ addingCalendarType: type, addingCalendarStep: 'planner-plan', plannerPlansLoading: true });
      const plans = await this.plannerService?.getUserPlans() || [];
      this.setState({ plannerPlans: plans, plannerPlansLoading: false });
    } else if (type === 'unifiedGroup') {
      this.setState({
        addingCalendarType: type,
        addingCalendarStep: 'unified-group-select',
        unifiedGroupsLoading: true,
        unifiedGroupsSelection: {},
        newCalendarColor: this.props.settings.organizationPrimaryColor || '#0078d4'
      });
      await this.loadUnifiedGroups();
    } else if (type === 'teamsShifts') {
      this.setState({
        addingCalendarType: type,
        addingCalendarStep: 'teams-shifts',
        newCalendarName: 'Teams Shifts',
        newCalendarColor: this.props.settings.organizationPrimaryColor || '#0078d4',
        teamsShiftsShowLogo: true
      });
    }
  };

  private loadUnifiedGroups = async (): Promise<void> => {
    if (!this.unifiedGroupService) {
      this.setState({ unifiedGroupsLoading: false });
      return;
    }

    try {
      const [groups, joinedTeamIds] = await Promise.all([
        this.unifiedGroupService.getUnifiedGroups(),
        this.unifiedGroupService.getJoinedTeamIds()
      ]);

      const mappedGroups = groups
        .map(group => ({
          ...group,
          isTeam: joinedTeamIds.has(group.id)
        }))
        .sort((a, b) => a.displayName.localeCompare(b.displayName));

      this.setState({
        unifiedGroups: mappedGroups,
        unifiedGroupsLoading: false
      });
    } catch (error) {
      console.error('Failed to load unified groups:', error);
      this.setState({
        unifiedGroups: [],
        unifiedGroupsLoading: false
      });
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
      plannerPlans: [],
      plannerPlansLoading: false,
      plannerSelectedPlanId: undefined,
      plannerAssignedToMeOnly: false,
      plannerShowCompleted: true,
      plannerShowLogo: true,
      teamsShiftsShowLogo: true,
      unifiedGroups: [],
      unifiedGroupsLoading: false,
      unifiedGroupsSelection: {},
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
    } else if (addingCalendarType === 'planner') {
      if (addingCalendarStep === 'planner-options') {
        this.setState({ addingCalendarStep: 'planner-plan' });
      } else if (addingCalendarStep === 'planner-plan') {
        this.handleBackToTypeSelection();
      }
    } else if (addingCalendarType === 'unifiedGroup') {
      this.handleBackToTypeSelection();
    } else if (addingCalendarType === 'teamsShifts') {
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
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      this.fetchSharePointListFields(list);
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
                  onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = '#e1dfdd'; }}
                  onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = '#f3f2f1'; }}
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

  // Planner flow
  private handleSelectPlannerPlan = (planId: string, planTitle: string): void => {
    this.setState({
      plannerSelectedPlanId: planId,
      addingCalendarStep: 'planner-options',
      newCalendarName: planTitle
    });
  };

  private handleConfirmPlannerPlan = (): void => {
    const selectedPlan = this.state.plannerPlans.find(p => p.id === this.state.plannerSelectedPlanId);
    if (!selectedPlan || !this.state.newCalendarName.trim()) {
      return;
    }

    const newSource: ICalendarSource = {
      id: this.generateId(),
      sourceType: 'planner',
      name: this.state.newCalendarName,
      color: this.state.newCalendarColor,
      isEnabled: true,
      plannerPlanId: selectedPlan.id,
      plannerPlanTitle: selectedPlan.title,
      plannerAssignedToMeOnly: this.state.plannerAssignedToMeOnly,
      showCompletedTasks: this.state.plannerShowCompleted,
      showSourceLogo: this.state.plannerShowLogo
    };

    const settings = {
      ...this.state.settings,
      sources: [...this.state.settings.sources, newSource]
    };

    this.setState({ settings }, () => this.handleCloseAddDialog());
  };

  private handleConfirmTeamsShifts = (): void => {
    if (!this.state.newCalendarName.trim()) {
      return;
    }

    const newSource: ICalendarSource = {
      id: this.generateId(),
      sourceType: 'teamsShifts',
      name: this.state.newCalendarName.trim(),
      color: this.state.newCalendarColor,
      isEnabled: true,
      showSourceLogo: this.state.teamsShiftsShowLogo
    };

    const settings = {
      ...this.state.settings,
      sources: [...this.state.settings.sources, newSource]
    };

    this.setState({ settings }, () => this.handleCloseAddDialog());
  };

  private handleToggleUnifiedGroupSelection = (groupId: string, checked?: boolean): void => {
    this.setState(prev => ({
      unifiedGroupsSelection: {
        ...prev.unifiedGroupsSelection,
        [groupId]: !!checked
      }
    }));
  };

  private handleConfirmUnifiedGroups = (): void => {
    const selectedIds = Object.keys(this.state.unifiedGroupsSelection)
      .filter(id => this.state.unifiedGroupsSelection[id]);

    if (selectedIds.length === 0) {
      return;
    }

    const selectedGroups = this.state.unifiedGroups.filter(group => selectedIds.indexOf(group.id) >= 0);
    const newSources = selectedGroups.map(group => ({
      id: this.generateId(),
      sourceType: 'unifiedGroup' as const,
      name: group.displayName,
      color: this.state.newCalendarColor,
      isEnabled: true,
      groupId: group.id,
      showSourceLogo: true
    }));

    const settings = {
      ...this.state.settings,
      sources: [...this.state.settings.sources, ...newSources]
    };

    this.setState({ settings }, () => this.handleCloseAddDialog());
  };

  private renderUnifiedGroupsFlow = (): React.ReactElement => {
    const { unifiedGroups, unifiedGroupsLoading, unifiedGroupsSelection } = this.state;
    const selectedCount = Object.keys(unifiedGroupsSelection).filter(id => unifiedGroupsSelection[id]).length;

    if (unifiedGroupsLoading) {
      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>Loading groups and teams...</Label>
          <Spinner size={SpinnerSize.large} />
        </Stack>
      );
    }

    if (unifiedGroups.length === 0) {
      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>No M365 groups found</Label>
          <div>You don&apos;t have access to any Microsoft 365 groups.</div>
          <DefaultButton text="Cancel" onClick={this.handleCloseAddDialog} />
        </Stack>
      );
    }

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>Select one or more groups or teams:</Label>
        <Stack tokens={{ childrenGap: 8 }}>
          {unifiedGroups.map(group => (
            <Stack
              key={group.id}
              horizontal
              verticalAlign="center"
              tokens={{ childrenGap: 8 }}
              style={{
                padding: '8px 12px',
                border: '1px solid #edebe9',
                borderRadius: 4,
                backgroundColor: unifiedGroupsSelection[group.id] ? '#f3f2f1' : 'white'
              }}
            >
              <Icon iconName={group.isTeam ? 'TeamsLogo' : 'Group'} style={{ fontSize: 16 }} />
              <Checkbox
                label={group.displayName}
                checked={!!unifiedGroupsSelection[group.id]}
                onChange={(_, checked) => this.handleToggleUnifiedGroupSelection(group.id, checked)}
              />
            </Stack>
          ))}
        </Stack>
        <div>
          <Label>Color</Label>
          <ColorPicker
            color={this.state.newCalendarColor}
            onChange={(_, color) => this.setState({ newCalendarColor: `#${color.hex}` })}
            alphaType="none"
          />
        </div>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text={selectedCount > 1 ? 'Add Calendars' : 'Add Calendar'}
            onClick={this.handleConfirmUnifiedGroups}
            disabled={selectedCount === 0}
          />
          <DefaultButton text="Cancel" onClick={this.handleCloseAddDialog} />
        </Stack>
      </Stack>
    );
  };

  private renderTeamsShiftsFlow = (): React.ReactElement => {
    const hasValidInput = this.state.newCalendarName.trim();

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>Teams Shifts</Label>
        <TextField
          label="Calendar Name"
          value={this.state.newCalendarName}
          onChange={(_, value) => this.setState({ newCalendarName: value || '' })}
          placeholder="e.g. Teams Shifts"
          required
        />

        <Toggle
          label="Bron logo tonen"
          checked={this.state.teamsShiftsShowLogo}
          onChange={(_, checked) => this.setState({ teamsShiftsShowLogo: checked || false })}
          onText="Ja"
          offText="Nee"
        />

        <ColorPicker
          color={this.state.newCalendarColor}
          onChange={(_, color) => this.setState({ newCalendarColor: `#${color.hex}` })}
          alphaType="none"
          showPreview={true}
        />

        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text="Add Calendar"
            onClick={this.handleConfirmTeamsShifts}
            disabled={!hasValidInput}
          />
          <DefaultButton text="Cancel" onClick={this.handleCloseAddDialog} />
        </Stack>
      </Stack>
    );
  };

  private renderPlannerFlow = (): React.ReactElement => {
    const { addingCalendarStep, plannerPlans, plannerPlansLoading } = this.state;

    // Step 1: Select Plan
    if (addingCalendarStep === 'planner-plan') {
      if (plannerPlansLoading) {
        return (
          <Stack tokens={{ childrenGap: 12 }}>
            <Label>Loading your Planner plans...</Label>
            <Spinner size={SpinnerSize.large} />
          </Stack>
        );
      }

      if (plannerPlans.length === 0) {
        return (
          <Stack tokens={{ childrenGap: 12 }}>
            <Label>No Planner plans found</Label>
            <div>You don&apos;t have access to any Planner plans, or there are no plans in your organization.</div>
            <DefaultButton text="Cancel" onClick={this.handleCloseAddDialog} />
          </Stack>
        );
      }

      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>Select a Planner plan to add:</Label>
          <Stack tokens={{ childrenGap: 8 }}>
            {plannerPlans.map(plan => (
              <div
                key={plan.id}
                onClick={() => this.handleSelectPlannerPlan(plan.id, plan.title)}
                style={{
                  padding: '12px',
                  border: '1px solid #ddd',
                  borderRadius: '4px',
                  cursor: 'pointer',
                  backgroundColor: this.state.plannerSelectedPlanId === plan.id ? '#f3f2f1' : 'white'
                }}
              >
                <Icon iconName="PlannerLogo" style={{ marginRight: 8, fontSize: 16 }} />
                <strong>{plan.title}</strong>
              </div>
            ))}
          </Stack>
        </Stack>
      );
    }

    // Step 2: Configure Options
    if (addingCalendarStep === 'planner-options') {
      const selectedPlan = plannerPlans.find(p => p.id === this.state.plannerSelectedPlanId);
      const hasValidInput = this.state.newCalendarName.trim();

      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>Configure Planner integration</Label>

          {selectedPlan && (
            <div style={{ padding: '8px', backgroundColor: '#f3f2f1', borderRadius: '4px' }}>
              <Icon iconName="PlannerLogo" style={{ marginRight: 8 }} />
              <strong>{selectedPlan.title}</strong>
            </div>
          )}

          <TextField
            label="Calendar Name"
            value={this.state.newCalendarName}
            onChange={(_, value) => this.setState({ newCalendarName: value || '' })}
            placeholder="e.g. Project Tasks"
            required
          />

          <Toggle
            label="Alleen aan mij toegewezen taken"
            checked={this.state.plannerAssignedToMeOnly}
            onChange={(_, checked) => this.setState({ plannerAssignedToMeOnly: checked || false })}
            onText="Ja"
            offText="Nee"
          />

          <Toggle
            label="Voltooide taken weergeven"
            checked={this.state.plannerShowCompleted}
            onChange={(_, checked) => this.setState({ plannerShowCompleted: checked || false })}
            onText="Ja"
            offText="Nee"
          />

          <Toggle
            label="Bron logo tonen"
            checked={this.state.plannerShowLogo}
            onChange={(_, checked) => this.setState({ plannerShowLogo: checked || false })}
            onText="Ja"
            offText="Nee"
          />

          <ColorPicker
            color={this.state.newCalendarColor}
            onChange={(_, color) => this.setState({ newCalendarColor: `#${color.hex}` })}
            alphaType="none"
            showPreview={true}
          />

          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text="Add Calendar"
              onClick={this.handleConfirmPlannerPlan}
              disabled={!hasValidInput}
            />
            <DefaultButton text="Cancel" onClick={this.handleCloseAddDialog} />
          </Stack>
        </Stack>
      );
    }

    return <div>Unknown step</div>;
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
              text="Microsoft Planner"
              secondaryText="Voeg taken toe vanuit een Planner plan"
              iconProps={{ iconName: 'PlannerLogo' }}
              onClick={() => this.handleSelectAddType('planner')}
              style={{ textAlign: 'left', height: 'auto', padding: '12px' }}
            />
            <PrimaryButton
              text="M365 Group"
              secondaryText="Add a calendar from a Microsoft 365 group"
              iconProps={{ iconName: 'Group' }}
              onClick={() => this.handleSelectAddType('unifiedGroup')}
              style={{ textAlign: 'left', height: 'auto', padding: '12px' }}
            />
            <PrimaryButton
              text="Teams"
              secondaryText="Add a calendar from a Team"
              iconProps={{ iconName: 'TeamsLogo' }}
              onClick={() => this.handleSelectAddType('unifiedGroup')}
              style={{ textAlign: 'left', height: 'auto', padding: '12px' }}
            />
            <PrimaryButton
              text="Teams Shifts"
              secondaryText="Toon diensten uit Teams shifts"
              iconProps={{ iconName: 'Clock' }}
              onClick={() => this.handleSelectAddType('teamsShifts')}
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
        {addingCalendarType === 'planner' && this.renderPlannerFlow()}
        {addingCalendarType === 'unifiedGroup' && this.renderUnifiedGroupsFlow()}
        {addingCalendarType === 'teamsShifts' && this.renderTeamsShiftsFlow()}
        {addingCalendarType === 'ics' && this.renderIcsFlow()}
      </Stack>
    );
  };

  private renderSourceSection = (params: {
    sectionKey: string;
    icon: string;
    iconBg: string;
    iconColor: string;
    title: string;
    subtitle: string;
    showLogoValue?: boolean;
    onShowLogoChange?: (checked: boolean) => void;
    headerActions?: React.ReactNode;
    children: React.ReactNode;
  }): React.ReactElement => {
    const { sectionKey, icon, iconBg, iconColor, title, subtitle, showLogoValue, onShowLogoChange, headerActions, children } = params;
    const isExpanded = this.state.expandedSections[sectionKey] !== false;
    const toggleExpanded = (): void => {
      this.setState(prev => ({
        expandedSections: { ...prev.expandedSections, [sectionKey]: !isExpanded }
      }));
    };

    return (
      <div style={{ border: '1px solid #edebe9', overflow: 'hidden' }}>
        {/* Section header */}
        <div
          role="button"
          tabIndex={0}
          aria-expanded={isExpanded}
          style={{
            display: 'flex',
            alignItems: 'center',
            padding: '12px 16px',
            gap: 12,
            cursor: 'pointer',
            userSelect: 'none',
            backgroundColor: 'rgba(0, 0, 0, 0.04)'
          }}
          onClick={toggleExpanded}
          onKeyDown={(e: React.KeyboardEvent) => { if (e.key === 'Enter' || e.key === ' ') { toggleExpanded(); e.preventDefault(); } }}
        >
          <div style={{
            width: 40, height: 40,
            backgroundColor: iconBg,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            flexShrink: 0
          }}>
            <Icon iconName={icon} style={{ fontSize: 20, color: iconColor }} />
          </div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontWeight: 600, fontSize: 14 }}>{title}</div>
            <div style={{ fontSize: 11, opacity: 0.55, textTransform: 'uppercase', letterSpacing: '0.5px', marginTop: 1 }}>{subtitle}</div>
          </div>
          {(headerActions !== undefined || showLogoValue !== undefined) && (
            <div
              style={{ display: 'flex', alignItems: 'center', gap: 8 }}
              onClick={e => e.stopPropagation()}
            >
              {headerActions}
              {showLogoValue !== undefined && onShowLogoChange && (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                  <span style={{ fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px', opacity: 0.6 }}>Show logos</span>
                  <Toggle
                    checked={showLogoValue}
                    onChange={(_, checked) => onShowLogoChange(!!checked)}
                    styles={{ root: { margin: 0 } }}
                  />
                </Stack>
              )}
            </div>
          )}
          <Icon
            iconName="ChevronDown"
            style={{
              fontSize: 14,
              opacity: 0.55,
              transform: isExpanded ? 'rotate(0deg)' : 'rotate(-90deg)',
              transition: 'transform 0.15s ease',
              flexShrink: 0
            }}
          />
        </div>
        {/* Expandable content */}
        {isExpanded && (
          <div style={{ padding: '8px 16px 16px', display: 'flex', flexDirection: 'column', gap: 6 }}>
            {children}
          </div>
        )}
      </div>
    );
  };

  private renderExchangeCalendarItem = (calendar: IExchangeCalendar, index: number = 0): React.ReactElement => {
    const isEnabled = this.isExchangeCalendarEnabled(calendar.id);

    return (
      <div key={calendar.id} style={{ backgroundColor: index % 2 === 1 ? 'rgba(0, 0, 0, 0.02)' : 'transparent', borderRadius: 4, padding: '4px 6px' }}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <div style={{
            width: 16,
            height: 16,
            backgroundColor: calendar.hexColor,
            borderRadius: 2,
            flexShrink: 0
          }} />
          <div style={{ flex: 1 }}>
            <strong style={{ fontSize: 13 }}>{calendar.name}</strong>
            {calendar.isDefaultCalendar && (
              <span style={{ fontSize: 11, color: '#605e5c', marginLeft: 8 }}>(Default)</span>
            )}
          </div>
          <IconButton
            iconProps={{ iconName: 'Edit' }}
            title="Edit"
            onClick={() => window.open('https://outlook.cloud.microsoft/calendar/', '_blank', 'noopener,noreferrer')}
            styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 14 } }}
          />
          <IconButton
            iconProps={{ iconName: isEnabled ? 'View' : 'Hide' }}
            title={isEnabled ? 'Hide calendar' : 'Show calendar'}
            onClick={() => this.handleToggleExchangeCalendar(calendar.id, !isEnabled)}
            styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 16, color: isEnabled ? 'inherit' : '#a19f9d' } }}
          />
        </Stack>
      </div>
    );
  };

  private renderCalendarSource = (source: ICalendarSource, index: number = 0): React.ReactElement => {
    const { editingSourceId } = this.state;
    const isEditing = editingSourceId === source.id;

    return (
      <div
        key={source.id}
        style={{
          borderRadius: 4,
          padding: isEditing ? 12 : '4px 6px',
          backgroundColor: isEditing ? '#f3f2f1' : (index % 2 === 1 ? 'rgba(0, 0, 0, 0.02)' : 'transparent'),
          border: isEditing ? '1px solid #edebe9' : 'none'
        }}
      >
        {isEditing ? (
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
              <strong style={{ fontSize: 13 }}>{source.name}</strong>
            </div>
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              title="Edit"
              onClick={() => this.toggleEdit(source.id)}
              styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 14 } }}
            />
            <IconButton
              iconProps={{ iconName: source.isEnabled ? 'View' : 'Hide' }}
              title={source.isEnabled ? 'Hide calendar' : 'Show calendar'}
              onClick={() => this.handleUpdateSource(source.id, { isEnabled: !source.isEnabled })}
              styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 16, color: source.isEnabled ? 'inherit' : '#a19f9d' } }}
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
              <Stack tokens={{ childrenGap: 12 }}>

                {/* Outlook */}
                {(userExchangeCalendarsLoading || userExchangeCalendars.length > 0) && this.renderSourceSection({
                  sectionKey: 'outlook',
                  icon: 'OutlookLogo',
                  iconBg: 'rgba(0, 120, 212, 0.12)',
                  iconColor: '#0078d4',
                  title: 'Outlook',
                  subtitle: 'Personal & shared mailboxes',
                  showLogoValue: settings.exchangeShowSourceLogo ?? true,
                  onShowLogoChange: (checked) => this.setState({ settings: { ...settings, exchangeShowSourceLogo: checked } }),
                  headerActions: (
                    <DefaultButton
                      text="Manage"
                      iconProps={{ iconName: 'OpenInNewWindow' }}
                      onClick={() => window.open('https://outlook.cloud.microsoft/calendar/', '_blank', 'noopener,noreferrer')}
                      styles={{ root: { height: 28, minWidth: 0, padding: '0 8px', fontSize: 12 } }}
                    />
                  ),
                  children: userExchangeCalendarsLoading
                    ? <Spinner size={SpinnerSize.small} label="Loading your calendars..." />
                    : userExchangeCalendars.map((calendar, i) => this.renderExchangeCalendarItem(calendar, i))
                })}

                {/* SharePoint */}
                {settings.sources.filter(s => s.sourceType === 'sharepoint').length > 0 && this.renderSourceSection({
                  sectionKey: 'sharepoint',
                  icon: 'SharepointLogo',
                  iconBg: 'rgba(3, 129, 134, 0.12)',
                  iconColor: '#038186',
                  title: 'SharePoint',
                  subtitle: 'Site lists, calendars & events',
                  showLogoValue: settings.sharePointShowSourceLogo ?? true,
                  onShowLogoChange: (checked) => this.setState({ settings: { ...settings, sharePointShowSourceLogo: checked } }),
                  children: settings.sources.filter(s => s.sourceType === 'sharepoint').map((source, i) => this.renderCalendarSource(source, i))
                })}

                {/* Planner */}
                {settings.sources.filter(s => s.sourceType === 'planner').length > 0 && this.renderSourceSection({
                  sectionKey: 'planner',
                  icon: 'PlannerLogo',
                  iconBg: 'rgba(16, 124, 65, 0.12)',
                  iconColor: '#107c41',
                  title: 'Planner',
                  subtitle: 'Task deadlines & milestones',
                  showLogoValue: settings.plannerShowSourceLogo ?? true,
                  onShowLogoChange: (checked) => this.setState({ settings: { ...settings, plannerShowSourceLogo: checked } }),
                  children: settings.sources.filter(s => s.sourceType === 'planner').map((source, i) => this.renderCalendarSource(source, i))
                })}

                {/* Groups & Teams */}
                {settings.sources.filter(s => s.sourceType === 'unifiedGroup').length > 0 && this.renderSourceSection({
                  sectionKey: 'unifiedGroup',
                  icon: 'Group',
                  iconBg: 'rgba(91, 95, 199, 0.12)',
                  iconColor: '#5b5fc7',
                  title: 'Groups & Teams',
                  subtitle: 'Shared group calendars',
                  showLogoValue: settings.unifiedGroupShowSourceLogo ?? true,
                  onShowLogoChange: (checked) => this.setState({ settings: { ...settings, unifiedGroupShowSourceLogo: checked } }),
                  children: settings.sources.filter(s => s.sourceType === 'unifiedGroup').map((source, i) => this.renderCalendarSource(source, i))
                })}

                {/* Teams Shifts */}
                {settings.sources.filter(s => s.sourceType === 'teamsShifts').length > 0 && this.renderSourceSection({
                  sectionKey: 'teamsShifts',
                  icon: 'Clock',
                  iconBg: 'rgba(74, 79, 190, 0.12)',
                  iconColor: '#4a4fbe',
                  title: 'Teams Shifts',
                  subtitle: 'Work schedules & rotas',
                  showLogoValue: settings.teamsShiftsShowSourceLogo ?? true,
                  onShowLogoChange: (checked) => this.setState({ settings: { ...settings, teamsShiftsShowSourceLogo: checked } }),
                  children: settings.sources.filter(s => s.sourceType === 'teamsShifts').map((source, i) => this.renderCalendarSource(source, i))
                })}

                {/* Empty state */}
                {!userExchangeCalendarsLoading && userExchangeCalendars.length === 0 && settings.sources.length === 0 && (
                  <Label style={{ color: '#605e5c', fontStyle: 'italic' }}>
                    No calendars found. Click &quot;Add Calendar&quot; to add a calendar source.
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
