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
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { HttpClient, type MSGraphClientV3 } from '@microsoft/sp-http';
import { ICalendarSource, ICalendarSettings, CalendarSourceType } from '../models/ICalendarSettings';
import * as strings from 'MyCalendarsWebPartStrings';
import { ExchangeCalendarService, IExchangeCalendar } from '../services/ExchangeCalendarService';
import { SharePointCalendarService, ISharePointSite, ISharePointList } from '../services/SharePointCalendarService';
import { PlannerTaskService, IPlannerPlan } from '../services/PlannerTaskService';
import { UnifiedGroupCalendarService, IUnifiedGroupItem } from '../services/UnifiedGroupCalendarService';
import { formatCalendarTime } from './views/calendarFormatting';
import { getBulkVisibilityTarget, getGroupVisibilityState, type GroupVisibilityState, setOutlookVisibility, setSharePointVisibility } from './calendarVisibility';
import { formatLocalizedString } from '../utils/localization';
import { findBestMatchingFieldKey, getFieldCandidates } from '../utils/sharePointFieldCandidates';

export interface ISettingsPanelProps {
  isOpen: boolean;
  onDismiss: () => void;
  settings: ICalendarSettings;
  onSave: (settings: ICalendarSettings) => void;
  onReset?: () => void;
  httpClient?: HttpClient;
  graphClient?: MSGraphClientV3;
  locale?: string;
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
      this.sharePointService = new SharePointCalendarService(props.graphClient);
      this.plannerService = new PlannerTaskService(props.graphClient);
      this.unifiedGroupService = new UnifiedGroupCalendarService(props.graphClient);
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
      expandedSections: { outlook: true, sharepoint: false, planner: false, unifiedGroup: false, teamsShifts: false }
    };
  }
  public componentDidMount(): void {
    if (this.props.graphClient) this.initializeGraphClient(this.props.graphClient);
  }

  public componentDidUpdate(prevProps: ISettingsPanelProps): void {
    if (this.props.graphClient && !prevProps.graphClient) this.initializeGraphClient(this.props.graphClient);
    if (prevProps.isOpen !== this.props.isOpen && this.props.isOpen) {
      this.setState({
        settings: JSON.parse(JSON.stringify(this.props.settings)), editingSourceId: undefined, showAddDialog: false,
        addingCalendarType: undefined, addingCalendarStep: 'initial', spCurrentPage: 0
      }, () => {
        this.enrichSharePointSiteNames().catch(err => console.error('Failed to enrich SharePoint site names:', err));
      });
      this.loadUserExchangeCalendars().catch(err => console.error('Failed to reload Exchange calendars:', err));
    }
  }

  private initializeGraphClient(client: MSGraphClientV3): void {
    if (this.sharePointService) this.sharePointService.setGraphClient(client);
    if (this.exchangeService) this.exchangeService.setGraphClient(client);
    if (this.plannerService) this.plannerService.setGraphClient(client);
    if (this.unifiedGroupService) this.unifiedGroupService.setGraphClient(client);
    this.loadUserExchangeCalendars().catch(err => console.error('Failed to load Exchange calendars:', err));
    this.enrichSharePointSiteNames().catch(err => console.error('Failed to enrich SharePoint site names:', err));
  }

  private enrichSharePointSiteNames = async (): Promise<void> => {
    if (!this.sharePointService) return;
    const missingSiteIds = Array.from(new Set(this.state.settings.sources
      .filter(source => source.sourceType === 'sharepoint' && source.sharePointSiteId && !source.sharePointSiteName)
      .map(source => source.sharePointSiteId as string)));
    if (missingSiteIds.length === 0) return;
    const resolved = await Promise.all(missingSiteIds.map(async siteId => ({ siteId, site: await this.sharePointService?.getSite(siteId) })));
    const names = new Map(resolved.filter(item => item.site?.name).map(item => [item.siteId, item.site?.name as string]));
    if (names.size === 0) return;
    this.setState(prev => ({
      settings: {
        ...prev.settings,
        sources: prev.settings.sources.map(source => source.sourceType === 'sharepoint' && source.sharePointSiteId && !source.sharePointSiteName && names.has(source.sharePointSiteId)
          ? { ...source, sharePointSiteName: names.get(source.sharePointSiteId) }
          : source)
      }
    }));
  };

  private loadUserExchangeCalendars = async (): Promise<void> => {
    if (!this.exchangeService) return;
    this.setState({ userExchangeCalendarsLoading: true });
    try {
      const calendars = await this.exchangeService.getCalendars();
      this.setState({ userExchangeCalendars: calendars, userExchangeCalendarsLoading: false });
    } catch (error) {
      console.error('Error loading user Exchange calendars:', error);
      this.setState({ userExchangeCalendars: [], userExchangeCalendarsLoading: false });
    }
  };

  private generateId(): string { return `source_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`; }

  private handleOpenAddDialog = (): void => { this.setState({ showAddDialog: true }); };
  private handleCloseAddDialog = (): void => {
    this.setState({
      showAddDialog: false, addingCalendarType: undefined, addingCalendarStep: 'initial',
      spSites: [], spSiteFilter: '', spCurrentPage: 0, spSelectedSite: undefined,
      spLists: [], spSelectedList: undefined, exchangeCalendars: [], exchangeMailbox: '',
      exchangeMailboxResolved: false, exchangeSelectedCalendarId: undefined, icsUrl: '',
      teamsShiftsShowLogo: true, unifiedGroups: [], unifiedGroupsLoading: false,
      unifiedGroupsSelection: {}, newCalendarColor: this.props.settings.organizationPrimaryColor || '#0078d4', newCalendarName: ''
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
      this.setState({ addingCalendarType: type, addingCalendarStep: 'unified-group-select', unifiedGroupsLoading: true, unifiedGroupsSelection: {}, newCalendarColor: this.props.settings.organizationPrimaryColor || '#0078d4' });
      await this.loadUnifiedGroups();
    } else if (type === 'teamsShifts') {
      this.setState({ addingCalendarType: type, addingCalendarStep: 'teams-shifts', newCalendarName: strings.TeamsShiftsLabel, newCalendarColor: this.props.settings.organizationPrimaryColor || '#0078d4', teamsShiftsShowLogo: true });
    }
  };

  private loadUnifiedGroups = async (): Promise<void> => {
    if (!this.unifiedGroupService) { this.setState({ unifiedGroupsLoading: false }); return; }
    try {
      const [groups, joinedTeamIds] = await Promise.all([
        this.unifiedGroupService.getUnifiedGroups(),
        this.unifiedGroupService.getJoinedTeamIds()
      ]);
      const mappedGroups = groups.map(group => ({ ...group, isTeam: joinedTeamIds.has(group.id) })).sort((a, b) => a.displayName.localeCompare(b.displayName));
      this.setState({ unifiedGroups: mappedGroups, unifiedGroupsLoading: false });
    } catch (error) {
      console.error('Failed to load unified groups:', error);
      this.setState({ unifiedGroups: [], unifiedGroupsLoading: false });
    }
  };

  private handleBackToTypeSelection = (): void => {
    this.setState({
      addingCalendarStep: 'initial', spSites: [], spSitesLoading: false, spSiteFilter: '', spCurrentPage: 0,
      spSelectedSite: undefined, spLists: [], spListsLoading: false, spSelectedList: undefined,
      exchangeCalendars: [], exchangeCalendarsLoading: false, exchangeMailbox: '', exchangeMailboxResolved: false,
      exchangeSelectedCalendarId: undefined, icsUrl: '', plannerPlans: [], plannerPlansLoading: false,
      plannerSelectedPlanId: undefined, plannerAssignedToMeOnly: false, plannerShowCompleted: true,
      plannerShowLogo: true, teamsShiftsShowLogo: true, unifiedGroups: [], unifiedGroupsLoading: false,
      unifiedGroupsSelection: {}, newCalendarColor: this.props.settings.organizationPrimaryColor || '#0078d4', newCalendarName: ''
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
            title={strings.BackLabel}
            ariaLabel={strings.BackLabel}
            onClick={this.handleBackOneStep}
            styles={{ root: { height: 32 } }}
          />
        )}
        <IconButton
          iconProps={{ iconName: 'Home' }}
          title={strings.HomeLabel}
          ariaLabel={strings.HomeLabel}
          onClick={this.handleBackToTypeSelection}
          styles={{ root: { height: 32 } }}
        />
      </Stack>
    );
  };

  // SharePoint flow
  private handleSharePointFilterChange = (value?: string): void => { this.setState({ spSiteFilter: value || '', spCurrentPage: 0 }); };

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
          title: getFieldCandidates('title'),
          start: getFieldCandidates('start'),
          end: getFieldCandidates('end')
        };

        const guessedMapping = { ...this.state.spFieldMapping };
        if (!guessedMapping.titleField) {
          guessedMapping.titleField = findBestMatchingFieldKey(fieldOptions, candidateLists.title) as string | undefined;
        }
        if (!guessedMapping.startDateField) {
          guessedMapping.startDateField = findBestMatchingFieldKey(fieldOptions, candidateLists.start) as string | undefined;
        }
        if (!guessedMapping.endDateField) {
          guessedMapping.endDateField = findBestMatchingFieldKey(fieldOptions, candidateLists.end) as string | undefined;
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

  private handleConfirmSharePointCalendar = (): void => {
    if (!this.state.spSelectedSite || !this.state.spSelectedList) {
      return;
    }

    const newSourceId = this.generateId();
    const newSource: ICalendarSource = {
      id: newSourceId,
      userSourceId: newSourceId,
      origin: 'user',
      sourceType: 'sharepoint',
      name: this.state.newCalendarName,
      color: this.state.newCalendarColor,
      isEnabled: true,
      sharePointSiteId: this.state.spSelectedSite.id,
      sharePointSiteName: this.state.spSelectedSite.name,
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
  private handleExchangeMailboxChange = (value?: string): void => { this.setState({ exchangeMailbox: value || '' }); };

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
      alert(strings.MailboxUnavailableAlertLabel);
    }
  };

  private handleSelectExchangeCalendar = (calendar: IExchangeCalendar): void => {
    this.setState({ exchangeSelectedCalendarId: calendar.id, newCalendarName: calendar.name, newCalendarColor: calendar.hexColor });
  };

  private handleConfirmExchangeCalendar = (): void => {
    const newSourceId = this.generateId();
    const newSource: ICalendarSource = {
      id: newSourceId,
      userSourceId: newSourceId,
      origin: 'user',
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
    const settings = { ...this.state.settings, sources: this.state.settings.sources.map(s => s.id === id ? { ...s, ...updates } : s) };
    this.setState({ settings });
  };

  private handleToggleExchangeCalendar = (calendarId: string, isEnabled: boolean): void => {
    const settings = { ...this.state.settings, exchangeCalendarStates: { ...(this.state.settings.exchangeCalendarStates || {}), [calendarId]: isEnabled } };
    this.setState({ settings });
  };

  private handleTogglePlannerShowAll = (checked: boolean): void => { this.setState(prev => ({ settings: { ...prev.settings, plannerShowAllCalendars: checked } })); };
  private handleTogglePlannerAssignedToMeOnly = (checked: boolean): void => { this.setState(prev => ({ settings: { ...prev.settings, plannerShowAllAssignedToMeOnly: checked } })); };
  private handleToggleUnifiedGroupShowAll = (checked: boolean): void => { this.setState(prev => ({ settings: { ...prev.settings, unifiedGroupShowAllCalendars: checked } })); };
  private handleToggleTeamsShiftsShowAll = (checked: boolean): void => { this.setState(prev => ({ settings: { ...prev.settings, teamsShiftsShowAllCalendars: checked } })); };

  private isExchangeCalendarEnabled = (calendarId: string): boolean => {
    const states = this.state.settings.exchangeCalendarStates || {};
    return states[calendarId] !== false;
  };

  private handleToggleOutlookGroupVisibility = (): void => {
    this.setState(prev => {
      const configured = prev.settings.sources.filter(source => source.sourceType === 'exchange').map(source => source.isEnabled);
      const discovered = prev.userExchangeCalendars.map(calendar => (prev.settings.exchangeCalendarStates || {})[calendar.id] !== false);
      const target = getBulkVisibilityTarget(getGroupVisibilityState([...discovered, ...configured]));
      return { settings: setOutlookVisibility(prev.settings, prev.userExchangeCalendars.map(calendar => calendar.id), target) };
    });
  };

  private handleToggleSharePointGroupVisibility = (): void => {
    this.setState(prev => {
      const values = prev.settings.sources.filter(source => source.sourceType === 'sharepoint').map(source => source.isEnabled);
      const target = getBulkVisibilityTarget(getGroupVisibilityState(values));
      return { settings: setSharePointVisibility(prev.settings, target) };
    });
  };

  private handleDeleteSource = (id: string): void => {
    const settings = { ...this.state.settings, sources: this.state.settings.sources.filter(s => s.id !== id) };
    this.setState({ settings });
  };

  private toggleEdit = (id: string | undefined): void => { this.setState({ editingSourceId: id }); };

  private handleSave = (): void => { this.props.onSave(this.state.settings); this.props.onDismiss(); };

  private handleReset = (): void => {
    if (confirm(strings.ResetSettingsConfirmationLabel)) {
      if (this.props.onReset) this.props.onReset();
      this.props.onDismiss();
    }
  };

  private onRenderFooterContent = (): React.ReactElement => {
    return (
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton onClick={this.handleSave} text={strings.SaveLabel} />
        <DefaultButton onClick={this.props.onDismiss} text={strings.CancelLabel} />
        <DefaultButton
          onClick={this.handleReset}
          text={strings.ResetToDefaultsLabel}
          title={strings.ResetCalendarSettingsTitle}
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
          <Label>{strings.FieldMappingLabel}</Label>
          <Dropdown
            label={strings.TitleSubjectFieldLabel}
            options={spAvailableFields}
            selectedKey={spFieldMapping.titleField || ''}
            onChange={(_, option) => this.setState({
              spFieldMapping: { ...spFieldMapping, titleField: option?.key as string }
            })}
          />
          <Dropdown
            label={strings.StartDateFieldLabel}
            options={spAvailableFields}
            selectedKey={spFieldMapping.startDateField || ''}
            onChange={(_, option) => this.setState({
              spFieldMapping: { ...spFieldMapping, startDateField: option?.key as string }
            })}
          />
          <Dropdown
            label={strings.EndDateFieldLabel}
            options={spAvailableFields}
            selectedKey={spFieldMapping.endDateField || ''}
            onChange={(_, option) => this.setState({
              spFieldMapping: { ...spFieldMapping, endDateField: option?.key as string }
            })}
          />
          <Dropdown
            label={strings.LocationFieldOptionalLabel}
            options={[{ key: '', text: strings.NoneLabel }, ...spAvailableFields]}
            selectedKey={spFieldMapping.locationField || ''}
            onChange={(_, option) => this.setState({
              spFieldMapping: { ...spFieldMapping, locationField: option?.key as string }
            })}
          />
          <Dropdown
            label={strings.DescriptionFieldOptionalLabel}
            options={[{ key: '', text: strings.NoneLabel }, ...spAvailableFields]}
            selectedKey={spFieldMapping.descriptionField || ''}
            onChange={(_, option) => this.setState({
              spFieldMapping: { ...spFieldMapping, descriptionField: option?.key as string }
            })}
          />
          <TextField
            label={strings.CalendarNameLabel}
            value={this.state.newCalendarName}
            onChange={(_, value) => this.setState({ newCalendarName: value || '' })}
          />
          <div>
            <Label>{strings.ColorLabel}</Label>
            <ColorPicker
              color={this.state.newCalendarColor}
              onChange={(_, color) => this.setState({ newCalendarColor: `#${color.hex}` })}
              alphaType="none"
            />
          </div>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text={strings.AddCalendarLabel} onClick={this.handleConfirmSharePointCalendar} />
          </Stack>
        </Stack>
      );
    }

    if (spSelectedList) {
      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>{formatLocalizedString(strings.SelectedListLabel, spSelectedList.name)}</Label>
          <Label style={{ color: '#605e5c' }}>{strings.ConfiguringFieldMappingLabel}</Label>
        </Stack>
      );
    }

    if (spSelectedSite) {
      if (spListsLoading) {
        return <Spinner size={SpinnerSize.medium} label={strings.LoadingLabel} />;
      }

      if (spLists.length === 0) {
        return (
          <Stack tokens={{ childrenGap: 12 }}>
            <Label>{strings.NoCalendarListsFoundLabel}</Label>
          </Stack>
        );
      }

      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>{formatLocalizedString(strings.SelectCalendarListLabel, spSelectedSite.name)}</Label>
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
      return <Spinner size={SpinnerSize.medium} label={strings.LoadingLabel} />;
    }

    // Pagination logic
    const filteredSites = spSites;
    const totalPages = Math.ceil(filteredSites.length / this.SITES_PER_PAGE);
    const startIndex = this.state.spCurrentPage * this.SITES_PER_PAGE;
    const endIndex = startIndex + this.SITES_PER_PAGE;
    const sitesOnPage = filteredSites.slice(startIndex, endIndex);

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>{strings.SearchSelectSharePointSiteLabel}</Label>
        <TextField
          placeholder={strings.FilterSitesPlaceholder}
          value={spSiteFilter}
          onChange={(_, value) => this.handleSharePointFilterChange(value)}
        />
        <PrimaryButton text={strings.SearchLabel} onClick={this.handleSharePointSearch} />

        {spSites.length === 0 && !spSiteFilter && (
          <Label style={{ color: '#605e5c', fontStyle: 'italic' }}>
            {strings.NoSitesFoundTrySearchLabel}
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
                    title={strings.SelectThisSiteLabel}
                    ariaLabel={strings.SelectThisSiteLabel}
                    styles={{ root: { pointerEvents: 'none' } }}
                  />
                </Stack>
              ))}
            </Stack>

            {/* Pagination Controls */}
            {totalPages > 1 && (
              <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="center">
                <DefaultButton
                  text={strings.PreviousLabel}
                  disabled={this.state.spCurrentPage === 0}
                  onClick={() => this.setState({ spCurrentPage: this.state.spCurrentPage - 1 })}
                />
                <Label style={{ margin: 0 }}>
                  {formatLocalizedString(strings.PageOfLabel, this.state.spCurrentPage + 1, totalPages)}
                </Label>
                <DefaultButton
                  text={strings.NextLabel}
                  disabled={this.state.spCurrentPage >= totalPages - 1}
                  onClick={() => this.setState({ spCurrentPage: this.state.spCurrentPage + 1 })}
                />
              </Stack>
            )}

            <Label style={{ fontSize: 12, color: '#605e5c' }}>
              {formatLocalizedString(strings.ShowingSitesLabel, sitesOnPage.length, filteredSites.length)}
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
            label={strings.CalendarNameLabel}
            value={this.state.newCalendarName}
            onChange={(_, value) => this.setState({ newCalendarName: value || '' })}
          />
          <div>
            <Label>{strings.ColorLabel}</Label>
            <ColorPicker
              color={this.state.newCalendarColor}
              onChange={(_, color) => this.setState({ newCalendarColor: `#${color.hex}` })}
              alphaType="none"
            />
          </div>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text={strings.AddCalendarLabel} onClick={this.handleConfirmExchangeCalendar} />
          </Stack>
        </Stack>
      );
    }

    if (exchangeCalendarsLoading) {
      return <Spinner size={SpinnerSize.medium} label={strings.LoadingLabel} />;
    }

    if (exchangeMailboxResolved && exchangeCalendars.length > 0) {
      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>{formatLocalizedString(strings.SelectCalendarFromMailboxLabel, exchangeMailbox || strings.YourMailboxLabel)}</Label>
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
                    <span style={{ fontSize: 11, color: '#605e5c', marginLeft: 8 }}>{strings.DefaultLabel}</span>
                  )}
                </div>
              </Stack>
            ))}
          </Stack>

          {/* Option to search another mailbox */}
          <div style={{ borderTop: '1px solid #edebe9', paddingTop: 12, marginTop: 8 }}>
            <Label>{strings.EnterMailboxEmailLabel}</Label>
            <TextField
              placeholder={strings.MailboxPlaceholder}
              value={exchangeMailbox}
              onChange={(_, value) => this.handleExchangeMailboxChange(value)}
            />
            <div style={{ marginTop: 8 }}>
              <PrimaryButton text={strings.LoadOtherMailboxLabel} onClick={this.handleExchangeLookupMailbox} />
            </div>
          </div>
        </Stack>
      );
    }

    // Default: Show instructions to open Outlook
    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <Label>{strings.OutlookCalendarDescription}</Label>
        <div style={{
          padding: '16px',
          backgroundColor: '#f3f2f1',
          borderRadius: 4,
          border: '1px solid #edebe9'
        }}>
          <ol style={{ margin: 0, paddingLeft: 20 }}>
            <li style={{ marginBottom: 8 }}>
              {strings.OpenOutlookInstructionLabel}
            </li>
            <li style={{ marginBottom: 8 }}>
              {strings.OutlookDirectoryInstructionLabel}
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
                <strong>{strings.AddFromDirectoryLabel}</strong>
              </div>
            </li>
            <li>
              {strings.CalendarAppearsAutomaticallyLabel}
            </li>
          </ol>
        </div>
        <PrimaryButton
          text={strings.OpenOutlookLabel}
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
    const { icsUrl, settings } = this.state;
    const hasValidInput = icsUrl.trim() && this.state.newCalendarName.trim();

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        {settings.availableAdminIcsCatalogItems.length > 0 && (
          <Stack tokens={{ childrenGap: 8 }}>
            <Label>{strings.AvailableViaAdminLabel}</Label>
            {settings.availableAdminIcsCatalogItems.map(item => (
              <DefaultButton
                key={item.adminIcsId}
                text={item.displayName}
                onClick={() => this.setState({
                  newCalendarName: item.displayName,
                  icsUrl: item.icsUrl
                })}
                style={{ textAlign: 'left', height: 'auto', padding: '8px 12px' }}
              />
            ))}
          </Stack>
        )}
        <TextField
          label={strings.CalendarNameLabel}
          value={this.state.newCalendarName}
          onChange={(_, value) => this.setState({ newCalendarName: value || '' })}
          placeholder={strings.IcsNamePlaceholder}
        />
        <TextField
          label={strings.IcsUrlLabel}
          value={icsUrl}
          onChange={(_, value) => this.setState({ icsUrl: value || '' })}
          placeholder={strings.IcsUrlPlaceholder}
        />

        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text={strings.OpenInOutlookLabel}
            onClick={this.handleOpenOutlookCalendar}
            disabled={!hasValidInput}
            iconProps={{ iconName: 'OpenInNewWindow' }}
          />
          <DefaultButton text={strings.CancelLabel} onClick={this.handleCloseAddDialog} />
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

    const newSourceId = this.generateId();
    const newSource: ICalendarSource = {
      id: newSourceId,
      userSourceId: newSourceId,
      origin: 'user',
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

    const newSourceId = this.generateId();
    const newSource: ICalendarSource = {
      id: newSourceId,
      userSourceId: newSourceId,
      origin: 'user',
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
    const newSources = selectedGroups.map(group => {
      const newSourceId = this.generateId();
      return {
        id: newSourceId,
        userSourceId: newSourceId,
        origin: 'user' as const,
        sourceType: 'unifiedGroup' as const,
        name: group.displayName,
        color: this.state.newCalendarColor,
        isEnabled: true,
        groupId: group.id,
        showSourceLogo: true
      };
    });

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
          <Label>{strings.LoadingGroupsAndTeamsLabel}</Label>
          <Spinner size={SpinnerSize.large} label={strings.LoadingLabel} />
        </Stack>
      );
    }

    if (unifiedGroups.length === 0) {
      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>{strings.NoGroupsFoundLabel}</Label>
          <div>{strings.NoGroupsAccessLabel}</div>
          <DefaultButton text={strings.CancelLabel} onClick={this.handleCloseAddDialog} />
        </Stack>
      );
    }

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>{strings.SelectGroupsOrTeamsLabel}</Label>
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
          <Label>{strings.ColorLabel}</Label>
          <ColorPicker
            color={this.state.newCalendarColor}
            onChange={(_, color) => this.setState({ newCalendarColor: `#${color.hex}` })}
            alphaType="none"
          />
        </div>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text={selectedCount > 1 ? strings.AddCalendarsLabel : strings.AddCalendarLabel}
            onClick={this.handleConfirmUnifiedGroups}
            disabled={selectedCount === 0}
          />
          <DefaultButton text={strings.CancelLabel} onClick={this.handleCloseAddDialog} />
        </Stack>
      </Stack>
    );
  };

  private renderTeamsShiftsFlow = (): React.ReactElement => {
    const hasValidInput = this.state.newCalendarName.trim();

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>{strings.TeamsShiftsLabel}</Label>
        <TextField
          label={strings.CalendarNameLabel}
          value={this.state.newCalendarName}
          onChange={(_, value) => this.setState({ newCalendarName: value || '' })}
          placeholder={strings.TeamsShiftsLabel}
          required
        />

        <Toggle
          label={strings.SourceLogoLabel}
          checked={this.state.teamsShiftsShowLogo}
          onChange={(_, checked) => this.setState({ teamsShiftsShowLogo: checked || false })}
          onText={strings.OnLabel}
          offText={strings.OffLabel}
        />

        <ColorPicker
          color={this.state.newCalendarColor}
          onChange={(_, color) => this.setState({ newCalendarColor: `#${color.hex}` })}
          alphaType="none"
          showPreview={true}
        />

        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text={strings.AddCalendarLabel}
            onClick={this.handleConfirmTeamsShifts}
            disabled={!hasValidInput}
          />
          <DefaultButton text={strings.CancelLabel} onClick={this.handleCloseAddDialog} />
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
            <Label>{strings.LoadingPlannerPlansLabel}</Label>
            <Spinner size={SpinnerSize.large} label={strings.LoadingLabel} />
          </Stack>
        );
      }

      if (plannerPlans.length === 0) {
        return (
          <Stack tokens={{ childrenGap: 12 }}>
            <Label>{strings.NoPlannerPlansFoundLabel}</Label>
            <div>{strings.NoPlannerPlansAccessLabel}</div>
            <DefaultButton text={strings.CancelLabel} onClick={this.handleCloseAddDialog} />
          </Stack>
        );
      }

      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>{strings.SelectPlannerPlanLabel}</Label>
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
          <Label>{strings.ConfigurePlannerIntegrationLabel}</Label>

          {selectedPlan && (
            <div style={{ padding: '8px', backgroundColor: '#f3f2f1', borderRadius: '4px' }}>
              <Icon iconName="PlannerLogo" style={{ marginRight: 8 }} />
              <strong>{selectedPlan.title}</strong>
            </div>
          )}

          <TextField
            label={strings.CalendarNameLabel}
            value={this.state.newCalendarName}
            onChange={(_, value) => this.setState({ newCalendarName: value || '' })}
            placeholder={strings.ProjectTasksPlaceholder}
            required
          />

          <Toggle
            label={strings.AssignedToMeOnlyLabel}
            checked={this.state.plannerAssignedToMeOnly}
            onChange={(_, checked) => this.setState({ plannerAssignedToMeOnly: checked || false })}
            onText={strings.OnLabel}
            offText={strings.OffLabel}
          />

          <Toggle
            label={strings.ShowCompletedTasksLabel}
            checked={this.state.plannerShowCompleted}
            onChange={(_, checked) => this.setState({ plannerShowCompleted: checked || false })}
            onText={strings.OnLabel}
            offText={strings.OffLabel}
          />

          <Toggle
            label={strings.SourceLogoLabel}
            checked={this.state.plannerShowLogo}
            onChange={(_, checked) => this.setState({ plannerShowLogo: checked || false })}
            onText={strings.OnLabel}
            offText={strings.OffLabel}
          />

          <ColorPicker
            color={this.state.newCalendarColor}
            onChange={(_, color) => this.setState({ newCalendarColor: `#${color.hex}` })}
            alphaType="none"
            showPreview={true}
          />

          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text={strings.AddCalendarLabel}
              onClick={this.handleConfirmPlannerPlan}
              disabled={!hasValidInput}
            />
            <DefaultButton text={strings.CancelLabel} onClick={this.handleCloseAddDialog} />
          </Stack>
        </Stack>
      );
    }

    return <div>{strings.UnknownStepLabel}</div>;
  };

  private renderAddCalendarFlow = (): React.ReactElement => {
    const { addingCalendarStep, addingCalendarType } = this.state;

    // Show type selection first
    if (addingCalendarStep === 'initial') {
      return (
        <Stack tokens={{ childrenGap: 16 }}>
      <Label>{strings.SelectCalendarTypeLabel}</Label>
          <Stack tokens={{ childrenGap: 12 }}>
            <PrimaryButton
              text={strings.AddCalendarSharePointLabel}
              secondaryText={strings.AddCalendarSharePointDescription}
              iconProps={{ iconName: 'SharepointLogo' }}
              onClick={() => this.handleSelectAddType('sharepoint')}
              style={{ textAlign: 'left', height: 'auto', padding: '12px' }}
            />
            <PrimaryButton
              text={strings.AddCalendarExchangeLabel}
              secondaryText={strings.OpenOutlookLabel}
              iconProps={{ iconName: 'OutlookLogo' }}
              onClick={() => this.handleSelectAddType('exchange')}
              style={{ textAlign: 'left', height: 'auto', padding: '12px' }}
            />
            <PrimaryButton
              text={strings.MicrosoftPlannerLabel}
              secondaryText={strings.AddFromPlannerLabel}
              iconProps={{ iconName: 'PlannerLogo' }}
              onClick={() => this.handleSelectAddType('planner')}
              style={{ textAlign: 'left', height: 'auto', padding: '12px' }}
            />
            <PrimaryButton
              text={strings.Microsoft365GroupLabel}
              secondaryText={strings.AddGroupCalendarDescription}
              iconProps={{ iconName: 'Group' }}
              onClick={() => this.handleSelectAddType('unifiedGroup')}
              style={{ textAlign: 'left', height: 'auto', padding: '12px' }}
            />
            <PrimaryButton
              text={strings.TeamsLabel}
              secondaryText={strings.AddTeamCalendarDescription}
              iconProps={{ iconName: 'TeamsLogo' }}
              onClick={() => this.handleSelectAddType('unifiedGroup')}
              style={{ textAlign: 'left', height: 'auto', padding: '12px' }}
            />
            <PrimaryButton
              text={strings.TeamsShiftsLabel}
              secondaryText={strings.TeamsShiftsDescription}
              iconProps={{ iconName: 'Clock' }}
              onClick={() => this.handleSelectAddType('teamsShifts')}
              style={{ textAlign: 'left', height: 'auto', padding: '12px' }}
            />
            <PrimaryButton
              text={strings.InternetCalendarLabel}
              secondaryText={strings.AddIcsCalendarDescription}
              iconProps={{ iconName: 'World' }}
              onClick={() => this.handleSelectAddType('ics')}
              style={{ textAlign: 'left', height: 'auto', padding: '12px' }}
            />
          </Stack>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton text={strings.CancelLabel} onClick={this.handleCloseAddDialog} />
          </Stack>
        </Stack>
      );
    }

    // Show flow for selected type
    if (!addingCalendarType) {
      return <div>{strings.UnknownCalendarTypeLabel}</div>;
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
    visibilityState?: GroupVisibilityState;
    visibilityDisabled?: boolean;
    onVisibilityChange?: () => void;
    headerActions?: React.ReactNode;
    children: React.ReactNode;
  }): React.ReactElement => {
    const { sectionKey, icon, iconBg, iconColor, title, subtitle, showLogoValue, onShowLogoChange, visibilityState, visibilityDisabled, onVisibilityChange, headerActions, children } = params;
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
          {(headerActions !== undefined || showLogoValue !== undefined || visibilityState !== undefined) && (
            <div
              style={{ display: 'flex', alignItems: 'center', gap: 8 }}
              onClick={e => e.stopPropagation()}
            >
              {headerActions}
              {visibilityState !== undefined && onVisibilityChange && (
                <Checkbox
                  checked={visibilityState === 'on'}
                  indeterminate={visibilityState === 'mixed'}
                  disabled={visibilityDisabled}
                  ariaLabel={visibilityState === 'on' ? strings.HideAllCalendarsLabel : strings.ShowAllCalendarsLabel}
                  title={visibilityState === 'on' ? strings.HideAllCalendarsLabel : strings.ShowAllCalendarsLabel}
                  onChange={onVisibilityChange}
                  styles={{ root: { margin: 0 } }}
                />
              )}
              {showLogoValue !== undefined && onShowLogoChange && (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                  <span style={{ fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px', opacity: 0.6 }}>{strings.ShowLogosLabel}</span>
                  <Toggle
                    checked={showLogoValue}
                    onChange={(_, checked) => onShowLogoChange(!!checked)}
                    onText={strings.OnLabel}
                    offText={strings.OffLabel}
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
              <span style={{ fontSize: 11, color: '#605e5c', marginLeft: 8 }}>{strings.DefaultLabel}</span>
            )}
          </div>
          <IconButton
            iconProps={{ iconName: 'Edit' }}
            title={strings.EditLabel}
            ariaLabel={strings.EditLabel}
            onClick={() => window.open('https://outlook.cloud.microsoft/calendar/', '_blank', 'noopener,noreferrer')}
            styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 14 } }}
          />
          <IconButton
            iconProps={{ iconName: isEnabled ? 'View' : 'Hide' }}
            title={isEnabled ? strings.HideCalendarLabel : strings.ShowCalendarLabel}
            ariaLabel={isEnabled ? strings.HideCalendarLabel : strings.ShowCalendarLabel}
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
    const isAdminSource = source.origin === 'admin';

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
              label={strings.NameLabel}
              value={source.name}
              onChange={(_, value) => this.handleUpdateSource(source.id, { name: value || '' })}
            />
            {source.sourceType === 'sharepoint' && (
              <div style={{ fontSize: 12, color: '#605e5c' }}>
                {strings.SiteLabel}: {source.sharePointSiteName || strings.SiteNameUnavailableLabel}
              </div>
            )}
            <div style={{ order: 3 }}>
              <Label>{strings.ColorLabel}</Label>
              <ColorPicker
                color={source.color}
                onChange={(_, color) => this.handleUpdateSource(source.id, { color: `#${color.hex}` })}
                alphaType="none"
              />
            </div>
            <Toggle
              label={strings.EnabledLabel}
              checked={source.isEnabled}
              onChange={(_, checked) => this.handleUpdateSource(source.id, { isEnabled: !!checked })}
              onText={strings.OnLabel}
              offText={strings.OffLabel}
            />
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton text={strings.DoneLabel} onClick={() => this.toggleEdit(undefined)} />
              <DefaultButton text={isAdminSource ? strings.RemoveForMeLabel : strings.DeleteLabel} onClick={() => this.handleDeleteSource(source.id)} />
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
              {source.sourceType === 'sharepoint' && (
                <div style={{ fontSize: 11, color: '#605e5c' }}>
                  {strings.SiteLabel}: {source.sharePointSiteName || strings.SiteNameUnavailableLabel}
                </div>
              )}
              {isAdminSource && (
                <span style={{ fontSize: 11, color: '#605e5c', marginLeft: 8 }}>({strings.AdminDefaultLabel})</span>
              )}
              {isAdminSource && source.audienceGroupNames && source.audienceGroupNames.length > 0 && (
                <div style={{ fontSize: 11, color: '#605e5c' }}>
                  {strings.ViaLabel}: {source.audienceGroupNames.join(', ')}
                </div>
              )}
            </div>
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              title={strings.EditLabel}
              ariaLabel={strings.EditLabel}
              onClick={() => this.toggleEdit(source.id)}
              styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 14 } }}
            />
            <IconButton
              iconProps={{ iconName: source.isEnabled ? 'View' : 'Hide' }}
              title={source.isEnabled ? strings.HideCalendarLabel : strings.ShowCalendarLabel}
              ariaLabel={source.isEnabled ? strings.HideCalendarLabel : strings.ShowCalendarLabel}
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
    const effectiveVisibleHours = settings.userVisibleHourCount ?? settings.visibleHourCount;
    const effectiveStartMinutes = settings.userPreferredStartMinutes ?? settings.preferredStartMinutes;
    const startOptions: IDropdownOption[] = [];
    const latestStart = Math.max(0, 24 * 60 - effectiveVisibleHours * 60);
    for (let minutes = 0; minutes <= latestStart; minutes += settings.slotDurationMinutes) {
      startOptions.push({ key: minutes, text: formatCalendarTime(new Date(2000, 0, 1, 0, minutes), this.props.locale) });
    }
    const visibleHourOptions = Array.from({ length: 24 }, (_, index) => ({ key: index + 1, text: String(index + 1) }));

    return (
      <Panel
        isOpen={isOpen}
        onDismiss={onDismiss}
        type={PanelType.medium}
        headerText={showAddDialog ? strings.AddCalendarLabel : strings.CalendarSettingsTitle}
        onRenderFooterContent={!showAddDialog ? this.onRenderFooterContent : undefined}
        isFooterAtBottom={true}
      >
        {showAddDialog ? (
          this.renderAddCalendarFlow()
        ) : (
          <Stack tokens={{ childrenGap: 16 }}>
            <div style={{ order: 3 }}>
              <Label>{strings.TimelinePreferencesLabel}</Label>
              <Toggle
                label={strings.UsePersonalStartTimeLabel}
                checked={settings.userPreferredStartMinutes !== undefined}
                onChange={(_, checked) => this.setState({ settings: { ...settings, userPreferredStartMinutes: checked ? settings.preferredStartMinutes : undefined } })}
                onText={strings.OnLabel}
                offText={strings.OffLabel}
              />
              <Dropdown
                label={strings.PreferredStartTimeLabel}
                disabled={settings.userPreferredStartMinutes === undefined}
                selectedKey={effectiveStartMinutes}
                options={startOptions}
                onChange={(_, option) => this.setState({ settings: { ...settings, userPreferredStartMinutes: Number(option?.key) } })}
              />
              <Toggle
                label={strings.UsePersonalVisibleHoursLabel}
                checked={settings.userVisibleHourCount !== undefined}
                onChange={(_, checked) => this.setState({ settings: { ...settings, userVisibleHourCount: checked ? settings.visibleHourCount : undefined } })}
                onText={strings.OnLabel}
                offText={strings.OffLabel}
              />
              <Dropdown
                label={strings.VisibleHoursLabel}
                disabled={settings.userVisibleHourCount === undefined}
                selectedKey={effectiveVisibleHours}
                options={visibleHourOptions}
                onChange={(_, option) => {
                  const userVisibleHourCount = Number(option?.key);
                  const userPreferredStartMinutes = Math.min(effectiveStartMinutes, 24 * 60 - userVisibleHourCount * 60);
                  this.setState({ settings: { ...settings, userVisibleHourCount, userPreferredStartMinutes: settings.userPreferredStartMinutes === undefined ? undefined : userPreferredStartMinutes } });
                }}
              />
              <Toggle
                label={strings.UsePersonalWeekendPreferenceLabel}
                checked={settings.userShowWeekends !== undefined}
                onChange={(_, checked) => this.setState({ settings: {
                  ...settings,
                  showWeekends: checked ? settings.showWeekends : settings.adminShowWeekends,
                  userShowWeekends: checked ? settings.showWeekends : undefined
                } })}
                onText={strings.OnLabel}
                offText={strings.OffLabel}
              />
              <Toggle
                label={strings.ShowWeekendsLabel}
                disabled={settings.userShowWeekends === undefined}
                checked={settings.showWeekends}
                onChange={(_, checked) => this.setState({ settings: { ...settings, showWeekends: !!checked, userShowWeekends: !!checked } })}
                onText={strings.OnLabel}
                offText={strings.OffLabel}
              />
              <div style={{ marginTop: 6, fontSize: 12, color: '#605e5c' }}>
                {formatLocalizedString(strings.GridSummaryLabel, settings.slotDurationMinutes, settings.showWeekends ? strings.WeekendsShownLabel : strings.WeekendsHiddenLabel)}
              </div>
            </div>

            {/* Add Calendar Button at Top */}
            <div style={{ order: 1 }}>
              <PrimaryButton
                text={strings.AddCalendarLabel}
                iconProps={{ iconName: 'Add' }}
                onClick={this.handleOpenAddDialog}
              />
            </div>

            {/* Calendar Sources Section */}
            <div style={{ order: 2 }}>
              <Stack tokens={{ childrenGap: 12 }}>

                {/* Outlook */}
                {(userExchangeCalendarsLoading || userExchangeCalendars.length > 0 || settings.sources.some(source => source.sourceType === 'exchange')) && this.renderSourceSection({
                  sectionKey: 'outlook',
                  icon: 'OutlookLogo',
                  iconBg: 'rgba(0, 120, 212, 0.12)',
                  iconColor: '#0078d4',
                  title: strings.OutlookLabel,
                  subtitle: strings.OutlookSectionSubtitle,
                  showLogoValue: settings.exchangeShowSourceLogo ?? true,
                  onShowLogoChange: (checked) => this.setState({ settings: { ...settings, exchangeShowSourceLogo: checked } }),
                  visibilityState: getGroupVisibilityState([
                    ...userExchangeCalendars.map(calendar => this.isExchangeCalendarEnabled(calendar.id)),
                    ...settings.sources.filter(source => source.sourceType === 'exchange').map(source => source.isEnabled)
                  ]),
                  visibilityDisabled: userExchangeCalendarsLoading || userExchangeCalendars.length + settings.sources.filter(source => source.sourceType === 'exchange').length === 0,
                  onVisibilityChange: this.handleToggleOutlookGroupVisibility,
                  headerActions: (
                    <DefaultButton
                      text={strings.ManageLabel}
                      iconProps={{ iconName: 'OpenInNewWindow' }}
                      onClick={() => window.open('https://outlook.cloud.microsoft/calendar/', '_blank', 'noopener,noreferrer')}
                      styles={{ root: { height: 28, minWidth: 0, padding: '0 8px', fontSize: 12 } }}
                    />
                  ),
                  children: (
                    <Stack tokens={{ childrenGap: 6 }}>
                      {userExchangeCalendarsLoading
                        ? <Spinner size={SpinnerSize.small} label={strings.LoadingLabel} />
                        : userExchangeCalendars.map((calendar, i) => this.renderExchangeCalendarItem(calendar, i))}
                      {settings.sources.filter(source => source.sourceType === 'exchange').map((source, i) => this.renderCalendarSource(source, userExchangeCalendars.length + i))}
                    </Stack>
                  )
                })}

                {/* SharePoint */}
                {settings.sources.filter(s => s.sourceType === 'sharepoint').length > 0 && this.renderSourceSection({
                  sectionKey: 'sharepoint',
                  icon: 'SharepointLogo',
                  iconBg: 'rgba(3, 129, 134, 0.12)',
                  iconColor: '#038186',
                  title: strings.SharePointLabel,
                  subtitle: strings.SharePointSectionSubtitle,
                  showLogoValue: settings.sharePointShowSourceLogo ?? true,
                  onShowLogoChange: (checked) => this.setState({ settings: { ...settings, sharePointShowSourceLogo: checked } }),
                  visibilityState: getGroupVisibilityState(settings.sources.filter(source => source.sourceType === 'sharepoint').map(source => source.isEnabled)),
                  onVisibilityChange: this.handleToggleSharePointGroupVisibility,
                  children: settings.sources.filter(s => s.sourceType === 'sharepoint').map((source, i) => this.renderCalendarSource(source, i))
                })}

                {/* Planner */}
                {this.renderSourceSection({
                  sectionKey: 'planner',
                  icon: 'PlannerLogo',
                  iconBg: 'rgba(16, 124, 65, 0.12)',
                  iconColor: '#107c41',
                  title: strings.PlannerLabel,
                  subtitle: strings.PlannerSectionSubtitle,
                  showLogoValue: settings.plannerShowSourceLogo ?? true,
                  onShowLogoChange: (checked) => this.setState({ settings: { ...settings, plannerShowSourceLogo: checked } }),
                  children: (() => {
                    const plannerSources = settings.sources.filter(s => s.sourceType === 'planner');
                    const showAllPlanner = settings.plannerShowAllCalendars ?? false;

                    return (
                      <Stack tokens={{ childrenGap: 8 }}>
                        <Toggle
                          label={strings.ShowAllPlannerPlansLabel}
                          checked={showAllPlanner}
                          onChange={(_, checked) => this.handleTogglePlannerShowAll(!!checked)}
                          onText={strings.OnLabel}
                          offText={strings.OffLabel}
                        />
                        {showAllPlanner && (
                          <Toggle
                            label={strings.AssignedToMeOnlyLabel}
                            checked={settings.plannerShowAllAssignedToMeOnly ?? false}
                            onChange={(_, checked) => this.handleTogglePlannerAssignedToMeOnly(!!checked)}
                            onText={strings.OnLabel}
                            offText={strings.OffLabel}
                          />
                        )}
                        {showAllPlanner && (
                          <MessageBar messageBarType={MessageBarType.info}>
                            {strings.AllPlannerPlansLoadedLabel}
                          </MessageBar>
                        )}
                        {plannerSources.length > 0 && (
                          <div style={showAllPlanner ? { opacity: 0.4, pointerEvents: 'none' } : undefined}>
                            {plannerSources.map((source, i) => this.renderCalendarSource(source, i))}
                          </div>
                        )}
                      </Stack>
                    );
                  })()
                })}

                {/* Groups & Teams */}
                {this.renderSourceSection({
                  sectionKey: 'unifiedGroup',
                  icon: 'Group',
                  iconBg: 'rgba(91, 95, 199, 0.12)',
                  iconColor: '#5b5fc7',
                  title: strings.GroupsAndTeamsLabel,
                  subtitle: strings.GroupsAndTeamsSectionSubtitle,
                  showLogoValue: settings.unifiedGroupShowSourceLogo ?? true,
                  onShowLogoChange: (checked) => this.setState({ settings: { ...settings, unifiedGroupShowSourceLogo: checked } }),
                  children: (() => {
                    const unifiedGroupSources = settings.sources.filter(s => s.sourceType === 'unifiedGroup');
                    const showAllUnifiedGroups = settings.unifiedGroupShowAllCalendars ?? false;

                    return (
                      <Stack tokens={{ childrenGap: 8 }}>
                        <Toggle
                          label={strings.ShowAllGroupsAndTeamsLabel}
                          checked={showAllUnifiedGroups}
                          onChange={(_, checked) => this.handleToggleUnifiedGroupShowAll(!!checked)}
                          onText={strings.OnLabel}
                          offText={strings.OffLabel}
                        />
                        {showAllUnifiedGroups && (
                          <MessageBar messageBarType={MessageBarType.info}>
                            {strings.AllGroupTeamCalendarsLoadedLabel}
                          </MessageBar>
                        )}
                        {unifiedGroupSources.length > 0 && (
                          <div style={showAllUnifiedGroups ? { opacity: 0.4, pointerEvents: 'none' } : undefined}>
                            {unifiedGroupSources.map((source, i) => this.renderCalendarSource(source, i))}
                          </div>
                        )}
                      </Stack>
                    );
                  })()
                })}

                {/* Teams Shifts */}
                {this.renderSourceSection({
                  sectionKey: 'teamsShifts',
                  icon: 'Clock',
                  iconBg: 'rgba(74, 79, 190, 0.12)',
                  iconColor: '#4a4fbe',
                  title: strings.TeamsShiftsLabel,
                  subtitle: strings.TeamsShiftsSectionSubtitle,
                  showLogoValue: settings.teamsShiftsShowSourceLogo ?? true,
                  onShowLogoChange: (checked) => this.setState({ settings: { ...settings, teamsShiftsShowSourceLogo: checked } }),
                  children: (() => {
                    const teamsShiftsSources = settings.sources.filter(s => s.sourceType === 'teamsShifts');
                    const showAllTeamsShifts = settings.teamsShiftsShowAllCalendars ?? false;

                    return (
                      <Stack tokens={{ childrenGap: 8 }}>
                        <Toggle
                          label={strings.ShowAllTeamsShiftsLabel}
                          checked={showAllTeamsShifts}
                          onChange={(_, checked) => this.handleToggleTeamsShiftsShowAll(!!checked)}
                          onText={strings.OnLabel}
                          offText={strings.OffLabel}
                        />
                        {showAllTeamsShifts && (
                          <MessageBar messageBarType={MessageBarType.info}>
                            {strings.AllTeamsShiftsLoadedLabel}
                          </MessageBar>
                        )}
                        {teamsShiftsSources.length > 0 && (
                          <div style={showAllTeamsShifts ? { opacity: 0.4, pointerEvents: 'none' } : undefined}>
                            {teamsShiftsSources.map((source, i) => this.renderCalendarSource(source, i))}
                          </div>
                        )}
                      </Stack>
                    );
                  })()
                })}

                {/* Empty state */}
                {!userExchangeCalendarsLoading && userExchangeCalendars.length === 0 && settings.sources.length === 0 && (
                  <Label style={{ color: '#605e5c', fontStyle: 'italic' }}>
                    {strings.NoCalendarsFoundLabel}
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
