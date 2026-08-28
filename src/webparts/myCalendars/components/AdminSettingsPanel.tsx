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
import {
  type IAdminAssignedSource,
  type IAdminIcsCatalogItem,
  type IAdminWebPartSettings,
  type IAudienceGroup,
  type ICalendarSourceBase,
  type CalendarSourceType,
  defaultAdminWebPartSettings
} from '../models/ICalendarSettings';
import { calendarSourceRegistry } from '../models/CalendarSourceRegistry';
import * as strings from 'MyCalendarsWebPartStrings';
import { ExchangeCalendarService, IExchangeCalendar } from '../services/ExchangeCalendarService';
import { SharePointCalendarService, ISharePointList, ISharePointSite } from '../services/SharePointCalendarService';
import { PlannerTaskService, IPlannerPlan } from '../services/PlannerTaskService';
import { UnifiedGroupCalendarService, IUnifiedGroupItem } from '../services/UnifiedGroupCalendarService';
import { AudienceService, IEntraSecurityGroup } from '../services/AudienceService';
import { createAdminAssignedSource, generateStableId } from '../services/CalendarSettingsService';
import { getSourceTypeDisplayName } from '../utils/sourceIconHelper';

type AdminAddStep =
  | 'initial'
  | 'sharepoint-site'
  | 'sharepoint-list'
  | 'sharepoint-fields'
  | 'exchange-calendar'
  | 'exchange-mailbox'
  | 'ics'
  | 'planner-plan'
  | 'planner-options'
  | 'teams-shifts'
  | 'unified-group-select'
  | 'admin-audience-select';

interface IGraphColumn {
  name?: string;
  displayName?: string;
  columnGroup?: string;
}

export interface IAdminSettingsPanelProps {
  isOpen: boolean;
  onDismiss: () => void;
  settings: IAdminWebPartSettings;
  onSave: (settings: IAdminWebPartSettings) => Promise<void> | void;
  loadNotice?: string;
  httpClient?: HttpClient;
  graphClient?: MSGraphClientV3;
}

interface IAdminSettingsPanelState {
  settings: IAdminWebPartSettings;
  editingSourceId: string | undefined;
  editingIcsId: string | undefined;
  showAddDialog: boolean;
  addingCalendarType: CalendarSourceType | undefined;
  addingCalendarStep: AdminAddStep;
  spSites: ISharePointSite[];
  spSitesLoading: boolean;
  spSiteFilter: string;
  spCurrentPage: number;
  spSelectedSite: ISharePointSite | undefined;
  spLists: ISharePointList[];
  spListsLoading: boolean;
  spSelectedList: ISharePointList | undefined;
  exchangeCalendars: IExchangeCalendar[];
  exchangeCalendarsLoading: boolean;
  exchangeMailbox: string;
  exchangeMailboxResolved: boolean;
  exchangeSelectedCalendarId: string | undefined;
  spAvailableFields: IDropdownOption[];
  spFieldMapping: {
    titleField?: string;
    startDateField?: string;
    endDateField?: string;
    descriptionField?: string;
    locationField?: string;
    allDayField?: string;
  };
  icsUrl: string;
  plannerPlans: IPlannerPlan[];
  plannerPlansLoading: boolean;
  plannerSelectedPlanId: string | undefined;
  plannerAssignedToMeOnly: boolean;
  plannerShowCompleted: boolean;
  plannerShowLogo: boolean;
  teamsShiftsShowLogo: boolean;
  unifiedGroups: IUnifiedGroupItem[];
  unifiedGroupsLoading: boolean;
  unifiedGroupsSelection: Record<string, boolean>;
  newCalendarColor: string;
  newCalendarName: string;
  securityGroups: IEntraSecurityGroup[];
  securityGroupsLoading: boolean;
  securityGroupSearch: string;
  selectedAudienceGroups: Record<string, string>;
  pendingAdminSources: ICalendarSourceBase[];
  pendingAdminIcs: { adminIcsId?: string; displayName: string; icsUrl: string } | undefined;
  audienceEditTarget: { kind: 'source'; id: string } | { kind: 'ics'; id: string } | undefined;
}

export class AdminSettingsPanel extends React.Component<IAdminSettingsPanelProps, IAdminSettingsPanelState> {
  private exchangeService: ExchangeCalendarService | null = null;
  private sharePointService: SharePointCalendarService | null = null;
  private plannerService: PlannerTaskService | null = null;
  private unifiedGroupService: UnifiedGroupCalendarService | null = null;
  private audienceService: AudienceService | null = null;
  private readonly SITES_PER_PAGE = 20;

  constructor(props: IAdminSettingsPanelProps) {
    super(props);

    if (props.httpClient) {
      this.exchangeService = new ExchangeCalendarService(props.httpClient, props.graphClient);
      this.sharePointService = new SharePointCalendarService(props.httpClient, props.graphClient);
      this.plannerService = new PlannerTaskService(props.httpClient, props.graphClient);
      this.unifiedGroupService = new UnifiedGroupCalendarService(props.httpClient, props.graphClient);
      if (props.graphClient) {
        this.exchangeService.setGraphClient(props.graphClient);
        this.sharePointService.setGraphClient(props.graphClient);
        this.plannerService.setGraphClient(props.graphClient);
        this.unifiedGroupService.setGraphClient(props.graphClient);
        this.audienceService = new AudienceService(props.graphClient);
      }
    }

    this.state = this.createStateFromProps(props);
  }

  public componentDidMount(): void {
    if (this.props.graphClient) {
      this.initializeGraphClient(this.props.graphClient);
    }
  }

  public componentDidUpdate(prevProps: IAdminSettingsPanelProps): void {
    if (this.props.graphClient && !prevProps.graphClient) {
      this.initializeGraphClient(this.props.graphClient);
    }

    if (prevProps.isOpen !== this.props.isOpen && this.props.isOpen) {
      this.setState(this.createStateFromProps(this.props));
    }
  }

  private createStateFromProps(props: IAdminSettingsPanelProps): IAdminSettingsPanelState {
    return {
      settings: JSON.parse(JSON.stringify(props.settings)),
      editingSourceId: undefined,
      editingIcsId: undefined,
      showAddDialog: false,
      addingCalendarType: undefined,
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
      spAvailableFields: [],
      spFieldMapping: {},
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
      newCalendarColor: props.settings.organizationPrimaryColor || '#0078d4',
      newCalendarName: '',
      securityGroups: [],
      securityGroupsLoading: false,
      securityGroupSearch: '',
      selectedAudienceGroups: {},
      pendingAdminSources: [],
      pendingAdminIcs: undefined,
      audienceEditTarget: undefined
    };
  }

  private initializeGraphClient(client: MSGraphClientV3): void {
    if (this.exchangeService) this.exchangeService.setGraphClient(client);
    if (this.sharePointService) this.sharePointService.setGraphClient(client);
    if (this.plannerService) this.plannerService.setGraphClient(client);
    if (this.unifiedGroupService) this.unifiedGroupService.setGraphClient(client);
    this.audienceService = new AudienceService(client);
  }

  private handleOpenAddDialog = (): void => {
    this.setState({
      showAddDialog: true,
      addingCalendarType: undefined,
      addingCalendarStep: 'initial'
    });
  };

  private handleCloseAddDialog = (): void => {
    this.setState({
      showAddDialog: false,
      addingCalendarType: undefined,
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
      spAvailableFields: [],
      spFieldMapping: {},
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
      newCalendarColor: this.state.settings.organizationPrimaryColor || '#0078d4',
      newCalendarName: '',
      securityGroups: [],
      securityGroupsLoading: false,
      securityGroupSearch: '',
      selectedAudienceGroups: {},
      pendingAdminSources: [],
      pendingAdminIcs: undefined,
      audienceEditTarget: undefined
    });
  };

  private handleResetDraft = (): void => {
    if (confirm('Reset the admin draft to hardcoded defaults? This only changes the draft until you click Save.')) {
      this.setState(this.createStateFromProps({
        ...this.props,
        settings: JSON.parse(JSON.stringify(defaultAdminWebPartSettings)) as IAdminWebPartSettings
      }));
    }
  };

  private handleSave = (): void => {
    Promise.resolve(this.props.onSave(this.state.settings)).catch(error => {
      console.error('Failed to save admin settings draft:', error);
    });
  };

  private handleSelectAddType = async (type: CalendarSourceType): Promise<void> => {
    if (type === 'sharepoint') {
      this.setState({ addingCalendarType: type, addingCalendarStep: 'sharepoint-site', spSitesLoading: true });
      const sites = await this.sharePointService?.getAccessibleSites() || [];
      this.setState({ spSites: sites, spSitesLoading: false });
      return;
    }

    if (type === 'exchange') {
      this.setState({ addingCalendarType: type, addingCalendarStep: 'exchange-mailbox' });
      return;
    }

    if (type === 'ics') {
      this.setState({ addingCalendarType: type, addingCalendarStep: 'ics' });
      return;
    }

    if (type === 'planner') {
      this.setState({ addingCalendarType: type, addingCalendarStep: 'planner-plan', plannerPlansLoading: true });
      const plans = await this.plannerService?.getUserPlans() || [];
      this.setState({ plannerPlans: plans, plannerPlansLoading: false });
      return;
    }

    if (type === 'unifiedGroup') {
      this.setState({
        addingCalendarType: type,
        addingCalendarStep: 'unified-group-select',
        unifiedGroupsLoading: true,
        unifiedGroupsSelection: {},
        newCalendarColor: this.state.settings.organizationPrimaryColor || '#0078d4'
      });
      await this.loadUnifiedGroups();
      return;
    }

    if (type === 'teamsShifts') {
      this.setState({
        addingCalendarType: type,
        addingCalendarStep: 'teams-shifts',
        newCalendarName: 'Teams Shifts',
        newCalendarColor: this.state.settings.organizationPrimaryColor || '#4a4fbe',
        teamsShiftsShowLogo: true
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
      securityGroups: [],
      securityGroupsLoading: false,
      securityGroupSearch: '',
      selectedAudienceGroups: {},
      pendingAdminSources: [],
      pendingAdminIcs: undefined,
      audienceEditTarget: undefined,
      newCalendarColor: this.state.settings.organizationPrimaryColor || '#0078d4',
      newCalendarName: ''
    });
  };

  private handleBackOneStep = (): void => {
    if (this.state.addingCalendarStep === 'admin-audience-select') {
      if (this.state.audienceEditTarget) {
        this.setState({ showAddDialog: false, audienceEditTarget: undefined, selectedAudienceGroups: {} });
      } else {
        this.setState({
          addingCalendarStep: this.state.pendingAdminIcs ? 'ics' : 'initial',
          securityGroups: [],
          securityGroupsLoading: false,
          securityGroupSearch: '',
          selectedAudienceGroups: {}
        });
      }
      return;
    }

    if (this.state.addingCalendarType === 'sharepoint') {
      if (this.state.addingCalendarStep === 'sharepoint-fields') {
        this.setState({ spSelectedList: undefined, addingCalendarStep: 'sharepoint-list' });
        return;
      }
      if (this.state.addingCalendarStep === 'sharepoint-list') {
        this.setState({ spSelectedSite: undefined, spLists: [], addingCalendarStep: 'sharepoint-site' });
        return;
      }
    }

    if (this.state.addingCalendarType === 'planner' && this.state.addingCalendarStep === 'planner-options') {
      this.setState({ addingCalendarStep: 'planner-plan' });
      return;
    }

    if (this.state.addingCalendarType === 'exchange' && this.state.exchangeSelectedCalendarId) {
      this.setState({ exchangeSelectedCalendarId: undefined });
      return;
    }

    this.handleBackToTypeSelection();
  };

  private renderNavigationHeader = (): React.ReactElement | null => {
    if (this.state.addingCalendarStep === 'initial') {
      return null;
    }

    return (
      <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 16 }}>
        <IconButton iconProps={{ iconName: 'Back' }} title="Back" ariaLabel="Back" onClick={this.handleBackOneStep} styles={{ root: { height: 32 } }} />
        <IconButton iconProps={{ iconName: 'Home' }} title="Home" ariaLabel="Home" onClick={this.handleBackToTypeSelection} styles={{ root: { height: 32 } }} />
      </Stack>
    );
  };

  private async loadUnifiedGroups(): Promise<void> {
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
  }

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
    this.setState({
      spSelectedList: list,
      newCalendarName: list.name,
      newCalendarColor: this.state.settings.organizationPrimaryColor || '#0078d4'
    }, () => {
      this.fetchSharePointListFields(list).catch(err => console.error('Failed to fetch SharePoint list fields:', err));
    });
  };

  private async fetchSharePointListFields(list: ISharePointList): Promise<void> {
    const { spSelectedSite } = this.state;
    if (!spSelectedSite || !this.props.graphClient) {
      return;
    }

    try {
      const columnsData = await this.props.graphClient
        .api(`/sites/${spSelectedSite.id}/lists/${list.id}/columns`)
        .query({ $select: 'name,displayName,columnGroup' })
        .get();

      const rawOptions: IDropdownOption[] = (columnsData.value || [])
        .filter((column: IGraphColumn) => column.name && !column.name.startsWith('_') && column.columnGroup !== '_Hidden')
        .map((column: IGraphColumn) => ({
          key: column.name as string,
          text: column.displayName || (column.name as string)
        }));

      this.setState({
        spAvailableFields: rawOptions,
        addingCalendarStep: 'sharepoint-fields',
        spFieldMapping: {
          titleField: this.findBestMatchingFieldKey(rawOptions, this.getFieldCandidates('title')),
          startDateField: this.findBestMatchingFieldKey(rawOptions, this.getFieldCandidates('start')),
          endDateField: this.findBestMatchingFieldKey(rawOptions, this.getFieldCandidates('end'))
        }
      });
    } catch (error) {
      console.error('Failed to load SharePoint field metadata:', error);
      this.setState({ addingCalendarStep: 'sharepoint-fields' });
    }
  }

  private parseFieldCandidates(rawValue: string): string[] {
    return (rawValue || '')
      .split(',')
      .map(value => value.trim())
      .filter(Boolean);
  }

  private getFieldCandidates(field: 'title' | 'start' | 'end'): string[] {
    const defaults = {
      title: ['Title', 'Subject', 'Event Title'],
      start: ['Start Time', 'Start', 'Start Date', 'StartDate', 'StartDateTime', 'EventDate', 'Starttijd', 'Begindatum', 'Begin', 'Startdatum'],
      end: ['End Time', 'End', 'End Date', 'EndDate', 'EndDateTime', 'Eindtijd', 'Einddatum', 'Einde']
    };

    const localized = this.parseFieldCandidates(
      field === 'title' ? strings.FieldTitleCandidates
        : field === 'start' ? strings.FieldStartCandidates
          : strings.FieldEndCandidates
    );

    return [...localized, ...defaults[field]];
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
      const exactMatch = normalizedOptions.find(entry => entry.text === candidate || entry.key === candidate);
      if (exactMatch) {
        return exactMatch.option.key as string;
      }
    }

    for (const candidate of normalizedCandidates) {
      const partialMatch = normalizedOptions.find(entry => entry.text.includes(candidate) || entry.key.includes(candidate));
      if (partialMatch) {
        return partialMatch.option.key as string;
      }
    }

    return undefined;
  }

  private handleExchangeMailboxChange = (value?: string): void => {
    this.setState({ exchangeMailbox: value || '' });
  };

  private handleExchangeLookupMailbox = async (): Promise<void> => {
    if (!this.state.exchangeMailbox.trim()) {
      return;
    }

    this.setState({ exchangeCalendarsLoading: true });
    const resolved = await this.exchangeService?.resolveMailbox(this.state.exchangeMailbox);

    if (!resolved) {
      this.setState({ exchangeCalendarsLoading: false, exchangeMailboxResolved: false });
      alert('Mailbox not found or not accessible');
      return;
    }

    const calendars = await this.exchangeService?.getCalendars(this.state.exchangeMailbox) || [];
    this.setState({
      exchangeCalendars: calendars,
      exchangeCalendarsLoading: false,
      exchangeMailboxResolved: true,
      addingCalendarStep: 'exchange-calendar'
    });
  };

  private handleSelectExchangeCalendar = (calendar: IExchangeCalendar): void => {
    this.setState({
      exchangeSelectedCalendarId: calendar.id,
      newCalendarName: calendar.name,
      newCalendarColor: calendar.hexColor
    });
  };

  private handleSelectPlannerPlan = (planId: string, planTitle: string): void => {
    this.setState({
      plannerSelectedPlanId: planId,
      addingCalendarStep: 'planner-options',
      newCalendarName: planTitle
    });
  };

  private handleToggleUnifiedGroupSelection = (groupId: string, checked?: boolean): void => {
    this.setState(prev => ({
      unifiedGroupsSelection: {
        ...prev.unifiedGroupsSelection,
        [groupId]: !!checked
      }
    }));
  };

  private beginAudienceSelectionForSources = async (
    sources: ICalendarSourceBase[],
    target?: { kind: 'source'; id: string }
  ): Promise<void> => {
    const selectedAudienceGroups: Record<string, string> = {};
    if (target) {
      const existing = this.state.settings.assignedSources.find(item => item.adminSourceId === target.id);
      existing?.audienceGroups.forEach(group => {
        selectedAudienceGroups[group.groupId] = group.displayName;
      });
    }

    this.setState({
      pendingAdminSources: sources,
      pendingAdminIcs: undefined,
      audienceEditTarget: target,
      selectedAudienceGroups,
      addingCalendarStep: 'admin-audience-select',
      securityGroupsLoading: true
    });

    const securityGroups = await this.audienceService?.getSecurityGroups() || [];
    this.setState({ securityGroups, securityGroupsLoading: false });
  };

  private beginAudienceSelectionForIcs = async (
    item: { adminIcsId?: string; displayName: string; icsUrl: string },
    target?: { kind: 'ics'; id: string }
  ): Promise<void> => {
    const selectedAudienceGroups: Record<string, string> = {};
    if (target) {
      const existing = this.state.settings.icsCatalog.find(entry => entry.adminIcsId === target.id);
      existing?.audienceGroups.forEach(group => {
        selectedAudienceGroups[group.groupId] = group.displayName;
      });
    }

    this.setState({
      pendingAdminSources: [],
      pendingAdminIcs: item,
      audienceEditTarget: target,
      selectedAudienceGroups,
      addingCalendarStep: 'admin-audience-select',
      securityGroupsLoading: true
    });

    const securityGroups = await this.audienceService?.getSecurityGroups() || [];
    this.setState({ securityGroups, securityGroupsLoading: false });
  };

  private createSelectedAudienceGroups(): IAudienceGroup[] {
    return Object.keys(this.state.selectedAudienceGroups).map(groupId => ({
      groupId,
      displayName: this.state.selectedAudienceGroups[groupId]
    }));
  }

  private handleConfirmSharePointCalendar = async (): Promise<void> => {
    if (!this.state.spSelectedSite || !this.state.spSelectedList) {
      return;
    }

    const source: ICalendarSourceBase = {
      sourceType: 'sharepoint',
      name: this.state.newCalendarName,
      color: this.state.newCalendarColor,
      isEnabled: true,
      sharePointSiteId: this.state.spSelectedSite.id,
      sharePointListId: this.state.spSelectedList.id,
      sharePointFieldMapping: this.state.spFieldMapping
    };

    await this.beginAudienceSelectionForSources([source]);
  };

  private handleConfirmExchangeCalendar = async (): Promise<void> => {
    const source: ICalendarSourceBase = {
      sourceType: 'exchange',
      name: this.state.newCalendarName,
      color: this.state.newCalendarColor,
      isEnabled: true,
      exchangeMailbox: this.state.exchangeMailbox || undefined,
      exchangeCalendarId: this.state.exchangeSelectedCalendarId || 'calendar'
    };

    await this.beginAudienceSelectionForSources([source]);
  };

  private handleConfirmPlannerPlan = async (): Promise<void> => {
    const selectedPlan = this.state.plannerPlans.find(plan => plan.id === this.state.plannerSelectedPlanId);
    if (!selectedPlan || !this.state.newCalendarName.trim()) {
      return;
    }

    const source: ICalendarSourceBase = {
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

    await this.beginAudienceSelectionForSources([source]);
  };

  private handleConfirmTeamsShifts = async (): Promise<void> => {
    if (!this.state.newCalendarName.trim()) {
      return;
    }

    const source: ICalendarSourceBase = {
      sourceType: 'teamsShifts',
      name: this.state.newCalendarName.trim(),
      color: this.state.newCalendarColor,
      isEnabled: true,
      showSourceLogo: this.state.teamsShiftsShowLogo
    };

    await this.beginAudienceSelectionForSources([source]);
  };

  private handleConfirmUnifiedGroups = async (): Promise<void> => {
    const selectedIds = Object.keys(this.state.unifiedGroupsSelection).filter(id => this.state.unifiedGroupsSelection[id]);
    if (selectedIds.length === 0) {
      return;
    }

    const selectedGroups = this.state.unifiedGroups.filter(group => selectedIds.indexOf(group.id) >= 0);
    const sources: ICalendarSourceBase[] = selectedGroups.map(group => ({
      sourceType: 'unifiedGroup',
      name: group.displayName,
      color: this.state.newCalendarColor,
      isEnabled: true,
      groupId: group.id,
      showSourceLogo: true
    }));

    await this.beginAudienceSelectionForSources(sources);
  };

  private handleConfirmAdminIcsItem = async (): Promise<void> => {
    if (!this.state.icsUrl.trim() || !this.state.newCalendarName.trim()) {
      return;
    }

    await this.beginAudienceSelectionForIcs({
      displayName: this.state.newCalendarName.trim(),
      icsUrl: this.state.icsUrl.trim()
    });
  };

  private handleSecurityGroupSearch = async (): Promise<void> => {
    this.setState({ securityGroupsLoading: true });
    const securityGroups = await this.audienceService?.getSecurityGroups(this.state.securityGroupSearch) || [];
    this.setState({ securityGroups, securityGroupsLoading: false });
  };

  private handleToggleAudienceGroup = (group: IEntraSecurityGroup, checked?: boolean): void => {
    this.setState(prev => {
      const selectedAudienceGroups = { ...prev.selectedAudienceGroups };
      if (checked) {
        selectedAudienceGroups[group.id] = group.displayName;
      } else {
        delete selectedAudienceGroups[group.id];
      }
      return { selectedAudienceGroups };
    });
  };

  private handleApplyAudienceSelection = (): void => {
    const audienceGroups = this.createSelectedAudienceGroups();
    if (audienceGroups.length === 0) {
      return;
    }

    if (this.state.audienceEditTarget?.kind === 'source') {
      const settings = {
        ...this.state.settings,
        assignedSources: this.state.settings.assignedSources.map(item =>
          item.adminSourceId === this.state.audienceEditTarget?.id
            ? { ...item, audienceGroups }
            : item
        )
      };
      this.setState({ settings }, () => this.handleCloseAddDialog());
      return;
    }

    if (this.state.audienceEditTarget?.kind === 'ics' && this.state.pendingAdminIcs) {
      const settings = {
        ...this.state.settings,
        icsCatalog: this.state.settings.icsCatalog.map(item =>
          item.adminIcsId === this.state.audienceEditTarget?.id
            ? {
              ...item,
              displayName: this.state.pendingAdminIcs?.displayName || item.displayName,
              icsUrl: this.state.pendingAdminIcs?.icsUrl || item.icsUrl,
              audienceGroups
            }
            : item
        )
      };
      this.setState({ settings }, () => this.handleCloseAddDialog());
      return;
    }

    if (this.state.pendingAdminIcs) {
      const newItem: IAdminIcsCatalogItem = {
        adminIcsId: this.state.pendingAdminIcs.adminIcsId || generateStableId('adminIcs'),
        displayName: this.state.pendingAdminIcs.displayName,
        icsUrl: this.state.pendingAdminIcs.icsUrl,
        audienceGroups
      };
      const settings = {
        ...this.state.settings,
        icsCatalog: [...this.state.settings.icsCatalog, newItem]
      };
      this.setState({ settings }, () => this.handleCloseAddDialog());
      return;
    }

    const newSources = this.state.pendingAdminSources.map(source => createAdminAssignedSource(source, audienceGroups));
    const settings = {
      ...this.state.settings,
      assignedSources: [...this.state.settings.assignedSources, ...newSources]
    };
    this.setState({ settings }, () => this.handleCloseAddDialog());
  };

  private handleUpdateAssignedSource = (adminSourceId: string, updates: Partial<ICalendarSourceBase>): void => {
    this.setState(prev => ({
      settings: {
        ...prev.settings,
        assignedSources: prev.settings.assignedSources.map(item =>
          item.adminSourceId === adminSourceId
            ? { ...item, source: { ...item.source, ...updates } }
            : item
        )
      }
    }));
  };

  private handleDeleteAssignedSource = (adminSourceId: string): void => {
    this.setState(prev => ({
      settings: {
        ...prev.settings,
        assignedSources: prev.settings.assignedSources.filter(item => item.adminSourceId !== adminSourceId)
      },
      editingSourceId: prev.editingSourceId === adminSourceId ? undefined : prev.editingSourceId
    }));
  };

  private handleEditAssignedSourceAudiences = async (adminSourceId: string): Promise<void> => {
    const existing = this.state.settings.assignedSources.find(item => item.adminSourceId === adminSourceId);
    if (!existing) {
      return;
    }

    await this.beginAudienceSelectionForSources([existing.source], { kind: 'source', id: adminSourceId });
    this.setState({ showAddDialog: true });
  };

  private handleUpdateIcsCatalogItem = (adminIcsId: string, updates: Partial<IAdminIcsCatalogItem>): void => {
    this.setState(prev => ({
      settings: {
        ...prev.settings,
        icsCatalog: prev.settings.icsCatalog.map(item =>
          item.adminIcsId === adminIcsId
            ? { ...item, ...updates }
            : item
        )
      }
    }));
  };

  private handleDeleteIcsCatalogItem = (adminIcsId: string): void => {
    this.setState(prev => ({
      settings: {
        ...prev.settings,
        icsCatalog: prev.settings.icsCatalog.filter(item => item.adminIcsId !== adminIcsId)
      },
      editingIcsId: prev.editingIcsId === adminIcsId ? undefined : prev.editingIcsId
    }));
  };

  private handleEditIcsAudiences = async (adminIcsId: string): Promise<void> => {
    const item = this.state.settings.icsCatalog.find(entry => entry.adminIcsId === adminIcsId);
    if (!item) {
      return;
    }

    await this.beginAudienceSelectionForIcs({
      adminIcsId: item.adminIcsId,
      displayName: item.displayName,
      icsUrl: item.icsUrl
    }, { kind: 'ics', id: adminIcsId });
    this.setState({ showAddDialog: true });
  };

  private toggleEditSource = (id: string | undefined): void => {
    this.setState({ editingSourceId: id });
  };

  private toggleEditIcs = (id: string | undefined): void => {
    this.setState({ editingIcsId: id });
  };

  private renderSharePointFlow(): React.ReactElement {
    const { spSites, spSitesLoading, spSiteFilter, spSelectedSite, spLists, spListsLoading, spSelectedList, spAvailableFields, spFieldMapping, addingCalendarStep } = this.state;

    if (spSelectedList && addingCalendarStep === 'sharepoint-fields') {
      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>Map SharePoint fields to calendar fields</Label>
          <Dropdown label="Title/Subject field" options={spAvailableFields} selectedKey={spFieldMapping.titleField || ''} onChange={(_, option) => this.setState({ spFieldMapping: { ...spFieldMapping, titleField: option?.key as string } })} />
          <Dropdown label="Start Date field" options={spAvailableFields} selectedKey={spFieldMapping.startDateField || ''} onChange={(_, option) => this.setState({ spFieldMapping: { ...spFieldMapping, startDateField: option?.key as string } })} />
          <Dropdown label="End Date field" options={spAvailableFields} selectedKey={spFieldMapping.endDateField || ''} onChange={(_, option) => this.setState({ spFieldMapping: { ...spFieldMapping, endDateField: option?.key as string } })} />
          <Dropdown label="Location field (optional)" options={[{ key: '', text: '(none)' }, ...spAvailableFields]} selectedKey={spFieldMapping.locationField || ''} onChange={(_, option) => this.setState({ spFieldMapping: { ...spFieldMapping, locationField: option?.key as string } })} />
          <Dropdown label="Description field (optional)" options={[{ key: '', text: '(none)' }, ...spAvailableFields]} selectedKey={spFieldMapping.descriptionField || ''} onChange={(_, option) => this.setState({ spFieldMapping: { ...spFieldMapping, descriptionField: option?.key as string } })} />
          <TextField label="Calendar Name" value={this.state.newCalendarName} onChange={(_, value) => this.setState({ newCalendarName: value || '' })} />
          <div>
            <Label>Color</Label>
            <ColorPicker color={this.state.newCalendarColor} onChange={(_, color) => this.setState({ newCalendarColor: `#${color.hex}` })} alphaType="none" />
          </div>
          <PrimaryButton text="Next: choose groups" onClick={() => this.handleConfirmSharePointCalendar().catch(err => console.error(err))} />
        </Stack>
      );
    }

    if (spSelectedSite) {
      if (spListsLoading) {
        return <Spinner size={SpinnerSize.medium} label="Loading lists..." />;
      }

      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>Select a calendar list from {spSelectedSite.name}:</Label>
          <Stack tokens={{ childrenGap: 8 }}>
            {spLists.map(list => (
              <DefaultButton key={list.id} text={list.name} onClick={() => this.handleSelectSharePointList(list)} style={{ textAlign: 'left', height: 'auto', padding: '8px' }} />
            ))}
          </Stack>
        </Stack>
      );
    }

    if (spSitesLoading) {
      return <Spinner size={SpinnerSize.medium} label="Loading sites..." />;
    }

    const totalPages = Math.ceil(spSites.length / this.SITES_PER_PAGE);
    const startIndex = this.state.spCurrentPage * this.SITES_PER_PAGE;
    const sitesOnPage = spSites.slice(startIndex, startIndex + this.SITES_PER_PAGE);

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>Search and select a SharePoint site:</Label>
        <TextField placeholder="Type to filter sites..." value={spSiteFilter} onChange={(_, value) => this.handleSharePointFilterChange(value)} />
        <PrimaryButton text="Search" onClick={() => this.handleSharePointSearch().catch(err => console.error(err))} />
        <Stack tokens={{ childrenGap: 8 }}>
          {sitesOnPage.map(site => (
            <Stack key={site.id} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} onClick={() => this.handleSelectSharePointSite(site).catch(err => console.error(err))} style={{ border: '1px solid #edebe9', borderRadius: 4, padding: '8px 12px', backgroundColor: '#f3f2f1', cursor: 'pointer' }}>
              <div style={{ flex: 1 }}>
                <strong>{site.name}</strong>
                <div style={{ fontSize: 12, color: '#605e5c' }}>{site.url}</div>
              </div>
              <Icon iconName="ChevronRight" />
            </Stack>
          ))}
        </Stack>
        {totalPages > 1 && (
          <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="center">
            <DefaultButton text="Previous" disabled={this.state.spCurrentPage === 0} onClick={() => this.setState({ spCurrentPage: this.state.spCurrentPage - 1 })} />
            <Label>Page {this.state.spCurrentPage + 1} of {totalPages}</Label>
            <DefaultButton text="Next" disabled={this.state.spCurrentPage >= totalPages - 1} onClick={() => this.setState({ spCurrentPage: this.state.spCurrentPage + 1 })} />
          </Stack>
        )}
      </Stack>
    );
  }

  private renderExchangeFlow(): React.ReactElement {
    const { exchangeMailbox, exchangeMailboxResolved, exchangeCalendars, exchangeCalendarsLoading, exchangeSelectedCalendarId } = this.state;

    if (exchangeSelectedCalendarId) {
      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <TextField label="Calendar Name" value={this.state.newCalendarName} onChange={(_, value) => this.setState({ newCalendarName: value || '' })} />
          <div>
            <Label>Color</Label>
            <ColorPicker color={this.state.newCalendarColor} onChange={(_, color) => this.setState({ newCalendarColor: `#${color.hex}` })} alphaType="none" />
          </div>
          <PrimaryButton text="Next: choose groups" onClick={() => this.handleConfirmExchangeCalendar().catch(err => console.error(err))} />
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
              <Stack key={cal.id} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ border: '1px solid #edebe9', borderRadius: 4, padding: '8px 12px', backgroundColor: '#f3f2f1', cursor: 'pointer' }} onClick={() => this.handleSelectExchangeCalendar(cal)}>
                <div style={{ width: 16, height: 16, backgroundColor: cal.hexColor, borderRadius: 2, flexShrink: 0 }} />
                <div style={{ flex: 1 }}>
                  <strong>{cal.name}</strong>
                  {cal.isDefaultCalendar && <span style={{ fontSize: 11, color: '#605e5c', marginLeft: 8 }}>(Default)</span>}
                </div>
              </Stack>
            ))}
          </Stack>
        </Stack>
      );
    }

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>Enter a mailbox email to load calendars:</Label>
        <TextField placeholder="user@example.com" value={exchangeMailbox} onChange={(_, value) => this.handleExchangeMailboxChange(value)} />
        <PrimaryButton text="Load Mailbox Calendars" onClick={() => this.handleExchangeLookupMailbox().catch(err => console.error(err))} />
      </Stack>
    );
  }

  private renderIcsFlow(): React.ReactElement {
    const hasValidInput = this.state.icsUrl.trim() && this.state.newCalendarName.trim();

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <TextField label="Display Name" value={this.state.newCalendarName} onChange={(_, value) => this.setState({ newCalendarName: value || '' })} placeholder="e.g. Office Holidays" />
        <TextField label="ICS URL" value={this.state.icsUrl} onChange={(_, value) => this.setState({ icsUrl: value || '' })} placeholder="https://example.com/calendar.ics" />
        <PrimaryButton text="Next: choose groups" onClick={() => this.handleConfirmAdminIcsItem().catch(err => console.error(err))} disabled={!hasValidInput} />
      </Stack>
    );
  }

  private renderPlannerFlow(): React.ReactElement {
    const { addingCalendarStep, plannerPlans, plannerPlansLoading } = this.state;

    if (addingCalendarStep === 'planner-plan') {
      if (plannerPlansLoading) {
        return <Spinner size={SpinnerSize.large} label="Loading your Planner plans..." />;
      }

      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>Select a Planner plan to add:</Label>
          <Stack tokens={{ childrenGap: 8 }}>
            {plannerPlans.map(plan => (
              <div key={plan.id} onClick={() => this.handleSelectPlannerPlan(plan.id, plan.title)} style={{ padding: 12, border: '1px solid #ddd', borderRadius: 4, cursor: 'pointer', backgroundColor: this.state.plannerSelectedPlanId === plan.id ? '#f3f2f1' : 'white' }}>
                <Icon iconName="PlannerLogo" style={{ marginRight: 8, fontSize: 16 }} />
                <strong>{plan.title}</strong>
              </div>
            ))}
          </Stack>
        </Stack>
      );
    }

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <TextField label="Calendar Name" value={this.state.newCalendarName} onChange={(_, value) => this.setState({ newCalendarName: value || '' })} required />
        <Toggle label="Alleen aan mij toegewezen taken" checked={this.state.plannerAssignedToMeOnly} onChange={(_, checked) => this.setState({ plannerAssignedToMeOnly: !!checked })} />
        <Toggle label="Voltooide taken weergeven" checked={this.state.plannerShowCompleted} onChange={(_, checked) => this.setState({ plannerShowCompleted: !!checked })} />
        <Toggle label="Bron logo tonen" checked={this.state.plannerShowLogo} onChange={(_, checked) => this.setState({ plannerShowLogo: !!checked })} />
        <ColorPicker color={this.state.newCalendarColor} onChange={(_, color) => this.setState({ newCalendarColor: `#${color.hex}` })} alphaType="none" />
        <PrimaryButton text="Next: choose groups" onClick={() => this.handleConfirmPlannerPlan().catch(err => console.error(err))} disabled={!this.state.newCalendarName.trim()} />
      </Stack>
    );
  }

  private renderUnifiedGroupsFlow(): React.ReactElement {
    const { unifiedGroups, unifiedGroupsLoading, unifiedGroupsSelection } = this.state;
    const selectedCount = Object.keys(unifiedGroupsSelection).filter(id => unifiedGroupsSelection[id]).length;

    if (unifiedGroupsLoading) {
      return <Spinner size={SpinnerSize.large} label="Loading groups and teams..." />;
    }

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>Select one or more groups or teams:</Label>
        <Stack tokens={{ childrenGap: 8 }}>
          {unifiedGroups.map(group => (
            <Stack key={group.id} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ padding: '8px 12px', border: '1px solid #edebe9', borderRadius: 4, backgroundColor: unifiedGroupsSelection[group.id] ? '#f3f2f1' : 'white' }}>
              <Icon iconName={group.isTeam ? 'TeamsLogo' : 'Group'} style={{ fontSize: 16 }} />
              <Checkbox label={group.displayName} checked={!!unifiedGroupsSelection[group.id]} onChange={(_, checked) => this.handleToggleUnifiedGroupSelection(group.id, checked)} />
            </Stack>
          ))}
        </Stack>
        <div>
          <Label>Color</Label>
          <ColorPicker color={this.state.newCalendarColor} onChange={(_, color) => this.setState({ newCalendarColor: `#${color.hex}` })} alphaType="none" />
        </div>
        <PrimaryButton text={selectedCount > 1 ? 'Next: choose groups for calendars' : 'Next: choose groups for calendar'} onClick={() => this.handleConfirmUnifiedGroups().catch(err => console.error(err))} disabled={selectedCount === 0} />
      </Stack>
    );
  }

  private renderTeamsShiftsFlow(): React.ReactElement {
    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <TextField label="Calendar Name" value={this.state.newCalendarName} onChange={(_, value) => this.setState({ newCalendarName: value || '' })} placeholder="e.g. Teams Shifts" required />
        <Toggle label="Bron logo tonen" checked={this.state.teamsShiftsShowLogo} onChange={(_, checked) => this.setState({ teamsShiftsShowLogo: !!checked })} />
        <ColorPicker color={this.state.newCalendarColor} onChange={(_, color) => this.setState({ newCalendarColor: `#${color.hex}` })} alphaType="none" showPreview={true} />
        <PrimaryButton text="Next: choose groups" onClick={() => this.handleConfirmTeamsShifts().catch(err => console.error(err))} disabled={!this.state.newCalendarName.trim()} />
      </Stack>
    );
  }

  private renderAudienceSelectionFlow(): React.ReactElement {
    const selectedCount = Object.keys(this.state.selectedAudienceGroups).length;
    const selectedGroupNames = Object.keys(this.state.selectedAudienceGroups).map(groupId => this.state.selectedAudienceGroups[groupId]);

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>Select one or more Entra security groups:</Label>
        <TextField placeholder="Search security groups..." value={this.state.securityGroupSearch} onChange={(_, value) => this.setState({ securityGroupSearch: value || '' })} />
        <PrimaryButton text="Search" onClick={() => this.handleSecurityGroupSearch().catch(err => console.error(err))} />
        {selectedCount > 0 && (
          <MessageBar messageBarType={MessageBarType.info}>
            Selected groups: {selectedGroupNames.join(', ')}
          </MessageBar>
        )}
        {this.state.securityGroupsLoading ? (
          <Spinner size={SpinnerSize.medium} label="Loading security groups..." />
        ) : (
          <Stack tokens={{ childrenGap: 8 }}>
            {this.state.securityGroups.map(group => (
              <Checkbox key={group.id} label={group.displayName} checked={!!this.state.selectedAudienceGroups[group.id]} onChange={(_, checked) => this.handleToggleAudienceGroup(group, checked)} />
            ))}
          </Stack>
        )}
        <PrimaryButton text={this.state.audienceEditTarget ? 'Apply groups' : 'Add item'} onClick={this.handleApplyAudienceSelection} disabled={selectedCount === 0} />
      </Stack>
    );
  }

  private renderAddFlow(): React.ReactElement {
    const { addingCalendarStep, addingCalendarType } = this.state;

    if (addingCalendarStep === 'initial') {
      return (
        <Stack tokens={{ childrenGap: 16 }}>
          <Label>Select the type of admin item to add:</Label>
          <Stack tokens={{ childrenGap: 12 }}>
            {calendarSourceRegistry.filter(definition => definition.adminSelectable).map(definition => (
              <PrimaryButton
                key={definition.type}
                text={definition.displayName}
                secondaryText={definition.adminCatalogOnly ? 'Publish this ICS feed as a selectable catalog item' : definition.description}
                iconProps={{ iconName: definition.iconName }}
                onClick={() => this.handleSelectAddType(definition.type).catch(err => console.error(err))}
                style={{ textAlign: 'left', height: 'auto', padding: '12px' }}
              />
            ))}
          </Stack>
          <DefaultButton text="Cancel" onClick={this.handleCloseAddDialog} />
        </Stack>
      );
    }

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        {this.renderNavigationHeader()}
        {addingCalendarStep === 'admin-audience-select' && this.renderAudienceSelectionFlow()}
        {addingCalendarType === 'sharepoint' && addingCalendarStep !== 'admin-audience-select' && this.renderSharePointFlow()}
        {addingCalendarType === 'exchange' && addingCalendarStep !== 'admin-audience-select' && this.renderExchangeFlow()}
        {addingCalendarType === 'planner' && addingCalendarStep !== 'admin-audience-select' && this.renderPlannerFlow()}
        {addingCalendarType === 'unifiedGroup' && addingCalendarStep !== 'admin-audience-select' && this.renderUnifiedGroupsFlow()}
        {addingCalendarType === 'teamsShifts' && addingCalendarStep !== 'admin-audience-select' && this.renderTeamsShiftsFlow()}
        {addingCalendarType === 'ics' && addingCalendarStep !== 'admin-audience-select' && this.renderIcsFlow()}
      </Stack>
    );
  }

  private renderAssignedSource(item: IAdminAssignedSource, index: number): React.ReactElement {
    const isEditing = this.state.editingSourceId === item.adminSourceId;
    const audienceText = item.audienceGroups.map(group => group.displayName).join(', ');

    return (
      <div key={item.adminSourceId} style={{ borderRadius: 4, padding: isEditing ? 12 : '6px 8px', backgroundColor: isEditing ? '#f3f2f1' : (index % 2 === 1 ? 'rgba(0,0,0,0.02)' : 'transparent') }}>
        {isEditing ? (
          <Stack tokens={{ childrenGap: 8 }}>
            <TextField label="Name" value={item.source.name} onChange={(_, value) => this.handleUpdateAssignedSource(item.adminSourceId, { name: value || '' })} />
            <div>
              <Label>Color</Label>
              <ColorPicker color={item.source.color} onChange={(_, color) => this.handleUpdateAssignedSource(item.adminSourceId, { color: `#${color.hex}` })} alphaType="none" />
            </div>
            <Toggle label="Enabled" checked={item.source.isEnabled} onChange={(_, checked) => this.handleUpdateAssignedSource(item.adminSourceId, { isEnabled: !!checked })} />
            <div style={{ fontSize: 12, color: '#605e5c' }}>Audiences: {audienceText || 'None'}</div>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton text="Groups" onClick={() => this.handleEditAssignedSourceAudiences(item.adminSourceId).catch(err => console.error(err))} />
              <PrimaryButton text="Done" onClick={() => this.toggleEditSource(undefined)} />
              <DefaultButton text="Delete" onClick={() => this.handleDeleteAssignedSource(item.adminSourceId)} />
            </Stack>
          </Stack>
        ) : (
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <div style={{ width: 16, height: 16, backgroundColor: item.source.color, borderRadius: 2, flexShrink: 0 }} />
            <div style={{ flex: 1 }}>
              <strong style={{ fontSize: 13 }}>{item.source.name}</strong>
              <div style={{ fontSize: 11, color: '#605e5c' }}>{getSourceTypeDisplayName(item.source.sourceType)} • {audienceText}</div>
            </div>
            <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={() => this.toggleEditSource(item.adminSourceId)} />
          </Stack>
        )}
      </div>
    );
  }

  private renderIcsCatalogItem(item: IAdminIcsCatalogItem, index: number): React.ReactElement {
    const isEditing = this.state.editingIcsId === item.adminIcsId;
    const audienceText = item.audienceGroups.map(group => group.displayName).join(', ');

    return (
      <div key={item.adminIcsId} style={{ borderRadius: 4, padding: isEditing ? 12 : '6px 8px', backgroundColor: isEditing ? '#f3f2f1' : (index % 2 === 1 ? 'rgba(0,0,0,0.02)' : 'transparent') }}>
        {isEditing ? (
          <Stack tokens={{ childrenGap: 8 }}>
            <TextField label="Display Name" value={item.displayName} onChange={(_, value) => this.handleUpdateIcsCatalogItem(item.adminIcsId, { displayName: value || '' })} />
            <TextField label="ICS URL" value={item.icsUrl} onChange={(_, value) => this.handleUpdateIcsCatalogItem(item.adminIcsId, { icsUrl: value || '' })} />
            <div style={{ fontSize: 12, color: '#605e5c' }}>Audiences: {audienceText || 'None'}</div>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton text="Groups" onClick={() => this.handleEditIcsAudiences(item.adminIcsId).catch(err => console.error(err))} />
              <PrimaryButton text="Done" onClick={() => this.toggleEditIcs(undefined)} />
              <DefaultButton text="Delete" onClick={() => this.handleDeleteIcsCatalogItem(item.adminIcsId)} />
            </Stack>
          </Stack>
        ) : (
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <Icon iconName="World" />
            <div style={{ flex: 1 }}>
              <strong style={{ fontSize: 13 }}>{item.displayName}</strong>
              <div style={{ fontSize: 11, color: '#605e5c' }}>{item.icsUrl}</div>
              <div style={{ fontSize: 11, color: '#605e5c' }}>{audienceText}</div>
            </div>
            <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={() => this.toggleEditIcs(item.adminIcsId)} />
          </Stack>
        )}
      </div>
    );
  }

  private onRenderFooterContent = (): React.ReactElement => (
    <Stack horizontal tokens={{ childrenGap: 8 }}>
      <PrimaryButton onClick={this.handleSave} text="Save" />
      <DefaultButton onClick={this.props.onDismiss} text="Cancel" />
      <DefaultButton onClick={this.handleResetDraft} text="Reset Draft to Defaults" />
    </Stack>
  );

  public render(): React.ReactElement {
    const { isOpen, onDismiss, loadNotice } = this.props;
    const { settings, showAddDialog } = this.state;
    const startOptions: IDropdownOption[] = [];
    const latestStart = Math.max(0, 24 * 60 - settings.visibleHourCount * 60);
    for (let minutes = 0; minutes <= latestStart; minutes += settings.slotDurationMinutes) {
      startOptions.push({ key: minutes, text: new Date(2000, 0, 1, 0, minutes).toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' }) });
    }
    const visibleHourOptions = Array.from({ length: 24 }, (_, index) => ({ key: index + 1, text: String(index + 1) }));

    return (
      <Panel isOpen={isOpen} onDismiss={onDismiss} type={PanelType.medium} headerText={showAddDialog ? 'Add Admin Calendar Default' : 'Admin Calendar Defaults'} onRenderFooterContent={!showAddDialog ? this.onRenderFooterContent : undefined} isFooterAtBottom={true}>
        {showAddDialog ? (
          this.renderAddFlow()
        ) : (
          <Stack tokens={{ childrenGap: 16 }}>
            {loadNotice && (
              <MessageBar messageBarType={MessageBarType.warning}>
                {loadNotice}
              </MessageBar>
            )}

            <Dropdown
              label="Default View"
              options={[
                { key: 'day', text: 'Day' },
                { key: 'week', text: 'Week' },
                { key: 'month', text: 'Month' }
              ]}
              selectedKey={settings.defaultView}
              onChange={(_, option) => this.setState(prev => ({
                settings: {
                  ...prev.settings,
                  defaultView: option?.key as IAdminWebPartSettings['defaultView']
                }
              }))}
            />

            <Label>Timeline defaults</Label>
            <Toggle
              label="Show weekends"
              checked={settings.showWeekends}
              onChange={(_, checked) => this.setState(prev => ({ settings: { ...prev.settings, showWeekends: checked !== false } }))}
            />
            <Dropdown
              label="Slot duration"
              selectedKey={settings.slotDurationMinutes}
              options={[{ key: 15, text: '15 minutes' }, { key: 30, text: '30 minutes' }, { key: 60, text: '60 minutes' }]}
              onChange={(_, option) => {
                const slot = option?.key as 15 | 30 | 60;
                this.setState(prev => ({ settings: { ...prev.settings, slotDurationMinutes: slot, preferredStartMinutes: Math.floor(Math.min(prev.settings.preferredStartMinutes, 24 * 60 - prev.settings.visibleHourCount * 60) / slot) * slot } }));
              }}
            />
            <Dropdown
              label="Preferred start time"
              selectedKey={settings.preferredStartMinutes}
              options={startOptions}
              onChange={(_, option) => this.setState(prev => ({ settings: { ...prev.settings, preferredStartMinutes: Number(option?.key) } }))}
            />
            <Dropdown
              label="Visible hours"
              selectedKey={settings.visibleHourCount}
              options={visibleHourOptions}
              onChange={(_, option) => {
                const visibleHourCount = Number(option?.key);
                this.setState(prev => ({ settings: { ...prev.settings, visibleHourCount, preferredStartMinutes: Math.min(prev.settings.preferredStartMinutes, 24 * 60 - visibleHourCount * 60) } }));
              }}
            />

            <div>
              <PrimaryButton text="Add Admin Default" iconProps={{ iconName: 'Add' }} onClick={this.handleOpenAddDialog} />
            </div>

            <div>
              <Label>Admin default calendars</Label>
              <Stack tokens={{ childrenGap: 8 }}>
                {settings.assignedSources.length > 0
                  ? settings.assignedSources.map((item, index) => this.renderAssignedSource(item, index))
                  : <div style={{ fontSize: 12, color: '#605e5c' }}>No admin default calendars configured.</div>}
              </Stack>
            </div>

            <div>
              <Label>Admin ICS catalog</Label>
              <Stack tokens={{ childrenGap: 8 }}>
                {settings.icsCatalog.length > 0
                  ? settings.icsCatalog.map((item, index) => this.renderIcsCatalogItem(item, index))
                  : <div style={{ fontSize: 12, color: '#605e5c' }}>No admin ICS catalog items configured.</div>}
              </Stack>
            </div>
          </Stack>
        )}
      </Panel>
    );
  }
}
