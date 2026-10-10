import * as React from 'react';
import { AssignmentPolicyControl, AllowedOverrideControls, AssignmentGroup, audienceTypeLabel, assignmentPolicyLabel } from './AssignmentControls';
import { LatestDiscovery } from './latestDiscovery';
import { updateAudienceDiscovery } from './audienceDiscoveryState';
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
import { Slider } from '@fluentui/react/lib/Slider';
import { HttpClient, type MSGraphClientV3 } from '@microsoft/sp-http';
import {
  type IAdminAssignedSource,
  type IAssignmentPolicy,
  type IExchangeMailboxAssignment,
  type IAdminAllowedOverrides,
  defaultAllowedOverrides,
  type IAdminIcsCatalogItem,
  type IAdminWebPartSettings,
  type IAudienceGroup,
  type ICalendarSourceBase,
  type CalendarSourceType,
  defaultAdminWebPartSettings
} from '../models/ICalendarSettings';
import { calendarSourceRegistry } from '../models/CalendarSourceRegistry';
import * as strings from 'MyCalendarsWebPartStrings';
import { ExchangeCalendarService, getExchangeDiscoveryErrorMessage, IExchangeCalendar, IMailboxSearchResult } from '../services/ExchangeCalendarService';
import { SharePointCalendarService, ISharePointList, ISharePointSite } from '../services/SharePointCalendarService';
import { PlannerTaskService, IPlannerPlan } from '../services/PlannerTaskService';
import { UnifiedGroupCalendarService, IUnifiedGroupItem } from '../services/UnifiedGroupCalendarService';
import { AudienceService, IEntraSecurityGroup } from '../services/AudienceService';
import { generateStableId, addAdministratorAssignments, addExchangeMailboxAssignment, canonicalizeExchangeSourceIdentities, getSourceIdentityKey } from '../services/CalendarSettingsService';
import { getSourceTypeDescription, getSourceTypeDisplayName } from '../utils/sourceIconHelper';
import { formatLocalizedString } from '../utils/localization';
import { findBestMatchingFieldKey, getFieldCandidates } from '../utils/sharePointFieldCandidates';
import { formatCalendarTime } from './views/calendarFormatting';
import { MailboxPeoplePicker } from './MailboxPeoplePicker';

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

export interface IAdminDefaultsPanelProps {
  isOpen: boolean;
  onDismiss: () => void;
  settings: IAdminWebPartSettings;
  onSave: (settings: IAdminWebPartSettings) => Promise<void>;
  loadNotice?: string;
  httpClient?: HttpClient;
  graphClient?: MSGraphClientV3;
  locale?: string;
  webAbsoluteUrl?: string;
}

interface IAdminDefaultsPanelState {
  settings: IAdminWebPartSettings;
  isSaving: boolean;
  saveError?: string;
  audienceFirst?: boolean;
  audiencePreset?: boolean;
  everyone?: boolean;
  selectedExchangeCalendars: Record<string, IAssignmentPolicy>;
  selectedSharePointLists: string[];
  pendingSharePointSources: Array<{ source: ICalendarSourceBase; policy: IAssignmentPolicy }>;
  exchangeAll: boolean;
  exchangeAllPolicy: IAssignmentPolicy;
  resolvedMailboxId?: string;
  mailboxRuleCalendars: Record<string, ICalendarSourceBase[]>;
  mailboxRuleErrors: Record<string, string>;
  selectedAudienceMetadata: Record<string, IAudienceGroup>;
  pendingPolicy: IAssignmentPolicy;
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
  exchangeDiscoveryError?: string;
  exchangeMailbox: string;
  exchangeMailboxPerson?: IMailboxSearchResult;
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
  securityGroupsError?: string;
  securityGroupsLoaded?: boolean;
  securityGroupSearch: string;
  selectedAudienceGroups: Record<string, string>;
  pendingAdminSources: ICalendarSourceBase[];
  pendingAdminIcs: { adminIcsId?: string; displayName: string; icsUrl: string } | undefined;
  audienceEditTarget: { kind: 'source'; id: string } | { kind: 'ics'; id: string } | undefined;
}

export class AdminDefaultsPanel extends React.Component<IAdminDefaultsPanelProps, IAdminDefaultsPanelState> {
  private exchangeService: ExchangeCalendarService | null = null;
  private sharePointService: SharePointCalendarService | null = null;
  private plannerService: PlannerTaskService | null = null;
  private unifiedGroupService: UnifiedGroupCalendarService | null = null;
  private audienceService: AudienceService | null = null;
  private mounted = false;
  private savePending = false;
  private editSession = 0;

  private isCurrentSession(session: number): boolean {
    return this.mounted && this.props.isOpen && session === this.editSession && !this.savePending;
  }

  private setSessionState<K extends keyof IAdminDefaultsPanelState>(
    session: number,
    state: Pick<IAdminDefaultsPanelState, K> | IAdminDefaultsPanelState | null | ((previous: Readonly<IAdminDefaultsPanelState>, props: Readonly<IAdminDefaultsPanelProps>) => Pick<IAdminDefaultsPanelState, K> | IAdminDefaultsPanelState | null),
    callback?: () => void
  ): void {
    if (!this.isCurrentSession(session)) return;
    this.setState((previous, props) => this.isCurrentSession(session)
      ? (typeof state === 'function' ? state(previous, props) : state) : null,
    () => { if (this.isCurrentSession(session)) callback?.(); });
  }

  private handleDismiss = (): void => {
    if (!this.savePending) this.props.onDismiss();
  };

  private readonly SITES_PER_PAGE = 20;
  private ruleDiscoveryGeneration = 0;
  private sharePointFieldsGeneration = 0;
  private readonly mailboxDiscovery = new LatestDiscovery();
  private readonly audienceDiscovery = new LatestDiscovery();
  private mailboxSuggestionGeneration = 0;

  constructor(props: IAdminDefaultsPanelProps) {
    super(props);

    if (props.httpClient || props.graphClient) {
      this.exchangeService = new ExchangeCalendarService(props.graphClient);
      this.sharePointService = new SharePointCalendarService(props.graphClient);
      this.plannerService = new PlannerTaskService(props.graphClient);
      this.unifiedGroupService = new UnifiedGroupCalendarService(props.graphClient);
      if (props.graphClient) {
        this.audienceService = new AudienceService(props.graphClient);
      }
    }

    this.state = this.createStateFromProps(props);
  }

  public componentDidMount(): void {
    this.mounted = true;
    if (this.props.graphClient) {
      this.initializeGraphClient(this.props.graphClient);
      this.refreshDraftMetadata().catch(error => console.error('Failed to refresh administrator draft metadata:', error));
    }
  }

  public componentDidUpdate(prevProps: IAdminDefaultsPanelProps): void {
    if (this.props.graphClient && !prevProps.graphClient) {
      this.initializeGraphClient(this.props.graphClient);
      this.refreshDraftMetadata().catch(error => console.error('Failed to refresh administrator draft metadata:', error));
    }

    if (prevProps.isOpen !== this.props.isOpen) {
      this.editSession++;
      this.invalidateDialogDiscovery();
      this.setState({ exchangeCalendarsLoading: false, exchangeDiscoveryError: undefined });
    }
    if (prevProps.isOpen !== this.props.isOpen && this.props.isOpen) {
      this.setState(this.createStateFromProps(this.props), () => {
        this.refreshDraftMetadata().catch(error => console.error('Failed to refresh administrator draft metadata:', error));
      });
    }
  }

  private refreshDraftMetadata = async (): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    await Promise.all([this.enrichSharePointSiteNames(), this.discoverMailboxRules()]);
    if (this.isCurrentSession(session)) await this.canonicalizeMailboxes();
  };

  private createStateFromProps(props: IAdminDefaultsPanelProps): IAdminDefaultsPanelState {
    return {
      settings: structuredClone(props.settings),
      isSaving: false, saveError: undefined,
      audienceFirst: false, audiencePreset: false, everyone: false, resolvedMailboxId: undefined,
      selectedSharePointLists: [], pendingSharePointSources: [],
      selectedExchangeCalendars: {}, exchangeAll: false, exchangeAllPolicy: { isMandatory: false, defaultEnabled: true },
      mailboxRuleCalendars: {}, mailboxRuleErrors: {}, selectedAudienceMetadata: {}, pendingPolicy: { isMandatory: false, defaultEnabled: true },
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
      exchangeMailboxPerson: undefined,
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
      securityGroupsError: undefined, securityGroupsLoaded: false,
      securityGroupSearch: '',
      selectedAudienceGroups: {},
      pendingAdminSources: [],
      pendingAdminIcs: undefined,
      audienceEditTarget: undefined
    };
  }

  private invalidateDialogDiscovery(): void {
    this.ruleDiscoveryGeneration++;
    this.sharePointFieldsGeneration++;
    this.mailboxDiscovery.invalidate();
    this.audienceDiscovery.invalidate();
  }

  public componentWillUnmount(): void {
    this.mounted = false;
    this.editSession++;
    this.invalidateDialogDiscovery();
  }

  private initializeGraphClient(client: MSGraphClientV3): void {
    this.exchangeService = this.exchangeService || new ExchangeCalendarService(client);
    this.exchangeService.setGraphClient(client);
    if (this.sharePointService) this.sharePointService.setGraphClient(client);
    if (this.plannerService) this.plannerService.setGraphClient(client);
    if (this.unifiedGroupService) this.unifiedGroupService.setGraphClient(client);
    this.audienceService = new AudienceService(client);
  }

  private enrichSharePointSiteNames = async (): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    if (!this.sharePointService) return;
    const missingSiteIds = Array.from(new Set(this.state.settings.assignedSources
      .map(item => item.source)
      .filter(source => source.sourceType === 'sharepoint' && source.sharePointSiteId && !source.sharePointSiteName)
      .map(source => source.sharePointSiteId as string)));
    if (missingSiteIds.length === 0) return;
    const resolved = await Promise.all(missingSiteIds.map(async siteId => ({ siteId, site: await this.sharePointService?.getSite(siteId) })));
    if (!this.isCurrentSession(session)) return;
    const names = new Map(resolved.filter(item => item.site?.name).map(item => [item.siteId, item.site?.name as string]));
    if (names.size === 0) return;
    this.setSessionState(session, prev => ({
      settings: {
        ...prev.settings,
        assignedSources: prev.settings.assignedSources.map(item => item.source.sourceType === 'sharepoint' && item.source.sharePointSiteId && !item.source.sharePointSiteName && names.has(item.source.sharePointSiteId)
          ? { ...item, source: { ...item.source, sharePointSiteName: names.get(item.source.sharePointSiteId) } }
          : item)
      }
    }));
  };

  private handleOpenAddDialog = (): void => {
    if (!this.isCurrentSession(this.editSession)) return;
    this.editSession++;
    this.setSessionState(this.editSession, {
      showAddDialog: true,
      addingCalendarType: undefined,
      addingCalendarStep: 'initial'
    });
  };

  private handleCloseAddDialog = (): void => {
    if (!this.isCurrentSession(this.editSession)) return;
    this.editSession++;
    this.invalidateDialogDiscovery();
    this.setSessionState(this.editSession, { exchangeDiscoveryError: undefined, exchangeCalendarsLoading: false });
    this.setSessionState(this.editSession, {
      showAddDialog: false,
      audienceFirst: false, audiencePreset: false, everyone: false, selectedExchangeCalendars: {}, exchangeAll: false, resolvedMailboxId: undefined,
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
      exchangeMailboxPerson: undefined,
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
      securityGroupsError: undefined, securityGroupsLoaded: false,
      securityGroupSearch: '',
      selectedAudienceGroups: {},
      pendingAdminSources: [],
      pendingAdminIcs: undefined,
      audienceEditTarget: undefined
    });
  };

  private handleResetDraft = (): void => {
    if (this.savePending) return;
    this.editSession++;
    if (confirm(strings.ResetAdminDraftConfirmationLabel)) {
      this.invalidateDialogDiscovery();
      this.setSessionState(this.editSession, this.createStateFromProps({
        ...this.props,
        settings: structuredClone(defaultAdminWebPartSettings)
      }));
    }
  };

  private handleSave = async (): Promise<void> => {
    if (this.savePending) return;
    this.savePending = true;
    const session = ++this.editSession;
    this.invalidateDialogDiscovery();
    this.setState({ isSaving: true, saveError: undefined, spSitesLoading: false, spListsLoading: false,
      plannerPlansLoading: false, unifiedGroupsLoading: false, exchangeCalendarsLoading: false, securityGroupsLoading: false });
    try {
      await this.props.onSave(structuredClone(this.state.settings));
      if (this.mounted && session === this.editSession) this.props.onDismiss();
    } catch (error) {
      console.error('Failed to save administrator defaults.', error);
      if (this.mounted && session === this.editSession) this.setState({ saveError: strings.AdminSettingsSaveErrorLabel });
    } finally {
      this.savePending = false;
      if (this.mounted && session === this.editSession) this.setState({ isSaving: false });
    }
  };

  private handleSelectAddType = async (type: CalendarSourceType): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    if ((type === 'exchange' || type === 'sharepoint') && !this.state.audiencePreset) {
      this.setSessionState(session, { addingCalendarType: type, addingCalendarStep: 'admin-audience-select', audienceFirst: true, everyone: false, selectedAudienceGroups: {}, selectedAudienceMetadata: {} });
      await this.loadSecurityGroups();
      if (!this.isCurrentSession(session)) return;
      return;
    }
    await this.handleSelectSourceType(type);
    if (!this.isCurrentSession(session)) return;
  };

  private handleSelectSourceType = async (type: CalendarSourceType): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    if (type === 'sharepoint') {
      this.setSessionState(session, { addingCalendarType: type, addingCalendarStep: 'sharepoint-site', spSitesLoading: true, selectedSharePointLists: [], pendingSharePointSources: [] });
      const sites = await this.sharePointService?.getAccessibleSites() || [];
      if (!this.isCurrentSession(session)) return;
      this.setSessionState(session, { spSites: sites, spSitesLoading: false });
      return;
    }

    if (type === 'exchange') {
      this.setSessionState(session, { addingCalendarType: type, addingCalendarStep: 'exchange-mailbox' });
      return;
    }

    if (type === 'ics') {
      this.setSessionState(session, { addingCalendarType: type, addingCalendarStep: 'ics' });
      return;
    }

    if (type === 'planner') {
      this.setSessionState(session, { addingCalendarType: type, addingCalendarStep: 'planner-plan', plannerPlansLoading: true });
      const plans = await this.plannerService?.getUserPlans() || [];
      if (!this.isCurrentSession(session)) return;
      this.setSessionState(session, { plannerPlans: plans, plannerPlansLoading: false });
      return;
    }

    if (type === 'unifiedGroup') {
      this.setSessionState(session, {
        addingCalendarType: type,
        addingCalendarStep: 'unified-group-select',
        unifiedGroupsLoading: true,
        unifiedGroupsSelection: {},
        newCalendarColor: this.state.settings.organizationPrimaryColor || '#0078d4'
      });
      await this.loadUnifiedGroups();
      if (!this.isCurrentSession(session)) return;
      return;
    }

    if (type === 'teamsShifts') {
      this.setSessionState(session, {
        addingCalendarType: type,
        addingCalendarStep: 'teams-shifts',
        newCalendarName: strings.TeamsShiftsLabel,
        newCalendarColor: this.state.settings.organizationPrimaryColor || '#4a4fbe',
        teamsShiftsShowLogo: true
      });
    }
  };

  private handleBackToTypeSelection = (): void => {
    if (!this.isCurrentSession(this.editSession)) return;
    this.editSession++;
    this.setSessionState(this.editSession, { audienceFirst: false, audiencePreset: false, everyone: false, selectedExchangeCalendars: {}, exchangeAll: false, pendingSharePointSources: [], selectedSharePointLists: [] });
    this.invalidateDialogDiscovery();
    this.setSessionState(this.editSession, { exchangeDiscoveryError: undefined, exchangeCalendarsLoading: false });
    this.setSessionState(this.editSession, {
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
      securityGroupsError: undefined, securityGroupsLoaded: false,
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
    if (!this.isCurrentSession(this.editSession)) return;
    this.editSession++;
    this.setSessionState(this.editSession, { spSitesLoading: false, spListsLoading: false, plannerPlansLoading: false,
      unifiedGroupsLoading: false, exchangeCalendarsLoading: false });
    this.sharePointFieldsGeneration++;
    this.audienceDiscovery.invalidate();
    this.setSessionState(this.editSession, { securityGroupsLoading: false, securityGroupsError: undefined, securityGroupsLoaded: false });
    if (this.state.addingCalendarStep === 'admin-audience-select') {
      if (this.state.audienceEditTarget) {
        this.setSessionState(this.editSession, { showAddDialog: false, audienceEditTarget: undefined, selectedAudienceGroups: {} });
      } else {
        this.setSessionState(this.editSession, {
          addingCalendarStep: this.state.pendingAdminIcs ? 'ics' : 'initial',
          securityGroups: [],
          securityGroupsLoading: false,
          securityGroupsError: undefined, securityGroupsLoaded: false,
          securityGroupSearch: '',
          selectedAudienceGroups: {}
        });
      }
      return;
    }

    if (this.state.addingCalendarType === 'sharepoint') {
      if (this.state.addingCalendarStep === 'sharepoint-fields') {
        this.setSessionState(this.editSession, { spSelectedList: undefined, addingCalendarStep: 'sharepoint-list' });
        return;
      }
      if (this.state.addingCalendarStep === 'sharepoint-list') {
        this.setSessionState(this.editSession, { spSelectedSite: undefined, spLists: [], addingCalendarStep: 'sharepoint-site' });
        return;
      }
    }

    if (this.state.addingCalendarType === 'planner' && this.state.addingCalendarStep === 'planner-options') {
      this.setSessionState(this.editSession, { addingCalendarStep: 'planner-plan' });
      return;
    }

    if (this.state.addingCalendarType === 'exchange' && this.state.addingCalendarStep === 'exchange-calendar') {
      this.mailboxDiscovery.invalidate();
      this.setSessionState(this.editSession, { addingCalendarStep: 'exchange-mailbox', exchangeMailboxResolved: false, exchangeCalendars: [], selectedExchangeCalendars: {}, exchangeAll: false });
      return;
    }
    if ((this.state.addingCalendarType === 'exchange' || this.state.addingCalendarType === 'sharepoint') && this.state.audiencePreset) {
      this.setSessionState(this.editSession, { addingCalendarStep: 'admin-audience-select', audienceFirst: true, audiencePreset: false });
      this.loadSecurityGroups().catch(error => console.error(error));
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
        <IconButton iconProps={{ iconName: 'Back' }} title={strings.BackLabel} ariaLabel={strings.BackLabel} onClick={this.handleBackOneStep} styles={{ root: { height: 32 } }} />
        <IconButton iconProps={{ iconName: 'Home' }} title={strings.HomeLabel} ariaLabel={strings.HomeLabel} onClick={this.handleBackToTypeSelection} styles={{ root: { height: 32 } }} />
      </Stack>
    );
  };

  private async loadUnifiedGroups(): Promise<void> {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    if (!this.unifiedGroupService) {
      this.setSessionState(session, { unifiedGroupsLoading: false });
      return;
    }

    try {
      const [groups, joinedTeamIds] = await Promise.all([
        this.unifiedGroupService.getUnifiedGroups(),
        this.unifiedGroupService.getJoinedTeamIds()
      ]);
      if (!this.isCurrentSession(session)) return;

      const mappedGroups = groups
        .map(group => ({
          ...group,
          isTeam: joinedTeamIds.has(group.id)
        }))
        .sort((a, b) => a.displayName.localeCompare(b.displayName));

      this.setSessionState(session, {
        unifiedGroups: mappedGroups,
        unifiedGroupsLoading: false
      });
    } catch (error) {
      console.error('Failed to load unified groups:', error);
      this.setSessionState(session, {
        unifiedGroups: [],
        unifiedGroupsLoading: false
      });
    }
  }

  private handleSharePointFilterChange = (value?: string): void => {
    this.setSessionState(this.editSession, { spSiteFilter: value || '', spCurrentPage: 0 });
  };

  private handleSharePointSearch = async (): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    this.setSessionState(session, { spSitesLoading: true, spCurrentPage: 0 });
    const sites = await this.sharePointService?.searchSites(this.state.spSiteFilter) || [];
    if (!this.isCurrentSession(session)) return;
    this.setSessionState(session, { spSites: sites, spSitesLoading: false });
  };

  private handleSelectSharePointSite = async (site: ISharePointSite): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    this.setSessionState(session, { spSelectedSite: site, spListsLoading: true, spLists: [] });
    const lists = await this.sharePointService?.getCalendarLists(site.id) || [];
    if (!this.isCurrentSession(session)) return;
    if (this.state.spSelectedSite?.id !== site.id) return;
    this.setSessionState(session, { spLists: lists, spListsLoading: false, addingCalendarStep: 'sharepoint-list' });
  };

  private handleSelectSharePointList = (list: ISharePointList): void => {
    this.setSessionState(this.editSession, {
      spSelectedList: list,
      spAvailableFields: [], spFieldMapping: {},
      newCalendarName: list.name,
      newCalendarColor: this.state.settings.organizationPrimaryColor || '#0078d4'
    }, () => {
      this.fetchSharePointListFields(list).catch(err => console.error('Failed to fetch SharePoint list fields:', err));
    });
  };

  private async fetchSharePointListFields(list: ISharePointList): Promise<void> {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    const generation = ++this.sharePointFieldsGeneration;
    const { spSelectedSite } = this.state;
    if (!spSelectedSite || !this.props.graphClient) {
      return;
    }

    try {
      const columnsData = await this.props.graphClient
        .api(`/sites/${spSelectedSite.id}/lists/${list.id}/columns`)
        .query({ $select: 'name,displayName,columnGroup' })
        .get();

      if (!this.isCurrentSession(session) || generation !== this.sharePointFieldsGeneration || this.state.spSelectedList?.id !== list.id) return;
      const rawOptions: IDropdownOption[] = (columnsData.value || [])
        .filter((column: IGraphColumn) => column.name && !column.name.startsWith('_') && column.columnGroup !== '_Hidden')
        .map((column: IGraphColumn) => ({
          key: column.name as string,
          text: column.displayName || (column.name as string)
        }));

      this.setSessionState(session, {
        spAvailableFields: rawOptions,
        addingCalendarStep: 'sharepoint-fields',
        spFieldMapping: {
          titleField: findBestMatchingFieldKey(rawOptions, getFieldCandidates('title')) as string | undefined,
          startDateField: findBestMatchingFieldKey(rawOptions, getFieldCandidates('start')) as string | undefined,
          endDateField: findBestMatchingFieldKey(rawOptions, getFieldCandidates('end')) as string | undefined
        }
      });
    } catch (error) {
      if (!this.isCurrentSession(session) || generation !== this.sharePointFieldsGeneration || this.state.spSelectedList?.id !== list.id) return;
      console.error('Failed to load SharePoint field metadata:', error);
      this.setSessionState(session, { addingCalendarStep: 'sharepoint-fields' });
    }
  }

  private handleExchangeMailboxChange = (value?: string): void => {
    this.mailboxDiscovery.invalidate();
    this.setSessionState(this.editSession, {
      exchangeMailbox: value || '', exchangeMailboxPerson: undefined, exchangeCalendars: [], exchangeCalendarsLoading: false,
      exchangeMailboxResolved: false, exchangeSelectedCalendarId: undefined, exchangeDiscoveryError: undefined
    });
  };

  private handleExchangeMailboxPersonChange = (mailbox?: IMailboxSearchResult): void => {
    if (!mailbox) {
      this.handleExchangeMailboxChange('');
      return;
    }

    this.mailboxDiscovery.invalidate();
    this.setSessionState(this.editSession, {
      exchangeMailbox: mailbox.userPrincipalName || mailbox.mail || mailbox.id,
      exchangeMailboxPerson: mailbox,
      exchangeCalendars: [], exchangeCalendarsLoading: false,
      exchangeMailboxResolved: false, exchangeSelectedCalendarId: undefined, exchangeDiscoveryError: undefined
    });
  };

  private handleExchangeMailboxInputChange = (value: string): void => {
    // NormalPeoplePicker clears its input after selecting a persona. Keep the selected
    // mailbox in that case; actual typed input still replaces the current selection.
    if (!value && this.state.exchangeMailboxPerson) return;
    this.handleExchangeMailboxChange(value);
  };

  private resolveExchangeMailboxSuggestions = async (query: string): Promise<IMailboxSearchResult[]> => {
    const session = this.editSession;
    const generation = ++this.mailboxSuggestionGeneration;
    if (!this.exchangeService || query.trim().length < 2) return [];

    try {
      const results = await this.exchangeService.searchMailboxes(query, 5, this.props.webAbsoluteUrl);
      return this.isCurrentSession(session) && generation === this.mailboxSuggestionGeneration ? results : [];
    } catch (error) {
      console.error('Exchange mailbox search failed:', error);
      return [];
    }
  };

  private handleExchangeLookupMailbox = async (): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    const mailbox = this.state.exchangeMailbox.trim();
    await this.mailboxDiscovery.run(async () => {
      if (!this.exchangeService) throw new Error('GraphClient not initialized');
      const resolved = await this.exchangeService.resolveMailbox(mailbox);
      const calendars = await this.exchangeService.getCalendars(resolved.id);
      return { calendars, mailboxId: resolved.id };
    }, {
      start: () => this.setSessionState(session, {
        exchangeMailbox: mailbox, exchangeCalendarsLoading: true, exchangeCalendars: [],
        exchangeMailboxResolved: false, exchangeSelectedCalendarId: undefined, exchangeDiscoveryError: undefined
      }),
      success: result => this.setSessionState(session, {
        exchangeCalendars: result.calendars, resolvedMailboxId: result.mailboxId, selectedExchangeCalendars: {}, exchangeMailboxResolved: true,
        addingCalendarStep: 'exchange-calendar'
      }),
      error: error => {
        console.error('Exchange mailbox discovery failed:', error);
        this.setSessionState(session, { exchangeDiscoveryError: getExchangeDiscoveryErrorMessage(error) });
      },
      finish: () => this.setSessionState(session, { exchangeCalendarsLoading: false })
    });
    if (!this.isCurrentSession(session)) return;
  };

  private handleSelectExchangeCalendar = (calendar: IExchangeCalendar): void => {
    this.setSessionState(this.editSession, {
      exchangeSelectedCalendarId: calendar.id,
      newCalendarName: calendar.name,
      newCalendarColor: calendar.hexColor
    });
  };

  private handleSelectPlannerPlan = (planId: string, planTitle: string): void => {
    this.setSessionState(this.editSession, {
      plannerSelectedPlanId: planId,
      addingCalendarStep: 'planner-options',
      newCalendarName: planTitle
    });
  };

  private handleToggleUnifiedGroupSelection = (groupId: string, checked?: boolean): void => {
    this.setSessionState(this.editSession, prev => ({
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
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    if (this.state.audiencePreset && !target) {
      this.addSourcesToAudience(sources);
      return;
    }
    const selectedAudienceGroups: Record<string, string> = {};
    if (target) {
      const existing = this.state.settings.assignedSources.find(item => (item.assignmentId || item.adminSourceId) === target.id);
      existing?.audienceGroups.forEach(group => {
        selectedAudienceGroups[group.groupId] = group.displayName;
      });
    }

    this.setSessionState(session, {
      audienceFirst: false, audiencePreset: false, everyone: !!target && Object.keys(selectedAudienceGroups).length === 0,
      pendingAdminSources: sources,
      pendingAdminIcs: undefined,
      audienceEditTarget: target,
      selectedAudienceGroups,
      addingCalendarStep: 'admin-audience-select',
      securityGroupsLoading: true
    });

    await this.loadSecurityGroups();
    if (!this.isCurrentSession(session)) return;
  };

  private beginAudienceSelectionForIcs = async (
    item: { adminIcsId?: string; displayName: string; icsUrl: string },
    target?: { kind: 'ics'; id: string }
  ): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    const selectedAudienceGroups: Record<string, string> = {};
    if (target) {
      const existing = this.state.settings.icsCatalog.find(entry => entry.adminIcsId === target.id);
      existing?.audienceGroups.forEach(group => {
        selectedAudienceGroups[group.groupId] = group.displayName;
      });
    }

    this.setSessionState(session, {
      audienceFirst: false, audiencePreset: false, everyone: !!target && Object.keys(selectedAudienceGroups).length === 0,
      pendingAdminSources: [],
      pendingAdminIcs: item,
      audienceEditTarget: target,
      selectedAudienceGroups,
      addingCalendarStep: 'admin-audience-select',
      securityGroupsLoading: true
    });

    await this.loadSecurityGroups();
    if (!this.isCurrentSession(session)) return;
  };

  private createSelectedAudienceGroups(): IAudienceGroup[] {
    return Object.keys(this.state.selectedAudienceGroups).map(groupId => ({
      groupId,
      displayName: this.state.selectedAudienceGroups[groupId],
      groupType: this.state.selectedAudienceMetadata[groupId]?.groupType || this.state.settings.audienceGroups?.find(group => group.groupId === groupId)?.groupType
    }));
  }

  private handleConfirmSharePointCalendar = async (): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    if (!this.state.spSelectedSite || !this.state.spSelectedList) {
      return;
    }

    const source: ICalendarSourceBase = {
      sourceType: 'sharepoint',
      name: this.state.newCalendarName,
      color: this.state.newCalendarColor,
      isEnabled: true,
      sharePointSiteId: this.state.spSelectedSite.id,
      sharePointSiteName: this.state.spSelectedSite.name,
      sharePointListId: this.state.spSelectedList.id,
      sharePointFieldMapping: this.state.spFieldMapping
    };

    const pending = [...this.state.pendingSharePointSources, { source, policy: this.state.pendingPolicy }];
    const remaining = this.state.selectedSharePointLists.filter(id => id !== this.state.spSelectedList?.id);
    if (remaining.length) {
      this.setSessionState(session, { pendingSharePointSources: pending, selectedSharePointLists: remaining, spFieldMapping: {}, spAvailableFields: [], pendingPolicy: { isMandatory: false, defaultEnabled: true } }, () => {
        const next = this.state.spLists.find(item => item.id === remaining[0]); if (next) this.handleSelectSharePointList(next);
      });
      return;
    }
    const settings = addAdministratorAssignments(this.state.settings, pending, this.createSelectedAudienceGroups());
    this.setSessionState(session, { settings, pendingSharePointSources: [], selectedSharePointLists: [] }, () => this.handleCloseAddDialog());
  };

  private handleConfirmExchangeCalendar = async (): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    const mailboxId = this.state.resolvedMailboxId;
    if (!mailboxId) return;
    if (this.state.exchangeAll) {
      const rule: IExchangeMailboxAssignment = { assignmentId: generateStableId('mailboxAssignment'), mailboxId,
        mailboxDisplayName: this.state.exchangeMailbox, audienceGroups: this.createSelectedAudienceGroups(),
        ...this.state.exchangeAllPolicy, allowedOverrides: { ...defaultAllowedOverrides }, exceptions: [] };
      const settings = addExchangeMailboxAssignment(this.state.settings, rule);
      this.setSessionState(session, { settings,
        mailboxRuleCalendars: { ...this.state.mailboxRuleCalendars, [mailboxId]: this.state.exchangeCalendars.map(calendar => ({ sourceType: 'exchange', name: calendar.name, color: calendar.hexColor, isEnabled: true, exchangeMailbox: mailboxId, exchangeCalendarId: calendar.id })) } }, () => this.handleCloseAddDialog());
      return;
    }
    const sources: ICalendarSourceBase[] = this.state.exchangeCalendars.filter(calendar => this.state.selectedExchangeCalendars[calendar.id]).map(calendar => ({
      sourceType: 'exchange', name: calendar.name, color: calendar.hexColor, isEnabled: true, exchangeMailbox: mailboxId, exchangeCalendarId: calendar.id
    }));
    this.addSourcesToAudience(sources, this.state.selectedExchangeCalendars);
  };

  private handleConfirmPlannerPlan = async (): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
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
    if (!this.isCurrentSession(session)) return;
  };

  private handleConfirmTeamsShifts = async (): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
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
    if (!this.isCurrentSession(session)) return;
  };

  private handleConfirmUnifiedGroups = async (): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
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
    if (!this.isCurrentSession(session)) return;
  };

  private handleConfirmAdminIcsItem = async (): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    if (!this.state.icsUrl.trim() || !this.state.newCalendarName.trim()) {
      return;
    }

    await this.beginAudienceSelectionForIcs({
      displayName: this.state.newCalendarName.trim(),
      icsUrl: this.state.icsUrl.trim()
    });
    if (!this.isCurrentSession(session)) return;
  };

  private loadSecurityGroups = async (searchText?: string): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    await this.audienceDiscovery.run(async () => {
      if (!this.audienceService) throw new Error('GraphClient not initialized');
      return this.audienceService.getAudienceGroups(searchText);
    }, {
      start: () => this.setSessionState(session, prev => updateAudienceDiscovery(prev, { type: 'start' })),
      success: securityGroups => this.setSessionState(session, prev => {
        const known = new Map(securityGroups.map(group => [group.id, group]));
        const enrich = (group: IAudienceGroup): IAudienceGroup => { const found = known.get(group.groupId); return found ? { ...group, displayName: found.displayName, groupType: found.groupType } : group; };
        const selectedAudienceMetadata = { ...prev.selectedAudienceMetadata };
        for (const id of Object.keys(prev.selectedAudienceGroups)) {
          const found = known.get(id); if (found) selectedAudienceMetadata[id] = { groupId: id, displayName: found.displayName, groupType: found.groupType };
        }
        return { ...prev, ...updateAudienceDiscovery(prev, { type: 'success', groups: securityGroups }), selectedAudienceMetadata,
          settings: { ...prev.settings, audienceGroups: (prev.settings.audienceGroups || []).map(enrich),
            assignedSources: prev.settings.assignedSources.map(item => ({ ...item, audienceGroups: item.audienceGroups.map(enrich) })),
            exchangeMailboxAssignments: (prev.settings.exchangeMailboxAssignments || []).map(item => ({ ...item, audienceGroups: item.audienceGroups.map(enrich) })),
            icsCatalog: prev.settings.icsCatalog.map(item => ({ ...item, audienceGroups: item.audienceGroups.map(enrich) })) } };
      }),
      error: error => {
        console.error('Audience discovery failed:', error);
        this.setSessionState(session, prev => updateAudienceDiscovery(prev, { type: 'error', message: strings.SecurityGroupsLoadErrorLabel }));
      },
      finish: () => this.setSessionState(session, prev => updateAudienceDiscovery(prev, { type: 'finish' }))
    });
    if (!this.isCurrentSession(session)) return;
  };

  private handleSecurityGroupSearch = async (): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    await this.loadSecurityGroups(this.state.securityGroupSearch);
    if (!this.isCurrentSession(session)) return;
  };

  private handleSecurityGroupSearchChange = (value?: string): void => {
    this.audienceDiscovery.invalidate();
    this.setSessionState(this.editSession, prev => ({ ...updateAudienceDiscovery(prev, { type: 'input' }), securityGroupSearch: value || '' }));
  };

  private handleToggleAudienceGroup = (group: IEntraSecurityGroup, checked?: boolean): void => {
    this.setSessionState(this.editSession, prev => {
      const selectedAudienceGroups = { ...prev.selectedAudienceGroups };
      if (checked) {
        selectedAudienceGroups[group.id] = group.displayName;
      } else {
        delete selectedAudienceGroups[group.id];
      }
      return { selectedAudienceGroups: prev.audienceFirst && checked ? { [group.id]: group.displayName } : selectedAudienceGroups,
        everyone: false, selectedAudienceMetadata: { ...prev.selectedAudienceMetadata, [group.id]: { groupId: group.id, displayName: group.displayName, groupType: group.groupType } } };
    });
  };

  private handleApplyAudienceSelection = (): void => {
    const audienceGroups = this.createSelectedAudienceGroups();
    if (audienceGroups.length === 0 && !this.state.everyone) return;
    if (this.state.audienceFirst && this.state.addingCalendarType) {
      this.setSessionState(this.editSession, { audienceFirst: false, audiencePreset: true }, () => { this.handleSelectSourceType(this.state.addingCalendarType as CalendarSourceType).catch(error => console.error(error)); });
      return;
    }

    if (this.state.audienceEditTarget?.kind === 'source') {
      const settings = {
        ...this.state.settings,
        assignedSources: this.state.settings.assignedSources.map(item =>
          (item.assignmentId || item.adminSourceId) === this.state.audienceEditTarget?.id
            ? { ...item, audienceGroups }
            : item
        )
      };
      this.setSessionState(this.editSession, { settings }, () => this.handleCloseAddDialog());
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
      this.setSessionState(this.editSession, { settings }, () => this.handleCloseAddDialog());
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
      this.setSessionState(this.editSession, { settings }, () => this.handleCloseAddDialog());
      return;
    }

    this.addSourcesToAudience(this.state.pendingAdminSources);
  };

  private addSourcesToAudience = (sources: ICalendarSourceBase[], policies?: Record<string, IAssignmentPolicy>): void => {
    const settings = addAdministratorAssignments(this.state.settings, sources.map(source => ({ source,
      policy: policies?.[source.exchangeCalendarId || getSourceIdentityKey(source)] || this.state.pendingPolicy })), this.createSelectedAudienceGroups());
    this.setSessionState(this.editSession, { settings }, () => this.handleCloseAddDialog());
  };

  private handleUpdateAssignedSource = (adminSourceId: string, updates: Partial<ICalendarSourceBase>): void => {
    this.setSessionState(this.editSession, prev => ({
      settings: {
        ...prev.settings,
        sourceCatalog: (prev.settings.sourceCatalog || []).map(item => item.adminSourceId === adminSourceId ? { ...item, source: { ...item.source, ...updates } } : item),
        assignedSources: prev.settings.assignedSources.map(item =>
          item.adminSourceId === adminSourceId
            ? { ...item, source: { ...item.source, ...updates } }
            : item
        )
      }
    }));
  };

  private handleDeleteAssignedSource = (adminSourceId: string): void => {
    this.setSessionState(this.editSession, prev => ({
      settings: {
        ...prev.settings,
        assignedSources: prev.settings.assignedSources.filter(item => (item.assignmentId || item.adminSourceId) !== adminSourceId)
      },
      editingSourceId: prev.editingSourceId === adminSourceId ? undefined : prev.editingSourceId
    }));
  };

  private handleEditAssignedSourceAudiences = async (adminSourceId: string): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    const existing = this.state.settings.assignedSources.find(item => (item.assignmentId || item.adminSourceId) === adminSourceId);
    if (!existing) {
      return;
    }

    await this.beginAudienceSelectionForSources([existing.source], { kind: 'source', id: adminSourceId });
    if (!this.isCurrentSession(session)) return;
    this.setSessionState(session, { showAddDialog: true });
  };

  private handleUpdateIcsCatalogItem = (adminIcsId: string, updates: Partial<IAdminIcsCatalogItem>): void => {
    this.setSessionState(this.editSession, prev => ({
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
    this.setSessionState(this.editSession, prev => ({
      settings: {
        ...prev.settings,
        icsCatalog: prev.settings.icsCatalog.filter(item => item.adminIcsId !== adminIcsId)
      },
      editingIcsId: prev.editingIcsId === adminIcsId ? undefined : prev.editingIcsId
    }));
  };

  private handleEditIcsAudiences = async (adminIcsId: string): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    const item = this.state.settings.icsCatalog.find(entry => entry.adminIcsId === adminIcsId);
    if (!item) {
      return;
    }

    await this.beginAudienceSelectionForIcs({
      adminIcsId: item.adminIcsId,
      displayName: item.displayName,
      icsUrl: item.icsUrl
    }, { kind: 'ics', id: adminIcsId });
    if (!this.isCurrentSession(session)) return;
    this.setSessionState(session, { showAddDialog: true });
  };

  private toggleEditSource = (id: string | undefined): void => {
    this.setSessionState(this.editSession, { editingSourceId: id });
  };

  private toggleEditIcs = (id: string | undefined): void => {
    this.setSessionState(this.editSession, { editingIcsId: id });
  };

  private renderSharePointFlow(): React.ReactElement {
    const { spSites, spSitesLoading, spSiteFilter, spSelectedSite, spLists, spListsLoading, spSelectedList, spAvailableFields, spFieldMapping, addingCalendarStep } = this.state;

    if (spSelectedList && addingCalendarStep === 'sharepoint-fields') {
      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>{strings.FieldMappingLabel}</Label>
          <Dropdown label={strings.TitleSubjectFieldLabel} options={spAvailableFields} selectedKey={spFieldMapping.titleField || ''} onChange={(_, option) => this.setSessionState(this.editSession, { spFieldMapping: { ...spFieldMapping, titleField: option?.key as string } })} />
          <Dropdown label={strings.StartDateFieldLabel} options={spAvailableFields} selectedKey={spFieldMapping.startDateField || ''} onChange={(_, option) => this.setSessionState(this.editSession, { spFieldMapping: { ...spFieldMapping, startDateField: option?.key as string } })} />
          <Dropdown label={strings.EndDateFieldLabel} options={spAvailableFields} selectedKey={spFieldMapping.endDateField || ''} onChange={(_, option) => this.setSessionState(this.editSession, { spFieldMapping: { ...spFieldMapping, endDateField: option?.key as string } })} />
          <Dropdown label={strings.LocationFieldOptionalLabel} options={[{ key: '', text: strings.NoneLabel }, ...spAvailableFields]} selectedKey={spFieldMapping.locationField || ''} onChange={(_, option) => this.setSessionState(this.editSession, { spFieldMapping: { ...spFieldMapping, locationField: option?.key as string } })} />
          <Dropdown label={strings.DescriptionFieldOptionalLabel} options={[{ key: '', text: strings.NoneLabel }, ...spAvailableFields]} selectedKey={spFieldMapping.descriptionField || ''} onChange={(_, option) => this.setSessionState(this.editSession, { spFieldMapping: { ...spFieldMapping, descriptionField: option?.key as string } })} />
          <AssignmentPolicyControl policy={this.state.pendingPolicy} onChange={pendingPolicy => this.setSessionState(this.editSession, { pendingPolicy })} />
          <TextField label={strings.CalendarNameLabel} value={this.state.newCalendarName} onChange={(_, value) => this.setSessionState(this.editSession, { newCalendarName: value || '' })} />
          <div>
            <Label>{strings.ColorLabel}</Label>
            <ColorPicker color={this.state.newCalendarColor} onChange={(_, color) => this.setSessionState(this.editSession, { newCalendarColor: `#${color.hex}` })} alphaType="none" />
          </div>
          <PrimaryButton text={strings.AddItemLabel} onClick={() => this.handleConfirmSharePointCalendar().catch(err => console.error(err))} />
        </Stack>
      );
    }

    if (spSelectedSite) {
      if (spListsLoading) {
        return <Spinner size={SpinnerSize.medium} label={strings.LoadingLabel} />;
      }

      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>{formatLocalizedString(strings.SelectCalendarListLabel, spSelectedSite.name)}</Label>
          <Stack tokens={{ childrenGap: 8 }}>
            {spLists.map(list => (
              <Checkbox key={list.id} label={list.name} checked={this.state.selectedSharePointLists.indexOf(list.id) >= 0} onChange={(_, checked) => this.setSessionState(this.editSession, prev => ({ selectedSharePointLists: checked ? [...prev.selectedSharePointLists, list.id] : prev.selectedSharePointLists.filter(id => id !== list.id) }))} />
            ))}
          </Stack>
          <PrimaryButton text={strings.NextLabel} disabled={!this.state.selectedSharePointLists.length} onClick={() => { const list = spLists.find(item => item.id === this.state.selectedSharePointLists[0]); if (list) this.handleSelectSharePointList(list); }} />
        </Stack>
      );
    }

    if (spSitesLoading) {
      return <Spinner size={SpinnerSize.medium} label={strings.LoadingLabel} />;
    }

    const totalPages = Math.ceil(spSites.length / this.SITES_PER_PAGE);
    const startIndex = this.state.spCurrentPage * this.SITES_PER_PAGE;
    const sitesOnPage = spSites.slice(startIndex, startIndex + this.SITES_PER_PAGE);

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>{strings.SearchSelectSharePointSiteLabel}</Label>
        <TextField placeholder={strings.FilterSitesPlaceholder} value={spSiteFilter} onChange={(_, value) => this.handleSharePointFilterChange(value)} />
        <PrimaryButton text={strings.SearchLabel} onClick={() => this.handleSharePointSearch().catch(err => console.error(err))} />
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
            <DefaultButton text={strings.PreviousLabel} disabled={this.state.spCurrentPage === 0} onClick={() => this.setSessionState(this.editSession, { spCurrentPage: this.state.spCurrentPage - 1 })} />
            <Label>{formatLocalizedString(strings.PageOfLabel, this.state.spCurrentPage + 1, totalPages)}</Label>
            <DefaultButton text={strings.NextLabel} disabled={this.state.spCurrentPage >= totalPages - 1} onClick={() => this.setSessionState(this.editSession, { spCurrentPage: this.state.spCurrentPage + 1 })} />
          </Stack>
        )}
      </Stack>
    );
  }

  private renderExchangeFlow(): React.ReactElement {
    const { exchangeMailbox, exchangeMailboxResolved, exchangeCalendars, exchangeCalendarsLoading } = this.state;

    if (exchangeCalendarsLoading) return <Spinner size={SpinnerSize.medium} label={strings.LoadingLabel} />;
    if (exchangeMailboxResolved) return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>{formatLocalizedString(strings.SelectCalendarFromMailboxLabel, exchangeMailbox)}</Label>
        <Checkbox label={strings.AllMailboxCalendarsLabel} checked={this.state.exchangeAll} onChange={(_, checked) => this.setSessionState(this.editSession, { exchangeAll: !!checked })} />
        {this.state.exchangeAll && <AssignmentPolicyControl policy={this.state.exchangeAllPolicy} onChange={exchangeAllPolicy => this.setSessionState(this.editSession, { exchangeAllPolicy })} />}
        {!this.state.exchangeAll && exchangeCalendars.map(calendar => <Stack key={calendar.id} tokens={{ childrenGap: 4 }}>
          <Checkbox label={calendar.name} checked={!!this.state.selectedExchangeCalendars[calendar.id]} onChange={(_, checked) => this.setSessionState(this.editSession, prev => {
            const selectedExchangeCalendars = { ...prev.selectedExchangeCalendars };
            if (checked) selectedExchangeCalendars[calendar.id] = { isMandatory: false, defaultEnabled: true }; else delete selectedExchangeCalendars[calendar.id];
            return { selectedExchangeCalendars };
          })} />
          {this.state.selectedExchangeCalendars[calendar.id] && <AssignmentPolicyControl policy={this.state.selectedExchangeCalendars[calendar.id]} onChange={policy => this.setSessionState(this.editSession, prev => ({ selectedExchangeCalendars: { ...prev.selectedExchangeCalendars, [calendar.id]: policy } }))} />}
        </Stack>)}
        <PrimaryButton text={strings.AddItemLabel} disabled={!this.state.exchangeAll && !Object.keys(this.state.selectedExchangeCalendars).length} onClick={() => this.handleConfirmExchangeCalendar().catch(error => console.error(error))} />
      </Stack>
    );

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>{strings.EnterMailboxEmailLabel}</Label>
        <MailboxPeoplePicker
          selectedMailbox={this.state.exchangeMailboxPerson}
          placeholder={strings.MailboxPlaceholder}
          disabled={exchangeCalendarsLoading}
          onResolveSuggestions={this.resolveExchangeMailboxSuggestions}
          onInputChange={this.handleExchangeMailboxInputChange}
          onChange={this.handleExchangeMailboxPersonChange}
        />
        <div style={{ fontSize: 12, color: '#605e5c' }}>{strings.MailboxInputHelpLabel}</div>
        <PrimaryButton text={strings.LoadCalendarsLabel} disabled={!exchangeMailbox.trim() || exchangeCalendarsLoading} onClick={() => this.handleExchangeLookupMailbox().catch(err => console.error(err))} />
      </Stack>
    );
  }

  private renderIcsFlow(): React.ReactElement {
    const hasValidInput = this.state.icsUrl.trim() && this.state.newCalendarName.trim();

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <TextField label={strings.DisplayNameLabel} value={this.state.newCalendarName} onChange={(_, value) => this.setSessionState(this.editSession, { newCalendarName: value || '' })} placeholder={strings.IcsNamePlaceholder} />
        <TextField label={strings.IcsUrlLabel} value={this.state.icsUrl} onChange={(_, value) => this.setSessionState(this.editSession, { icsUrl: value || '' })} placeholder={strings.IcsUrlPlaceholder} />
        <PrimaryButton text={strings.NextChooseGroupsLabel} onClick={() => this.handleConfirmAdminIcsItem().catch(err => console.error(err))} disabled={!hasValidInput} />
      </Stack>
    );
  }

  private renderPlannerFlow(): React.ReactElement {
    const { addingCalendarStep, plannerPlans, plannerPlansLoading } = this.state;

    if (addingCalendarStep === 'planner-plan') {
      if (plannerPlansLoading) {
        return <Spinner size={SpinnerSize.large} label={strings.LoadingPlannerPlansLabel} />;
      }

      return (
        <Stack tokens={{ childrenGap: 12 }}>
          <Label>{strings.SelectPlannerPlanLabel}</Label>
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
        <TextField label={strings.CalendarNameLabel} value={this.state.newCalendarName} onChange={(_, value) => this.setSessionState(this.editSession, { newCalendarName: value || '' })} required />
        <Toggle label={strings.AssignedToMeOnlyLabel} checked={this.state.plannerAssignedToMeOnly} onText={strings.OnLabel} offText={strings.OffLabel} onChange={(_, checked) => this.setSessionState(this.editSession, { plannerAssignedToMeOnly: !!checked })} />
        <Toggle label={strings.ShowCompletedTasksLabel} checked={this.state.plannerShowCompleted} onText={strings.OnLabel} offText={strings.OffLabel} onChange={(_, checked) => this.setSessionState(this.editSession, { plannerShowCompleted: !!checked })} />
        <Toggle label={strings.SourceLogoLabel} checked={this.state.plannerShowLogo} onText={strings.OnLabel} offText={strings.OffLabel} onChange={(_, checked) => this.setSessionState(this.editSession, { plannerShowLogo: !!checked })} />
        <ColorPicker color={this.state.newCalendarColor} onChange={(_, color) => this.setSessionState(this.editSession, { newCalendarColor: `#${color.hex}` })} alphaType="none" />
        <PrimaryButton text={strings.NextChooseGroupsLabel} onClick={() => this.handleConfirmPlannerPlan().catch(err => console.error(err))} disabled={!this.state.newCalendarName.trim()} />
      </Stack>
    );
  }

  private renderUnifiedGroupsFlow(): React.ReactElement {
    const { unifiedGroups, unifiedGroupsLoading, unifiedGroupsSelection } = this.state;
    const selectedCount = Object.keys(unifiedGroupsSelection).filter(id => unifiedGroupsSelection[id]).length;

    if (unifiedGroupsLoading) {
      return <Spinner size={SpinnerSize.large} label={strings.LoadingGroupsAndTeamsLabel} />;
    }

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>{strings.SelectGroupsOrTeamsLabel}</Label>
        <Stack tokens={{ childrenGap: 8 }}>
          {unifiedGroups.map(group => (
            <Stack key={group.id} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ padding: '8px 12px', border: '1px solid #edebe9', borderRadius: 4, backgroundColor: unifiedGroupsSelection[group.id] ? '#f3f2f1' : 'white' }}>
              <Icon iconName={group.isTeam ? 'TeamsLogo' : 'Group'} style={{ fontSize: 16 }} />
              <Checkbox label={group.displayName} checked={!!unifiedGroupsSelection[group.id]} onChange={(_, checked) => this.handleToggleUnifiedGroupSelection(group.id, checked)} />
            </Stack>
          ))}
        </Stack>
        <div>
          <Label>{strings.ColorLabel}</Label>
          <ColorPicker color={this.state.newCalendarColor} onChange={(_, color) => this.setSessionState(this.editSession, { newCalendarColor: `#${color.hex}` })} alphaType="none" />
        </div>
        <PrimaryButton text={selectedCount > 1 ? strings.NextChooseGroupsForCalendarsLabel : strings.NextChooseGroupsForCalendarLabel} onClick={() => this.handleConfirmUnifiedGroups().catch(err => console.error(err))} disabled={selectedCount === 0} />
      </Stack>
    );
  }

  private renderTeamsShiftsFlow(): React.ReactElement {
    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <TextField label={strings.CalendarNameLabel} value={this.state.newCalendarName} onChange={(_, value) => this.setSessionState(this.editSession, { newCalendarName: value || '' })} placeholder={strings.TeamsShiftsLabel} required />
        <Toggle label={strings.SourceLogoLabel} checked={this.state.teamsShiftsShowLogo} onText={strings.OnLabel} offText={strings.OffLabel} onChange={(_, checked) => this.setSessionState(this.editSession, { teamsShiftsShowLogo: !!checked })} />
        <ColorPicker color={this.state.newCalendarColor} onChange={(_, color) => this.setSessionState(this.editSession, { newCalendarColor: `#${color.hex}` })} alphaType="none" showPreview={true} />
        <PrimaryButton text={strings.NextChooseGroupsLabel} onClick={() => this.handleConfirmTeamsShifts().catch(err => console.error(err))} disabled={!this.state.newCalendarName.trim()} />
      </Stack>
    );
  }

  private renderAudienceSelectionFlow(): React.ReactElement {
    const selectedCount = Object.keys(this.state.selectedAudienceGroups).length;
    const selectedGroupNames = Object.keys(this.state.selectedAudienceGroups).map(groupId => this.state.selectedAudienceGroups[groupId]);

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Label>{strings.SelectAdminAudienceGroupsLabel}</Label>
        <MessageBar>{strings.AudiencePermissionsHelpLabel}</MessageBar>
        <Checkbox label={strings.EveryoneAudienceLabel} checked={!!this.state.everyone} onChange={(_, checked) => this.setSessionState(this.editSession, { everyone: !!checked, selectedAudienceGroups: {} })} />
        <TextField placeholder={strings.SearchSecurityGroupsPlaceholder} value={this.state.securityGroupSearch} onChange={(_, value) => this.handleSecurityGroupSearchChange(value)} />
        <PrimaryButton text={strings.SearchLabel} onClick={() => this.handleSecurityGroupSearch().catch(err => console.error(err))} />
        {selectedCount > 0 && (
          <MessageBar messageBarType={MessageBarType.info}>
            {formatLocalizedString(strings.SelectedGroupsLabel, selectedGroupNames.join(', '))}
          </MessageBar>
        )}
        {this.state.securityGroupsError && <MessageBar messageBarType={MessageBarType.error}>{this.state.securityGroupsError}</MessageBar>}
        {this.state.securityGroupsLoaded && !this.state.securityGroupsLoading && !this.state.securityGroupsError && this.state.securityGroups.length === 0 && <MessageBar>{strings.NoSecurityGroupsLabel}</MessageBar>}
        {this.state.securityGroupsLoading ? (
          <Spinner size={SpinnerSize.medium} label={strings.LoadingSecurityGroupsLabel} />
        ) : (
          <Stack tokens={{ childrenGap: 8 }}>
            {this.state.securityGroups.map(group => (
              <Checkbox key={group.id} label={group.displayName + (group.groupType ? ' — ' + audienceTypeLabel({ groupId: group.id, displayName: group.displayName, groupType: group.groupType }) : '')} checked={!!this.state.selectedAudienceGroups[group.id]} onChange={(_, checked) => this.handleToggleAudienceGroup(group, checked)} />
            ))}
          </Stack>
        )}
        <PrimaryButton text={this.state.audienceFirst ? strings.NextLabel : this.state.audienceEditTarget ? strings.ApplyGroupsLabel : strings.AddItemLabel} onClick={this.handleApplyAudienceSelection} disabled={selectedCount === 0 && !this.state.everyone} />
      </Stack>
    );
  }

  private renderAddFlow(): React.ReactElement {
    const { addingCalendarStep, addingCalendarType } = this.state;

    if (addingCalendarStep === 'initial') {
      return (
        <Stack tokens={{ childrenGap: 16 }}>
          <Label>{strings.SelectAdminItemTypeLabel}</Label>
          <Stack tokens={{ childrenGap: 12 }}>
            {calendarSourceRegistry.filter(definition => definition.adminSelectable).map(definition => (
              <PrimaryButton
                key={definition.type}
                text={strings[definition.displayNameKey]}
                secondaryText={definition.adminCatalogOnly ? strings.PublishIcsCatalogItemDescription : getSourceTypeDescription(definition.type)}
                iconProps={{ iconName: definition.iconName }}
                onClick={() => this.handleSelectAddType(definition.type).catch(err => console.error(err))}
                style={{ textAlign: 'left', height: 'auto', padding: '12px' }}
              />
            ))}
          </Stack>
          <DefaultButton text={strings.CancelLabel} onClick={this.handleCloseAddDialog} />
        </Stack>
      );
    }

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        {this.renderNavigationHeader()}
        {addingCalendarStep === 'admin-audience-select' && this.renderAudienceSelectionFlow()}
        {addingCalendarType === 'sharepoint' && addingCalendarStep !== 'admin-audience-select' && this.renderSharePointFlow()}
        {addingCalendarType === 'exchange' && addingCalendarStep !== 'admin-audience-select' && this.renderExchangeFlow()}
        {addingCalendarType === 'exchange' && addingCalendarStep !== 'admin-audience-select' && this.state.exchangeDiscoveryError && <MessageBar messageBarType={MessageBarType.error}>{this.state.exchangeDiscoveryError}</MessageBar>}
        {addingCalendarType === 'exchange' && addingCalendarStep !== 'admin-audience-select' && this.state.exchangeMailboxResolved && !this.state.exchangeCalendarsLoading && this.state.exchangeCalendars.length === 0 && <MessageBar>{strings.NoMailboxCalendarsLabel}</MessageBar>}
        {addingCalendarType === 'planner' && addingCalendarStep !== 'admin-audience-select' && this.renderPlannerFlow()}
        {addingCalendarType === 'unifiedGroup' && addingCalendarStep !== 'admin-audience-select' && this.renderUnifiedGroupsFlow()}
        {addingCalendarType === 'teamsShifts' && addingCalendarStep !== 'admin-audience-select' && this.renderTeamsShiftsFlow()}
        {addingCalendarType === 'ics' && addingCalendarStep !== 'admin-audience-select' && this.renderIcsFlow()}
      </Stack>
    );
  }

  private handleUpdatePolicy = (id: string, policy: IAssignmentPolicy): void => {
    this.setSessionState(this.editSession, prev => ({ settings: { ...prev.settings, assignedSources: prev.settings.assignedSources.map(item => (item.assignmentId || item.adminSourceId) === id ? { ...item, ...policy } : item) } }));
  };

  private handleUpdateAllowedOverrides = (id: string, allowedOverrides: IAdminAllowedOverrides): void => {
    this.setSessionState(this.editSession, prev => ({ settings: { ...prev.settings,
      sourceCatalog: (prev.settings.sourceCatalog || []).map(item => item.adminSourceId === id ? { ...item, allowedOverrides } : item),
      assignedSources: prev.settings.assignedSources.map(item => item.adminSourceId === id ? { ...item, allowedOverrides } : item) } }));
  };

  private handleUpdateMailboxRule = (id: string, updates: Partial<IExchangeMailboxAssignment>): void => {
    this.setSessionState(this.editSession, prev => ({ settings: { ...prev.settings, exchangeMailboxAssignments: (prev.settings.exchangeMailboxAssignments || []).map(rule => rule.assignmentId === id ? { ...rule, ...updates } : rule) } }));
  };

  private canonicalizeMailboxes = async (): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    if (!this.exchangeService) return;
    const generation = this.ruleDiscoveryGeneration;
    const aliases = Array.from(new Set(this.state.settings.assignedSources.filter(item => item.source.sourceType === 'exchange' && item.source.exchangeMailbox).map(item => item.source.exchangeMailbox as string)));
    const identities = await Promise.all(aliases.map(async alias => {
      try { return { alias, id: (await this.exchangeService?.resolveMailbox(alias))?.id }; } catch { return { alias, id: undefined }; }
    }));
    if (!this.isCurrentSession(session)) return;
    if (generation !== this.ruleDiscoveryGeneration) return;
    const byAlias = new Map(identities.filter(item => item.id).map(item => [item.alias.toLowerCase(), item.id as string]));
    this.setSessionState(session, prev => ({ settings: canonicalizeExchangeSourceIdentities(prev.settings, byAlias) }));
  };

  private discoverMailboxRules = async (): Promise<void> => {
    const session = this.editSession;
    if (!this.isCurrentSession(session)) return;
    if (!this.exchangeService) return;
    const generation = ++this.ruleDiscoveryGeneration;
    const ids = Array.from(new Set((this.state.settings.exchangeMailboxAssignments || []).map(rule => rule.mailboxId)));
    const results = await Promise.all(ids.map(async mailboxId => {
      try {
        const calendars = await this.exchangeService?.getCalendars(mailboxId) || [];
        return { mailboxId, sources: calendars.map(calendar => ({ sourceType: 'exchange' as const, name: calendar.name, color: calendar.hexColor, isEnabled: true, exchangeMailbox: mailboxId, exchangeCalendarId: calendar.id })), error: '' };
      } catch (error) { return { mailboxId, sources: [], error: getExchangeDiscoveryErrorMessage(error) }; }
    }));
    if (!this.isCurrentSession(session)) return;
    if (generation !== this.ruleDiscoveryGeneration) return;
    this.setSessionState(session, { mailboxRuleCalendars: results.reduce<Record<string, ICalendarSourceBase[]>>((all, item) => ({ ...all, [item.mailboxId]: item.sources }), {}),
      mailboxRuleErrors: results.reduce<Record<string, string>>((all, item) => ({ ...all, [item.mailboxId]: item.error }), {}) });
  };

  private updateDiscoveredSource = (source: ICalendarSourceBase, allowed: IAdminAllowedOverrides, updates: Partial<ICalendarSourceBase>, allowedUpdates?: IAdminAllowedOverrides): void => {
    this.setSessionState(this.editSession, prev => {
      const catalog = (prev.settings.sourceCatalog || []).slice();
      const existing = catalog.find(item => getSourceIdentityKey(item.source) === getSourceIdentityKey(source));
      const id = existing?.adminSourceId || 'exchange|' + encodeURIComponent((source.exchangeMailbox || '').toLowerCase()) + '|' + encodeURIComponent(source.exchangeCalendarId || '');
      const definition = { adminSourceId: id, source: { ...(existing?.source || source), ...updates }, allowedOverrides: allowedUpdates || existing?.allowedOverrides || allowed };
      const index = catalog.findIndex(item => item.adminSourceId === id);
      if (index >= 0) catalog[index] = definition; else catalog.push(definition);
      return { settings: { ...prev.settings, sourceCatalog: catalog, assignedSources: prev.settings.assignedSources.map(item => item.adminSourceId === id ? { ...item, source: definition.source, allowedOverrides: definition.allowedOverrides } : item) } };
    });
  };

  private renderMailboxRule(rule: IExchangeMailboxAssignment): React.ReactElement {
    return <Stack key={rule.assignmentId} tokens={{ childrenGap: 8 }}>
      <Label>{strings.AllMailboxCalendarsLabel}</Label>
      <AssignmentPolicyControl policy={rule} onChange={policy => this.handleUpdateMailboxRule(rule.assignmentId, policy)} />
      <AllowedOverrideControls value={rule.allowedOverrides} onChange={allowedOverrides => this.setSessionState(this.editSession, prev => ({ settings: { ...prev.settings,
        exchangeMailboxAssignments: (prev.settings.exchangeMailboxAssignments || []).map(item => item.mailboxId === rule.mailboxId ? { ...item, allowedOverrides } : item) } }))} />
      {this.state.mailboxRuleErrors[rule.mailboxId] && <MessageBar messageBarType={MessageBarType.error}>{this.state.mailboxRuleErrors[rule.mailboxId]}</MessageBar>}
      {(this.state.mailboxRuleCalendars[rule.mailboxId] || []).map(source => {
        const exception = rule.exceptions.find(item => item.calendarId === source.exchangeCalendarId);
        const key = exception?.excluded ? 'exclude' : exception?.isMandatory !== undefined ? exception.isMandatory ? 'mandatory' : exception.defaultEnabled ? 'default' : 'available' : 'inherit';
        const definition = this.state.settings.sourceCatalog?.find(item => getSourceIdentityKey(item.source) === getSourceIdentityKey(source));
        const shared = definition?.source || source;
        return <AssignmentGroup key={source.exchangeCalendarId} title={shared.name}>
          <Dropdown label={strings.PolicyLabel} selectedKey={key} options={[{ key: 'inherit', text: strings.InheritPolicyLabel }, { key: 'exclude', text: strings.ExcludeAssignmentLabel }, { key: 'mandatory', text: strings.MandatoryPolicyLabel }, { key: 'default', text: strings.DefaultPolicyLabel }, { key: 'available', text: strings.AvailablePolicyLabel }]}
            onChange={(_, option) => {
              const exceptions = rule.exceptions.filter(item => item.calendarId !== source.exchangeCalendarId);
              if (option?.key !== 'inherit') exceptions.push({ calendarId: source.exchangeCalendarId as string,
                ...(option?.key === 'exclude' ? { excluded: true } : { isMandatory: option?.key === 'mandatory', defaultEnabled: option?.key !== 'available' }) });
              this.handleUpdateMailboxRule(rule.assignmentId, { exceptions });
            }} />
          <MessageBar>{strings.SharedCalendarPropertiesLabel}</MessageBar>
          <TextField label={strings.NameLabel} value={shared.name} onChange={(_, name) => this.updateDiscoveredSource(source, rule.allowedOverrides, { name: name || '' })} />
          <ColorPicker color={shared.color} alphaType="none" onChange={(_, color) => this.updateDiscoveredSource(source, rule.allowedOverrides, { color: '#' + color.hex })} />
          <Toggle label={strings.SourceLogoLabel} checked={shared.showSourceLogo ?? true} onChange={(_, checked) => this.updateDiscoveredSource(source, rule.allowedOverrides, { showSourceLogo: !!checked })} />
          <AllowedOverrideControls value={definition?.allowedOverrides || rule.allowedOverrides} onChange={allowed => this.updateDiscoveredSource(source, rule.allowedOverrides, {}, allowed)} />
        </AssignmentGroup>;
      })}
      <DefaultButton text={strings.DeleteLabel} onClick={() => this.setSessionState(this.editSession, prev => ({ settings: { ...prev.settings, exchangeMailboxAssignments: (prev.settings.exchangeMailboxAssignments || []).filter(item => item.assignmentId !== rule.assignmentId) } }))} />
    </Stack>;
  }

  private renderGroupedAssignments(): React.ReactElement {
    const types: CalendarSourceType[] = ['exchange', 'sharepoint', 'planner', 'unifiedGroup', 'teamsShifts'];
    return <Stack tokens={{ childrenGap: 12 }}>{types.map(type => {
      const items = this.state.settings.assignedSources.filter(item => item.source.sourceType === type);
      const rules = type === 'exchange' ? this.state.settings.exchangeMailboxAssignments || [] : [];
      if (!items.length && !rules.length) return null;
      const groups = new Map<string, IAudienceGroup>();
      [...items, ...rules].forEach(item => item.audienceGroups.length ? item.audienceGroups.forEach(group => groups.set(group.groupId, group)) : groups.set('', { groupId: '', displayName: strings.EveryoneAudienceLabel }));
      return <AssignmentGroup key={type} title={getSourceTypeDisplayName(type)}>{Array.from(groups.values()).map(group => {
        const selectedItems = items.filter(item => group.groupId ? item.audienceGroups.some(entry => entry.groupId === group.groupId) : !item.audienceGroups.length);
        const selectedRules = rules.filter(item => group.groupId ? item.audienceGroups.some(entry => entry.groupId === group.groupId) : !item.audienceGroups.length);
        const containers = new Map<string, { name: string; sources: IAdminAssignedSource[]; rules: IExchangeMailboxAssignment[] }>();
        selectedItems.forEach(item => {
          const key = item.source.exchangeMailbox || item.source.sharePointSiteId || type;
          const container = containers.get(key) || { name: item.source.sharePointSiteName || item.source.exchangeMailbox || getSourceTypeDisplayName(type), sources: [], rules: [] };
          container.sources.push(item); containers.set(key, container);
        });
        selectedRules.forEach(rule => {
          const container = containers.get(rule.mailboxId) || { name: rule.mailboxDisplayName, sources: [], rules: [] };
          container.rules.push(rule); containers.set(rule.mailboxId, container);
        });
        return <AssignmentGroup key={group.groupId} title={group.displayName} description={audienceTypeLabel(group)}>
          {Array.from(containers.entries()).map(([key, container]) => <AssignmentGroup key={key} title={container.name}>
            {container.sources.map((item, index) => this.renderAssignedSource(item, index))}
            {container.rules.map(rule => this.renderMailboxRule(rule))}
          </AssignmentGroup>)}
          <DefaultButton text={strings.AddMoreCalendarsLabel} onClick={() => this.setSessionState(this.editSession, { showAddDialog: true, audienceFirst: false, audiencePreset: true, everyone: !group.groupId,
            selectedAudienceGroups: group.groupId ? { [group.groupId]: group.displayName } : {}, selectedAudienceMetadata: group.groupId ? { [group.groupId]: group } : {},
            selectedExchangeCalendars: {}, exchangeAll: false, pendingPolicy: { isMandatory: false, defaultEnabled: true } }, () => { this.handleSelectSourceType(type).catch(error => console.error(error)); })} />
        </AssignmentGroup>;
      })}</AssignmentGroup>;
    })}</Stack>;
  }

  private renderAssignedSource(item: IAdminAssignedSource, index: number): React.ReactElement {
    const isEditing = this.state.editingSourceId === (item.assignmentId || item.adminSourceId);
    const audienceText = item.audienceGroups.map(group => group.displayName).join(', ');

    return (
      <div key={item.assignmentId || item.adminSourceId} style={{ borderRadius: 4, padding: isEditing ? 12 : '6px 8px', backgroundColor: isEditing ? '#f3f2f1' : (index % 2 === 1 ? 'rgba(0,0,0,0.02)' : 'transparent') }}>
        {isEditing ? (
          <Stack tokens={{ childrenGap: 8 }}>
            <TextField label={strings.NameLabel} value={item.source.name} onChange={(_, value) => this.handleUpdateAssignedSource(item.adminSourceId, { name: value || '' })} />
            {item.source.sourceType === 'sharepoint' && (
              <div style={{ fontSize: 12, color: '#605e5c' }}>
                {strings.SiteLabel}: {item.source.sharePointSiteName || strings.SiteNameUnavailableLabel}
              </div>
            )}
            <div>
              <Label>{strings.ColorLabel}</Label>
              <ColorPicker color={item.source.color} onChange={(_, color) => this.handleUpdateAssignedSource(item.adminSourceId, { color: `#${color.hex}` })} alphaType="none" />
            </div>
            <AssignmentPolicyControl policy={item} onChange={policy => this.handleUpdatePolicy(item.assignmentId || item.adminSourceId, policy)} />
            <MessageBar>{strings.SharedCalendarPropertiesLabel}</MessageBar>
            <Toggle label={strings.SourceLogoLabel} checked={item.source.showSourceLogo ?? true} onChange={(_, checked) => this.handleUpdateAssignedSource(item.adminSourceId, { showSourceLogo: !!checked })} />
            {item.source.sourceType === 'planner' && <Stack tokens={{ childrenGap: 6 }}>
              <Toggle label={strings.AssignedToMeOnlyLabel} checked={!!item.source.plannerAssignedToMeOnly} onChange={(_, checked) => this.handleUpdateAssignedSource(item.adminSourceId, { plannerAssignedToMeOnly: !!checked })} />
              <Toggle label={strings.ShowCompletedTasksLabel} checked={item.source.showCompletedTasks !== false} onChange={(_, checked) => this.handleUpdateAssignedSource(item.adminSourceId, { showCompletedTasks: !!checked })} />
            </Stack>}
            <AllowedOverrideControls value={item.allowedOverrides} planner={item.source.sourceType === 'planner'} onChange={allowedOverrides => this.handleUpdateAllowedOverrides(item.adminSourceId, allowedOverrides)} />
            <div style={{ fontSize: 12, color: '#605e5c' }}>{strings.AudiencesLabel}: {audienceText || strings.NoAudienceLabel}</div>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton text={strings.GroupsLabel} onClick={() => this.handleEditAssignedSourceAudiences(item.assignmentId || item.adminSourceId).catch(err => console.error(err))} />
              <PrimaryButton text={strings.DoneLabel} onClick={() => this.toggleEditSource(undefined)} />
              <DefaultButton text={strings.DeleteLabel} onClick={() => this.handleDeleteAssignedSource(item.assignmentId || item.adminSourceId)} />
            </Stack>
          </Stack>
        ) : (
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <div style={{ width: 16, height: 16, backgroundColor: item.source.color, borderRadius: 2, flexShrink: 0 }} />
            <div style={{ flex: 1 }}>
              <strong style={{ fontSize: 13 }}>{item.source.name}</strong>
              <div>{assignmentPolicyLabel(item)}</div>
              {item.source.sourceType === 'sharepoint' && (
                <div style={{ fontSize: 11, color: '#605e5c' }}>
                  {strings.SiteLabel}: {item.source.sharePointSiteName || strings.SiteNameUnavailableLabel}
                </div>
              )}
              <div style={{ fontSize: 11, color: '#605e5c' }}>{getSourceTypeDisplayName(item.source.sourceType)} • {strings.AudiencesLabel}: {audienceText || strings.NoAudienceLabel}</div>
            </div>
            <IconButton iconProps={{ iconName: 'Edit' }} title={strings.EditLabel} ariaLabel={strings.EditLabel} onClick={() => this.toggleEditSource(item.assignmentId || item.adminSourceId)} />
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
            <TextField label={strings.DisplayNameLabel} value={item.displayName} onChange={(_, value) => this.handleUpdateIcsCatalogItem(item.adminIcsId, { displayName: value || '' })} />
            <TextField label={strings.IcsUrlLabel} value={item.icsUrl} onChange={(_, value) => this.handleUpdateIcsCatalogItem(item.adminIcsId, { icsUrl: value || '' })} />
            <div style={{ fontSize: 12, color: '#605e5c' }}>{strings.AudiencesLabel}: {audienceText || strings.NoAudienceLabel}</div>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton text={strings.GroupsLabel} onClick={() => this.handleEditIcsAudiences(item.adminIcsId).catch(err => console.error(err))} />
              <PrimaryButton text={strings.DoneLabel} onClick={() => this.toggleEditIcs(undefined)} />
              <DefaultButton text={strings.DeleteLabel} onClick={() => this.handleDeleteIcsCatalogItem(item.adminIcsId)} />
            </Stack>
          </Stack>
        ) : (
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <Icon iconName="World" />
            <div style={{ flex: 1 }}>
              <strong style={{ fontSize: 13 }}>{item.displayName}</strong>
              <div style={{ fontSize: 11, color: '#605e5c' }}>{item.icsUrl}</div>
              <div style={{ fontSize: 11, color: '#605e5c' }}>{strings.AudiencesLabel}: {audienceText || strings.NoAudienceLabel}</div>
            </div>
            <IconButton iconProps={{ iconName: 'Edit' }} title={strings.EditLabel} ariaLabel={strings.EditLabel} onClick={() => this.toggleEditIcs(item.adminIcsId)} />
          </Stack>
        )}
      </div>
    );
  }

  private onRenderFooterContent = (): React.ReactElement => (
    <Stack horizontal tokens={{ childrenGap: 8 }}>
      <PrimaryButton onClick={this.handleSave} disabled={this.state.isSaving} text={this.state.isSaving ? strings.SavingLabel : strings.SaveLabel} />
      <DefaultButton onClick={this.handleDismiss} disabled={this.state.isSaving} text={strings.CancelLabel} />
      <DefaultButton onClick={this.handleResetDraft} disabled={this.state.isSaving} text={strings.ResetDraftToDefaultsLabel} />
    </Stack>
  );

  public render(): React.ReactElement {
    const { isOpen, loadNotice } = this.props;
    const { settings, showAddDialog } = this.state;
    const startOptions: IDropdownOption[] = [];
    const latestStart = Math.max(0, 24 * 60 - settings.visibleHourCount * 60);
    for (let minutes = 0; minutes <= latestStart; minutes += settings.slotDurationMinutes) {
      startOptions.push({ key: minutes, text: formatCalendarTime(new Date(2000, 0, 1, 0, minutes), this.props.locale) });
    }
    const visibleHourOptions = Array.from({ length: 24 }, (_, index) => ({ key: index + 1, text: String(index + 1) }));

    return (
    <Panel isOpen={isOpen} onDismiss={this.handleDismiss} isBlocking={this.state.isSaving} isLightDismiss={false} hasCloseButton={!this.state.isSaving} type={PanelType.medium} headerText={showAddDialog ? strings.AddAdminDefaultLabel : strings.AdminCalendarDefaultsTitle} onRenderFooterContent={!showAddDialog ? this.onRenderFooterContent : undefined} isFooterAtBottom={true}>
        {this.state.saveError && <MessageBar messageBarType={MessageBarType.error}>{this.state.saveError}</MessageBar>}
        <fieldset disabled={this.state.isSaving} style={{ border: 0, padding: 0, margin: 0, minWidth: 0 }}>

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
              label={strings.DefaultViewLabel}
              options={[
                { key: 'day', text: strings.DayLabel },
                { key: 'week', text: strings.WeekLabel },
                { key: 'month', text: strings.MonthLabel }
              ]}
              selectedKey={settings.defaultView}
              onChange={(_, option) => this.setSessionState(this.editSession, prev => ({
                settings: {
                  ...prev.settings,
                  defaultView: option?.key as IAdminWebPartSettings['defaultView']
                }
              }))}
            />

            <Label>{strings.CacheDefaultsLabel}</Label>
            <Toggle
              label={strings.EnableCacheLabel}
              checked={settings.enableCache}
              onText={strings.OnLabel}
              offText={strings.OffLabel}
              onChange={(_, checked) => this.setSessionState(this.editSession, prev => ({ settings: { ...prev.settings, enableCache: checked !== false } }))}
            />
            <Slider
              label={strings.CacheDurationMinutesLabel}
              min={1}
              max={60}
              step={1}
              value={settings.cacheDurationMinutes}
              disabled={!settings.enableCache}
              showValue={true}
              onChange={value => this.setSessionState(this.editSession, prev => ({ settings: { ...prev.settings, cacheDurationMinutes: value } }))}
            />

            <Label>{strings.TimelineDefaultsLabel}</Label>
            <Toggle
              label={strings.ShowWeekendsLabel}
              checked={settings.showWeekends}
              onText={strings.OnLabel}
              offText={strings.OffLabel}
              onChange={(_, checked) => this.setSessionState(this.editSession, prev => ({ settings: { ...prev.settings, showWeekends: checked !== false } }))}
            />
            <Dropdown
              label={strings.SlotDurationLabel}
              selectedKey={settings.slotDurationMinutes}
              options={[{ key: 15, text: formatLocalizedString(strings.MinutesLabel, 15) }, { key: 30, text: formatLocalizedString(strings.MinutesLabel, 30) }, { key: 60, text: formatLocalizedString(strings.MinutesLabel, 60) }]}
              onChange={(_, option) => {
                const slot = option?.key as 15 | 30 | 60;
                this.setSessionState(this.editSession, prev => ({ settings: { ...prev.settings, slotDurationMinutes: slot, preferredStartMinutes: Math.floor(Math.min(prev.settings.preferredStartMinutes, 24 * 60 - prev.settings.visibleHourCount * 60) / slot) * slot } }));
              }}
            />
            <Dropdown
              label={strings.PreferredStartTimeLabel}
              selectedKey={settings.preferredStartMinutes}
              options={startOptions}
              onChange={(_, option) => this.setSessionState(this.editSession, prev => ({ settings: { ...prev.settings, preferredStartMinutes: Number(option?.key) } }))}
            />
            <Dropdown
              label={strings.VisibleHoursLabel}
              selectedKey={settings.visibleHourCount}
              options={visibleHourOptions}
              onChange={(_, option) => {
                const visibleHourCount = Number(option?.key);
                this.setSessionState(this.editSession, prev => ({ settings: { ...prev.settings, visibleHourCount, preferredStartMinutes: Math.min(prev.settings.preferredStartMinutes, 24 * 60 - visibleHourCount * 60) } }));
              }}
            />

            <div>
              <PrimaryButton text={strings.AddAdminDefaultLabel} iconProps={{ iconName: 'Add' }} onClick={this.handleOpenAddDialog} />
            </div>

            <div>
              <Label>{strings.AdminDefaultCalendarsLabel}</Label>
              <Stack tokens={{ childrenGap: 8 }}>
                {settings.assignedSources.length > 0 || (settings.exchangeMailboxAssignments || []).length > 0
                  ? this.renderGroupedAssignments()
                  : <div style={{ fontSize: 12, color: '#605e5c' }}>{strings.NoAdminDefaultCalendarsLabel}</div>}
              </Stack>
            </div>

            <div>
              <Label>{strings.AdminIcsCatalogLabel}</Label>
              <Stack tokens={{ childrenGap: 8 }}>
                {settings.icsCatalog.length > 0
                  ? settings.icsCatalog.map((item, index) => this.renderIcsCatalogItem(item, index))
                  : <div style={{ fontSize: 12, color: '#605e5c' }}>{strings.NoAdminIcsCatalogItemsLabel}</div>}
              </Stack>
            </div>
          </Stack>
        )}
        </fieldset>
      </Panel>
    );
  }
}
