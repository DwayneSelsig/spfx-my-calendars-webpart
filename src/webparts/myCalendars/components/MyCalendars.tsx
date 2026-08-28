import * as React from 'react';
import type { IMyCalendarsProps } from './IMyCalendarsProps';
import type { ICalendarEvent as IEvent } from '../models/ICalendarEvent';
import { CalendarViewType } from '../models/ICalendarSettings';
import { ExchangeCalendarService, type IExchangeCalendar } from '../services/ExchangeCalendarService';
import { SharePointCalendarService } from '../services/SharePointCalendarService';
import { PlannerTaskService, type IPlannerPlan } from '../services/PlannerTaskService';
import { TeamsShiftsService } from '../services/TeamsShiftsService';
import { UnifiedGroupCalendarService, type IUnifiedGroupItem } from '../services/UnifiedGroupCalendarService';
import { DayView } from './views/DayView';
import { WeekView } from './views/WeekView';
import { MonthView } from './views/MonthView';
import { SearchResultsView } from './views/SearchResultsView';
import { CalendarToolbar } from './CalendarToolbar';
import { CommandBar, ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';
import { Callout } from '@fluentui/react/lib/Callout';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Text } from '@fluentui/react/lib/Text';
import { CommandBarButton } from '@fluentui/react/lib/Button';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { SettingsPanel } from './SettingsPanel';
//import { CalendarToolbar } from './CalendarToolbar';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { getSourceIconName, getSourceTypeDisplayName } from '../utils/sourceIconHelper';

type ServiceKey = 'exchange' | 'ics' | 'sharepoint' | 'planner' | 'teamsShifts' | 'unifiedGroup';
type ServiceStatus = 'loading' | 'ready' | 'error';
type IndexedEvent = IEvent & { searchIndexText?: string };

const defaultLoadingSources: Record<ServiceKey, ServiceStatus> = {
  exchange: 'ready',
  ics: 'ready',
  sharepoint: 'ready',
  planner: 'ready',
  teamsShifts: 'ready',
  unifiedGroup: 'ready'
};

const defaultLoadErrors: Record<ServiceKey, string | undefined> = {
  exchange: undefined,
  ics: undefined,
  sharepoint: undefined,
  planner: undefined,
  teamsShifts: undefined,
  unifiedGroup: undefined
};

const styles = mergeStyleSets({
  myCalendars: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    color: 'var(--bodyText, #323130)'
  },
  commandBar: {
    flexShrink: 0,
    borderBottom: '1px solid var(--neutralLight, #edebe9)',
    paddingBottom: 10
  },
  calendarContainer: {
    flex: 1,
    overflow: 'hidden',
    position: 'relative'
  },
  loadingStatusWrapper: {
    display: 'inline-flex',
    alignItems: 'center'
  },
  loadingStatusCallout: {
    padding: '12px 14px',
    minWidth: 260
  },
  loadingStatusTitle: {
    display: 'block',
    marginBottom: 8,
    fontWeight: 600
  },
  loadingStatusList: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: 8
  },
  loadingStatusRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    columnGap: 12
  },
  loadingStatusLabel: {
    fontSize: 12,
    color: 'var(--neutralPrimary, #323130)'
  },
  loadingStatusState: {
    display: 'inline-flex',
    alignItems: 'center',
    columnGap: 6,
    fontSize: 12,
    color: 'var(--neutralSecondary, #605e5c)'
  },
  loadingStatusError: {
    fontSize: 11,
    color: 'var(--errorText, #a80000)',
    marginLeft: 4
  }
});

interface IMyCalendarsState {
  appointments: IEvent[];
  currentDate: Date;
  currentView: CalendarViewType;
  previousView: CalendarViewType;
  isLoading: boolean;
  isSettingsPanelOpen: boolean;
  searchQuery: string;
  graphClient: MSGraphClientV3 | undefined;
  loadingSources: Record<ServiceKey, ServiceStatus>;
  loadErrors: Record<ServiceKey, string | undefined>;
  isLoadingStatusOpen: boolean;
  showRefreshButton: boolean;
}

export default class MyCalendars extends React.Component<IMyCalendarsProps, IMyCalendarsState> {
  private searchDebounceTimer: number | null = null;
  private activeLoadId = 0;
  private loadingStatusWrapperRef = React.createRef<HTMLDivElement>();
  private refreshTimer: number | null = null;
  private loadedMonthsBySource = new Map<string, Set<string>>();
  private knownSourceIdsByService: Record<ServiceKey, Set<string>> = {
    exchange: new Set(), ics: new Set(), sharepoint: new Set(), planner: new Set(), teamsShifts: new Set(), unifiedGroup: new Set()
  };
  private rangeLoadPromise: Promise<void> = Promise.resolve();
  private loadGeneration = 0;
  private exchangeCalendarsPromise?: Promise<IExchangeCalendar[]>;
  private plannerPlansPromise?: Promise<IPlannerPlan[]>;
  private unifiedGroupsPromise?: Promise<IUnifiedGroupItem[]>;
  private joinedTeamIdsPromise?: Promise<Set<string>>;
  private teamsShiftsService?: TeamsShiftsService;

  constructor(props: IMyCalendarsProps) {
    super(props);
    
    this.state = {
      appointments: [],
      currentDate: new Date(),
      currentView: props.settings.defaultView,
      previousView: props.settings.defaultView,
      isLoading: false,
      isSettingsPanelOpen: false,
      searchQuery: '',
      graphClient: undefined,
      loadingSources: { ...defaultLoadingSources },
      loadErrors: { ...defaultLoadErrors },
      isLoadingStatusOpen: false,
      showRefreshButton: true
    };
  }

  public componentDidMount(): void {
    this.loadAppointments().then(() => this.ensureVisibleRange()).catch(err => console.error('Failed to load appointments:', err));
    // Resolve and store graphClient for use in SettingsPanel
    this.props.context.msGraphClientFactory.getClient('3')
      .then(client => this.setState({ graphClient: client }))
      .catch(err => console.error('Failed to initialize graph client:', err));
  }

  private getLoadingSettingsFingerprint(settings: import('../models/ICalendarSettings').ICalendarSettings): string {
    const loadingSettings = { ...settings } as unknown as Record<string, unknown>;
    ['defaultView', 'preferredStartMinutes', 'visibleHourCount', 'slotDurationMinutes', 'showWeekends', 'userPreferredStartMinutes', 'userVisibleHourCount']
      .forEach(key => delete loadingSettings[key]);
    return JSON.stringify(loadingSettings);
  }

  public componentDidUpdate(prevProps: IMyCalendarsProps): void {
    const defaultViewChanged = prevProps.settings.defaultView !== this.props.settings.defaultView;
    if (defaultViewChanged) {
      const defaultView = this.props.settings.defaultView;
      this.setState(prevState => ({
        currentView: prevState.currentView === 'search' ? 'search' : defaultView,
        previousView: defaultView
      }), () => {
        this.ensureVisibleRange(this.state.currentDate, defaultView).catch(err => console.error('Failed to load visible range:', err));
      });
    }

    try {
      const prevJson = JSON.stringify(prevProps.settings);
      const currJson = JSON.stringify(this.props.settings);
      if (prevJson !== currJson) {
        if (this.getLoadingSettingsFingerprint(prevProps.settings) !== this.getLoadingSettingsFingerprint(this.props.settings)) {
          this.loadAppointments().then(() => this.ensureVisibleRange()).catch(err => console.error('Failed to load appointments:', err));
        } else if (!defaultViewChanged) {
          this.ensureVisibleRange().catch(err => console.error('Failed to load visible range:', err));
        }
      }
    } catch {
      // Fallback to sources change detection
      if (prevProps.settings.sources !== this.props.settings.sources) {
        this.loadAppointments().then(() => this.ensureVisibleRange()).catch(err => console.error('Failed to load appointments:', err));
      }
    }
  }

  public componentWillUnmount(): void {
    if (this.searchDebounceTimer) {
      clearTimeout(this.searchDebounceTimer);
      this.searchDebounceTimer = null;
    }

    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer);
      this.refreshTimer = null;
    }
  }

  private getMonthKeys(startDate: Date, endDate: Date): string[] {
    const keys: string[] = [];
    const firstMonth = startDate.getFullYear() * 12 + startDate.getMonth();
    const lastInstant = new Date(Math.max(startDate.getTime(), endDate.getTime() - 1));
    const lastMonth = lastInstant.getFullYear() * 12 + lastInstant.getMonth();
    for (let monthIndex = firstMonth; monthIndex <= lastMonth; monthIndex++) {
      const year = Math.floor(monthIndex / 12);
      keys.push(`${year}-${monthIndex - year * 12 + 1}`);
    }
    return keys;
  }

  private clearRangeCache(): void {
    this.loadedMonthsBySource.clear();
    (Object.keys(this.knownSourceIdsByService) as ServiceKey[]).forEach(key => this.knownSourceIdsByService[key].clear());
    this.exchangeCalendarsPromise = undefined;
    this.plannerPlansPromise = undefined;
    this.unifiedGroupsPromise = undefined;
    this.joinedTeamIdsPromise = undefined;
    this.teamsShiftsService = undefined;
  }

  private registerSource(service: ServiceKey, sourceId: string): void {
    this.knownSourceIdsByService[service].add(sourceId);
  }

  private areMonthsLoaded(sourceId: string, monthKeys: string[]): boolean {
    const loaded = this.loadedMonthsBySource.get(sourceId);
    return !!loaded && monthKeys.every(key => loaded.has(key));
  }

  private markMonthsLoaded(sourceId: string, monthKeys: string[]): void {
    const loaded = this.loadedMonthsBySource.get(sourceId) || new Set<string>();
    monthKeys.forEach(key => loaded.add(key));
    this.loadedMonthsBySource.set(sourceId, loaded);
  }

  private getVisibleRange(date: Date = this.state.currentDate, view: CalendarViewType = this.state.previousView): { start: Date; end: Date } {
    if (view === 'month') {
      const first = new Date(date.getFullYear(), date.getMonth(), 1);
      const start = new Date(first);
      start.setDate(first.getDate() - first.getDay());
      const end = new Date(start);
      end.setDate(start.getDate() + 42);
      return { start, end };
    }
    if (view === 'week') {
      const start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      if (!this.props.settings.showWeekends) {
        while (start.getDay() === 0 || start.getDay() === 6) start.setDate(start.getDate() + 1);
      }
      const end = new Date(start);
      if (this.props.settings.showWeekends) {
        end.setDate(end.getDate() + 7);
      } else {
        let weekdays = 0;
        while (weekdays < 5) {
          if (end.getDay() !== 0 && end.getDay() !== 6) weekdays++;
          end.setDate(end.getDate() + 1);
        }
      }
      return { start, end };
    }
    const start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const end = new Date(start);
    end.setDate(end.getDate() + 1);
    return { start, end };
  }

  private ensureVisibleRange(date: Date = this.state.currentDate, view: CalendarViewType = this.state.previousView): Promise<void> {
    const range = this.getVisibleRange(date, view);
    const monthKeys = this.getMonthKeys(range.start, range.end);
    const missingServices = new Set(this.getEnabledServiceKeys().filter(service => {
      const knownSources = Array.from(this.knownSourceIdsByService[service]);
      return knownSources.length === 0 || knownSources.some(sourceId => !this.areMonthsLoaded(sourceId, monthKeys));
    }));
    if (missingServices.size === 0) return Promise.resolve();
    const generation = this.loadGeneration;
    this.rangeLoadPromise = this.rangeLoadPromise
      .catch(() => undefined)
      .then(() => generation === this.loadGeneration
        ? this.loadAppointments(range.start, range.end, false, missingServices)
        : Promise.resolve());
    return this.rangeLoadPromise;
  }

  private async loadAppointments(requestStart?: Date, requestEnd?: Date, reset: boolean = true, requestedServices?: Set<ServiceKey>): Promise<void> {
    if (reset) this.loadGeneration++;
    const currentGeneration = this.loadGeneration;
    const loadId = ++this.activeLoadId;
    if (reset) this.clearRangeCache();
    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer);
      this.refreshTimer = null;
    }
    const enabledServices = this.getEnabledServiceKeys();
    const servicesToLoad = requestedServices || new Set(enabledServices);
    const initialLoadingSources = Array.from(servicesToLoad).reduce((acc, key) => {
      acc[key] = 'loading';
      return acc;
    }, reset ? { ...defaultLoadingSources } : { ...this.state.loadingSources });

    this.setState({
      isLoading: true,
      appointments: reset ? [] : this.state.appointments,
      loadingSources: initialLoadingSources,
      loadErrors: reset ? { ...defaultLoadErrors } : this.state.loadErrors,
      isLoadingStatusOpen: false,
      showRefreshButton: false
    });
    const httpClient = this.props.context.httpClient;
    const graphClientPromise = this.props.context.msGraphClientFactory.getClient('3');
    const graphClient = await graphClientPromise;
    const exchangeService = new ExchangeCalendarService(httpClient, graphClient);
    exchangeService.setGraphClient(graphClient);
    const sharePointService = new SharePointCalendarService(httpClient, graphClient);
    sharePointService.setGraphClient(graphClient);
    const plannerService = new PlannerTaskService(httpClient, graphClient);
    plannerService.setGraphClient(graphClient);
    const teamsShiftsService = this.teamsShiftsService || new TeamsShiftsService(httpClient, graphClient);
    teamsShiftsService.setGraphClient(graphClient);
    this.teamsShiftsService = teamsShiftsService;
    const unifiedGroupService = new UnifiedGroupCalendarService(httpClient, graphClient);
    unifiedGroupService.setGraphClient(graphClient);
    
    // Calculate date range for filtering (current month ± 3 months)
    const today = new Date();
    const startDate = requestStart || new Date(today.getFullYear(), today.getMonth() - 3, 1);
    const endDate = requestEnd || new Date(today.getFullYear(), today.getMonth() + 4, 1);
    const requestedMonthKeys = this.getMonthKeys(startDate, endDate);
    const markLoaded = (sourceId: string): void => {
      if (loadId === this.activeLoadId && currentGeneration === this.loadGeneration) {
        this.markMonthsLoaded(sourceId, requestedMonthKeys);
      }
    };
    const registerLoadedSource = (service: ServiceKey, sourceId: string): void => {
      if (loadId === this.activeLoadId && currentGeneration === this.loadGeneration) this.registerSource(service, sourceId);
    };
    Array.from(servicesToLoad).forEach(service => this.registerSource(service, `$service:${service}`));
    
    const updateStatus = (service: ServiceKey, status: ServiceStatus, errorMessage?: string): void => {
      if (loadId !== this.activeLoadId) {
        return;
      }

      if (status === 'ready') markLoaded(`$service:${service}`);

      this.setState(prev => ({
        loadingSources: {
          ...prev.loadingSources,
          [service]: status
        },
        loadErrors: {
          ...prev.loadErrors,
          [service]: errorMessage
        }
      }));
    };

    const appendAppointments = (appointments: IEvent[]): void => {
      if (loadId !== this.activeLoadId || appointments.length === 0) {
        return;
      }

      const normalizedAppointments: IndexedEvent[] = appointments.map(apt => ({
        ...apt,
        colorHex: apt.colorHex || '#0078d4',
        sourceIconName: getSourceIconName(apt.sourceType, apt.sourceIconName),
        sourceDisplayName: apt.sourceDisplayName || getSourceTypeDisplayName(apt.sourceType, apt.sourceIconName),
        // Build a reusable lowercase search index per event to avoid repeated string work while typing.
        searchIndexText: this.buildSearchIndexText(apt)
      }));

      this.setState(prev => {
        const byIdentity = new Map<string, IndexedEvent>();
        prev.appointments.forEach(event => byIdentity.set(`${event.sourceId}:${event.id}`, event as IndexedEvent));
        normalizedAppointments.forEach(event => byIdentity.set(`${event.sourceId}:${event.id}`, event));
        return { appointments: Array.from(byIdentity.values()) };
      });
    };

    const enabledSources = this.props.settings.sources.filter(source => source.isEnabled);

    const sourceGroups: Record<ServiceKey, typeof enabledSources> = {
      exchange: enabledSources.filter(source => source.sourceType === 'exchange'),
      ics: enabledSources.filter(source => source.sourceType === 'ics'),
      sharepoint: enabledSources.filter(source => source.sourceType === 'sharepoint'),
      planner: enabledSources.filter(source => source.sourceType === 'planner'),
      teamsShifts: enabledSources.filter(source => source.sourceType === 'teamsShifts'),
      unifiedGroup: enabledSources.filter(source => source.sourceType === 'unifiedGroup')
    };
    (Object.keys(sourceGroups) as ServiceKey[]).forEach(service => {
      sourceGroups[service].forEach(source => this.registerSource(service, source.id));
    });

    const tasks: Array<Promise<void>> = [];

    // Load Exchange calendars (user's own calendars)
    if (servicesToLoad.has('exchange')) {
      tasks.push((async () => {
        let hadError = false;

        try {
          this.exchangeCalendarsPromise = this.exchangeCalendarsPromise || exchangeService.getCalendars();
          const userCalendars = await this.exchangeCalendarsPromise;
          const exchangeCalendarStates = this.props.settings.exchangeCalendarStates || {};
          userCalendars.filter(calendar => exchangeCalendarStates[calendar.id] !== false)
            .forEach(calendar => registerLoadedSource('exchange', `exchange_${calendar.id}`));
          const calendarPromises = userCalendars
            .filter(calendar => exchangeCalendarStates[calendar.id] !== false)
            .filter(calendar => !this.areMonthsLoaded(`exchange_${calendar.id}`, requestedMonthKeys))
            .map(async calendar => {
              try {
                const events = await exchangeService.getCalendarEvents(
                  calendar.id,
                  startDate,
                  endDate,
                  undefined // current user
                );
                markLoaded(`exchange_${calendar.id}`);
                return events.map(event => ({
                  ...event,
                  sourceId: `exchange_${calendar.id}`,
                  sourceDisplayName: calendar.name,
                  colorHex: calendar.hexColor,
                  sourceType: 'exchange' as const,
                  showSourceLogo: this.props.settings.exchangeShowSourceLogo ?? true
                } as IEvent));
              } catch (error) {
                hadError = true;
                console.error(`Failed to load Exchange calendar ${calendar.name}:`, error);
                return [] as IEvent[];
              }
            });

          const appointmentGroups = await Promise.all(calendarPromises);
          const flattenedAppointments = appointmentGroups.reduce<IEvent[]>((acc, group) => acc.concat(group), []);
          appendAppointments(flattenedAppointments);

          if (sourceGroups.exchange.length > 0) {
            const manualExchangeAppointments = await Promise.all(sourceGroups.exchange.map(async source => {
              try {
                if (!source.exchangeCalendarId || this.areMonthsLoaded(source.id, requestedMonthKeys)) {
                  return [] as IEvent[];
                }

                const events = await exchangeService.getCalendarEvents(
                  source.exchangeCalendarId,
                  startDate,
                  endDate,
                  source.exchangeMailbox
                );
                markLoaded(source.id);

                return events.map(event => ({
                  ...event,
                  sourceId: source.id,
                  sourceDisplayName: source.name,
                  colorHex: source.color,
                  sourceType: 'exchange' as const,
                  showSourceLogo: source.showSourceLogo ?? this.props.settings.exchangeShowSourceLogo ?? true
                } as IEvent));
              } catch (error) {
                hadError = true;
                console.error(`Failed to load Exchange source ${source.name}:`, error);
                return [] as IEvent[];
              }
            }));

            const flattenedManualAppointments = manualExchangeAppointments.reduce<IEvent[]>((acc, group) => acc.concat(group), []);
            appendAppointments(flattenedManualAppointments);
          }
        } catch (error) {
          hadError = true;
          if (currentGeneration === this.loadGeneration) this.exchangeCalendarsPromise = undefined;
          console.error('Failed to load user Exchange calendars:', error);
        }

        updateStatus('exchange', hadError ? 'error' : 'ready', hadError ? 'One or more Exchange calendars failed.' : undefined);
      })());
    }

    if (servicesToLoad.has('sharepoint') && sourceGroups.sharepoint.length > 0) {
      tasks.push((async () => {
        let hadError = false;
        const appointmentsBySource = await Promise.all(sourceGroups.sharepoint.map(async source => {
          try {
            if (source.sharePointSiteId && source.sharePointListId && !this.areMonthsLoaded(source.id, requestedMonthKeys)) {
              const items = await sharePointService.getListEvents(
                source.sharePointSiteId,
                source.sharePointListId,
                startDate,
                endDate,
                source.sharePointFieldMapping
              );
              markLoaded(source.id);
              return items.map(item => ({
                ...item,
                sourceId: source.id,
                sourceDisplayName: source.name,
                colorHex: source.color,
                sourceType: 'sharepoint' as const,
                showSourceLogo: this.props.settings.sharePointShowSourceLogo ?? true
              } as IEvent));
            }
          } catch (error) {
            hadError = true;
            console.error(`Failed to load SharePoint calendar ${source.name}:`, error);
          }

          return [] as IEvent[];
        }));

        const flattenedAppointments = appointmentsBySource.reduce<IEvent[]>((acc, group) => acc.concat(group), []);
        appendAppointments(flattenedAppointments);
        updateStatus('sharepoint', hadError ? 'error' : 'ready', hadError ? 'One or more SharePoint calendars failed.' : undefined);
      })());
    }

    if (servicesToLoad.has('planner') && this.props.settings.plannerShowAllCalendars) {
      tasks.push((async () => {
        let hadError = false;

        try {
          this.plannerPlansPromise = this.plannerPlansPromise || plannerService.getUserPlans();
          const plans = await this.plannerPlansPromise;
          const autoPlannerSources = plans.map(plan => ({
            id: `auto_planner_${plan.id}`,
            origin: 'user' as const,
            sourceType: 'planner' as const,
            name: plan.title,
            color: this.props.settings.organizationPrimaryColor || '#0078d4',
            isEnabled: true,
            plannerPlanId: plan.id,
            plannerPlanTitle: plan.title,
            plannerAssignedToMeOnly: this.props.settings.plannerShowAllAssignedToMeOnly ?? false,
            showCompletedTasks: true,
            showSourceLogo: this.props.settings.plannerShowSourceLogo ?? true
          }));
          autoPlannerSources.forEach(source => registerLoadedSource('planner', source.id));

          const appointmentsBySource = await Promise.all(autoPlannerSources.map(async source => {
            try {
              if (!source.plannerPlanId || this.areMonthsLoaded(source.id, requestedMonthKeys)) {
                return [] as IEvent[];
              }

              const events = await plannerService.getTasks(
                source.plannerPlanId,
                startDate,
                endDate,
                source.plannerAssignedToMeOnly ?? false,
                source.showCompletedTasks ?? true,
                source,
                this.props.settings.plannerShowSourceLogo ?? true
              );
              markLoaded(source.id);
              return events;
            } catch (error) {
              hadError = true;
              console.error(`Failed to load Planner tasks ${source.name}:`, error);
              return [] as IEvent[];
            }
          }));

          const flattenedAppointments = appointmentsBySource.reduce<IEvent[]>((acc, group) => acc.concat(group), []);
          appendAppointments(flattenedAppointments);
        } catch (error) {
          hadError = true;
          if (currentGeneration === this.loadGeneration) this.plannerPlansPromise = undefined;
          console.error('Failed to load Planner plans for auto mode:', error);
        }

        updateStatus('planner', hadError ? 'error' : 'ready', hadError ? 'One or more Planner sources failed.' : undefined);
      })());
    } else if (servicesToLoad.has('planner') && sourceGroups.planner.length > 0) {
      tasks.push((async () => {
        let hadError = false;
        const appointmentsBySource = await Promise.all(sourceGroups.planner.map(async source => {
          try {
            if (source.plannerPlanId && !this.areMonthsLoaded(source.id, requestedMonthKeys)) {
              const events = await plannerService.getTasks(
                source.plannerPlanId,
                startDate,
                endDate,
                source.plannerAssignedToMeOnly ?? false,
                source.showCompletedTasks ?? true,
                source,
                this.props.settings.plannerShowSourceLogo ?? true
              );
              markLoaded(source.id);
              return events;
            }
          } catch (error) {
            hadError = true;
            console.error(`Failed to load Planner tasks ${source.name}:`, error);
          }

          return [] as IEvent[];
        }));

        const flattenedAppointments = appointmentsBySource.reduce<IEvent[]>((acc, group) => acc.concat(group), []);
        appendAppointments(flattenedAppointments);
        updateStatus('planner', hadError ? 'error' : 'ready', hadError ? 'One or more Planner sources failed.' : undefined);
      })());
    }

    if (servicesToLoad.has('teamsShifts') && this.props.settings.teamsShiftsShowAllCalendars) {
      tasks.push((async () => {
        let hadError = false;
        const autoSource = {
          id: 'auto_teamsShifts',
          origin: 'user' as const,
          sourceType: 'teamsShifts' as const,
          name: 'Teams Shifts',
          color: this.props.settings.organizationPrimaryColor || '#4a4fbe',
          isEnabled: true,
          showSourceLogo: this.props.settings.teamsShiftsShowSourceLogo ?? true
        };

        try {
          registerLoadedSource('teamsShifts', autoSource.id);
          if (!this.areMonthsLoaded(autoSource.id, requestedMonthKeys)) {
            const events = await teamsShiftsService.getShiftsForJoinedTeams(
              startDate,
              endDate,
              autoSource,
              this.props.settings.teamsShiftsShowSourceLogo ?? true
            );
            markLoaded(autoSource.id);
            appendAppointments(events);
          }
        } catch (error) {
          hadError = true;
          console.error('Failed to load Teams shifts for auto mode:', error);
        }

        updateStatus('teamsShifts', hadError ? 'error' : 'ready', hadError ? 'One or more Teams Shifts sources failed.' : undefined);
      })());
    } else if (servicesToLoad.has('teamsShifts') && sourceGroups.teamsShifts.length > 0) {
      tasks.push((async () => {
        let hadError = false;
        const appointmentsBySource = await Promise.all(sourceGroups.teamsShifts.map(async source => {
          try {
            if (this.areMonthsLoaded(source.id, requestedMonthKeys)) return [] as IEvent[];
            const events = await teamsShiftsService.getShiftsForJoinedTeams(
              startDate,
              endDate,
              source,
              this.props.settings.teamsShiftsShowSourceLogo ?? true
            );
            markLoaded(source.id);
            return events;
          } catch (error) {
            hadError = true;
            console.error(`Failed to load Teams shifts ${source.name}:`, error);
            return [] as IEvent[];
          }
        }));

        const flattenedAppointments = appointmentsBySource.reduce<IEvent[]>((acc, group) => acc.concat(group), []);
        appendAppointments(flattenedAppointments);
        updateStatus('teamsShifts', hadError ? 'error' : 'ready', hadError ? 'One or more Teams Shifts sources failed.' : undefined);
      })());
    }

    if (servicesToLoad.has('unifiedGroup') && this.props.settings.unifiedGroupShowAllCalendars) {
      tasks.push((async () => {
        let hadError = false;

        try {
          this.unifiedGroupsPromise = this.unifiedGroupsPromise || unifiedGroupService.getUnifiedGroups();
          this.joinedTeamIdsPromise = this.joinedTeamIdsPromise || unifiedGroupService.getJoinedTeamIds();
          const [groups, joinedTeamIds] = await Promise.all([this.unifiedGroupsPromise, this.joinedTeamIdsPromise]);
          groups.forEach(group => registerLoadedSource('unifiedGroup', `auto_unifiedGroup_${group.id}`));

          const appointmentsByGroup = await Promise.all(groups.map(async group => {
            try {
              const sourceId = `auto_unifiedGroup_${group.id}`;
              if (this.areMonthsLoaded(sourceId, requestedMonthKeys)) return [] as IEvent[];
              const iconName = joinedTeamIds.has(group.id) ? 'TeamsLogo' : 'Group';
              const events = await unifiedGroupService.getGroupEvents(group.id, startDate, endDate);
              markLoaded(sourceId);
              return events.map(event => ({
                ...event,
                sourceId: `auto_unifiedGroup_${group.id}`,
                sourceDisplayName: group.displayName,
                colorHex: this.props.settings.organizationPrimaryColor || '#5b5fc7',
                sourceType: 'unifiedGroup' as const,
                showSourceLogo: this.props.settings.unifiedGroupShowSourceLogo ?? true,
                sourceIconName: iconName
              } as IEvent));
            } catch (error) {
              hadError = true;
              console.error(`Failed to load group calendar ${group.displayName}:`, error);
              return [] as IEvent[];
            }
          }));

          const flattenedAppointments = appointmentsByGroup.reduce<IEvent[]>((acc, group) => acc.concat(group), []);
          appendAppointments(flattenedAppointments);
        } catch (error) {
          hadError = true;
          if (currentGeneration === this.loadGeneration) {
            this.unifiedGroupsPromise = undefined;
            this.joinedTeamIdsPromise = undefined;
          }
          console.error('Failed to load group calendars for auto mode:', error);
        }

        updateStatus('unifiedGroup', hadError ? 'error' : 'ready', hadError ? 'One or more group calendars failed.' : undefined);
      })());
    } else if (servicesToLoad.has('unifiedGroup') && sourceGroups.unifiedGroup.length > 0) {
      tasks.push((async () => {
        let hadError = false;
        let joinedTeamIds: Set<string> = new Set();

        try {
          this.joinedTeamIdsPromise = this.joinedTeamIdsPromise || unifiedGroupService.getJoinedTeamIds();
          joinedTeamIds = await this.joinedTeamIdsPromise;
        } catch (error) {
          hadError = true;
          if (currentGeneration === this.loadGeneration) this.joinedTeamIdsPromise = undefined;
          console.error('Failed to load joined Teams for group calendars:', error);
        }

        const appointmentsBySource = await Promise.all(sourceGroups.unifiedGroup.map(async source => {
          try {
            if (!source.groupId || this.areMonthsLoaded(source.id, requestedMonthKeys)) {
              return [] as IEvent[];
            }

            const iconName = joinedTeamIds.has(source.groupId) ? 'TeamsLogo' : 'Group';
            const events = await unifiedGroupService.getGroupEvents(source.groupId, startDate, endDate);
            markLoaded(source.id);
            return events.map(event => ({
              ...event,
              sourceId: source.id,
              sourceDisplayName: source.name,
              colorHex: source.color,
              sourceType: 'unifiedGroup' as const,
              showSourceLogo: this.props.settings.unifiedGroupShowSourceLogo ?? true,
              sourceIconName: iconName
            } as IEvent));
          } catch (error) {
            hadError = true;
            console.error(`Failed to load group calendar ${source.name}:`, error);
            return [] as IEvent[];
          }
        }));

        const flattenedAppointments = appointmentsBySource.reduce<IEvent[]>((acc, group) => acc.concat(group), []);
        appendAppointments(flattenedAppointments);
        updateStatus('unifiedGroup', hadError ? 'error' : 'ready', hadError ? 'One or more group calendars failed.' : undefined);
      })());
    }

    if (servicesToLoad.has('ics')) {
      sourceGroups.ics.forEach(source => markLoaded(source.id));
      updateStatus('ics', 'ready');
    }
    await Promise.all(tasks.map(task => task.catch(() => undefined)));

    if (loadId === this.activeLoadId) {
      this.setState({ isLoading: false });
      this.refreshTimer = window.setTimeout(() => {
        if (loadId === this.activeLoadId) {
          this.setState({ showRefreshButton: true });
        }
      }, 2000);
    }
  }

  private handleDateChange = (date: Date): void => {
    this.setState({ currentDate: date }, () => {
      this.ensureVisibleRange(date).catch(err => console.error('Failed to load visible range:', err));
    });
  };

  private handleViewChange = (view: CalendarViewType): void => {
    if (view !== 'search') {
      this.setState({ currentView: view, previousView: view }, () => {
        this.ensureVisibleRange(this.state.currentDate, view).catch(err => console.error('Failed to load visible range:', err));
      });
      this.props.onDefaultViewChange(view);
    }
  };

  private navigateDate = (direction: 'prev' | 'next'): void => {
    const { currentDate, currentView } = this.state;
    const newDate = new Date(currentDate);

    switch (currentView) {
      case 'day':
        newDate.setDate(newDate.getDate() + (direction === 'next' ? 1 : -1));
        break;
      case 'week':
        if (this.props.settings.showWeekends) {
          newDate.setDate(newDate.getDate() + (direction === 'next' ? 7 : -7));
        } else {
          const step = direction === 'next' ? 1 : -1;
          let remaining = 5;
          while (remaining > 0) {
            newDate.setDate(newDate.getDate() + step);
            if (newDate.getDay() !== 0 && newDate.getDay() !== 6) remaining--;
          }
        }
        break;
      case 'month':
        newDate.setMonth(newDate.getMonth() + (direction === 'next' ? 1 : -1));
        break;
    }

    this.handleDateChange(newDate);
  };

  private getCommandBarItems = (): ICommandBarItemProps[] => {
    const searchItem: ICommandBarItemProps = {
      key: 'search',
      onRender: () => (
        <SearchBox
          placeholder="Search appointments..."
          onChange={(_event, newValue) => this.handleSearch(newValue || '')}
          styles={{
            root: {
              width: '100%',
              marginTop: 8
            }
          }}
        />
      )
    };

    return [
      searchItem
    ];
  };

  private getFarCommandBarItems = (): ICommandBarItemProps[] => {
    const { showRefreshButton } = this.state;

    const refreshItem: ICommandBarItemProps = {
      key: 'refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: () => {
        this.loadAppointments().then(() => this.ensureVisibleRange()).catch(err => console.error('Failed to load appointments:', err));
      }
    };

    const loadingStatusItem: ICommandBarItemProps = {
      key: 'loadingStatus',
      iconOnly: true,
      onRender: () => this.renderLoadingStatusButton()
    };

    return [
      /*{
        key: 'views',
        text: displayView.charAt(0).toUpperCase() + displayView.slice(1),
        iconProps: { iconName: 'View' },
        subMenuProps: {
          items: [
            {
              key: 'day',
              text: 'Day',
              iconProps: { iconName: 'CalendarDay' },
              canCheck: true,
              checked: displayView === 'day',
              onClick: () => this.handleViewChange('day')
            },
            {
              key: 'week',
              text: 'Week',
              iconProps: { iconName: 'CalendarWeek' },
              canCheck: true,
              checked: displayView === 'week',
              onClick: () => this.handleViewChange('week')
            },
            {
              key: 'month',
              text: 'Month',
              iconProps: { iconName: 'Calendar' },
              canCheck: true,
              checked: displayView === 'month',
              onClick: () => this.handleViewChange('month')
            },
            {
              key: 'schedule',
              text: 'Schedule',
              iconProps: { iconName: 'BulletedList' },
              canCheck: true,
              checked: displayView === 'schedule',
              onClick: () => this.handleViewChange('schedule')
            }
          ]
        }
      },*/
      {
        key: 'settings',
        iconProps: { iconName: 'Settings' },
        onClick: this.openSettingsPanel
      },
      showRefreshButton ? refreshItem : loadingStatusItem
    ];
  };

  private openSettingsPanel = (): void => {
    if (this.props.context.propertyPane.isPropertyPaneOpen()) {
      this.props.context.propertyPane.close();
    }
    this.setState({ isSettingsPanelOpen: true });
  };

  private getEnabledServiceKeys(): ServiceKey[] {
    const enabledSources = this.props.settings.sources.filter(source => source.isEnabled);
    const serviceKeys = new Set<ServiceKey>();

    serviceKeys.add('exchange');
    for (const source of enabledSources) {
      if (source.sourceType === 'ics') {
        serviceKeys.add('ics');
      } else if (source.sourceType === 'sharepoint') {
        serviceKeys.add('sharepoint');
      } else if (source.sourceType === 'planner') {
        serviceKeys.add('planner');
      } else if (source.sourceType === 'teamsShifts') {
        serviceKeys.add('teamsShifts');
      } else if (source.sourceType === 'unifiedGroup') {
        serviceKeys.add('unifiedGroup');
      }
    }

    if (this.props.settings.plannerShowAllCalendars) {
      serviceKeys.add('planner');
    }
    if (this.props.settings.teamsShiftsShowAllCalendars) {
      serviceKeys.add('teamsShifts');
    }
    if (this.props.settings.unifiedGroupShowAllCalendars) {
      serviceKeys.add('unifiedGroup');
    }

    return Array.from(serviceKeys);
  }

  private toggleLoadingStatus = (): void => {
    this.setState(prev => ({ isLoadingStatusOpen: !prev.isLoadingStatusOpen }));
  };

  private renderLoadingStatusButton(): React.ReactElement {
    const { isLoadingStatusOpen, loadingSources, loadErrors } = this.state;
    const enabledServices = this.getEnabledServiceKeys();
    const serviceLabels: Record<ServiceKey, string> = {
      exchange: 'Exchange',
      ics: 'ICS',
      sharepoint: 'SharePoint',
      planner: 'Planner',
      teamsShifts: 'Teams Shifts',
      unifiedGroup: 'Groups/Teams'
    };
    const hasLoading = enabledServices.reduce((acc, service) => acc || loadingSources[service] === 'loading', false);
    const buttonLabel = hasLoading ? 'Show loading status' : 'Show loading summary';

    return (
      <div ref={this.loadingStatusWrapperRef} className={styles.loadingStatusWrapper}>
        <CommandBarButton
          onClick={this.toggleLoadingStatus}
          ariaLabel={buttonLabel}
          iconProps={{ iconName: 'Refresh' }}
          onRenderIcon={() => (hasLoading ? <Spinner size={SpinnerSize.small} /> : <Icon iconName="CheckMark" />)}
        />
        {isLoadingStatusOpen && this.loadingStatusWrapperRef.current && (
          <Callout
            target={this.loadingStatusWrapperRef.current}
            onDismiss={() => this.setState({ isLoadingStatusOpen: false })}
            setInitialFocus
            className={styles.loadingStatusCallout}
          >
            <Text variant="medium" className={styles.loadingStatusTitle}>Loading status</Text>
            <div className={styles.loadingStatusList}>
              {enabledServices.map(service => {
                const status = loadingSources[service] ?? 'ready';
                const statusLabel = status === 'loading' ? 'Loading' : status === 'error' ? 'Error' : 'Ready';
                const errorMessage = loadErrors[service];

                return (
                  <div key={service}>
                    <div className={styles.loadingStatusRow}>
                      <div className={styles.loadingStatusLabel}>{serviceLabels[service]}</div>
                      <div className={styles.loadingStatusState}>
                        {status === 'loading' && <Icon iconName="Clock" />}
                        {status === 'ready' && <Icon iconName="CheckMark" />}
                        {status === 'error' && <Icon iconName="ErrorBadge" />}
                        <span>{statusLabel}</span>
                      </div>
                    </div>
                    {status === 'error' && errorMessage && (
                      <div className={styles.loadingStatusError}>{errorMessage}</div>
                    )}
                  </div>
                );
              })}
            </div>
          </Callout>
        )}
      </div>
    );
  }

  private getDateRangeText = (): string => {
    const { currentDate, currentView, searchQuery } = this.state;
    
    if (currentView === 'search') {
      return searchQuery ? `Search results for "${searchQuery}"` : 'Search';
    }
    
    const options: Intl.DateTimeFormatOptions = { month: 'long', year: 'numeric' };

    switch (currentView) {
      case 'day':
        return currentDate.toLocaleDateString(undefined, { ...options, day: 'numeric' });
      case 'week': {
        const weekStart = new Date(currentDate);
        if (!this.props.settings.showWeekends) {
          while (weekStart.getDay() === 0 || weekStart.getDay() === 6) weekStart.setDate(weekStart.getDate() + 1);
        }
        const weekEnd = new Date(weekStart);
        if (this.props.settings.showWeekends) {
          weekEnd.setDate(weekStart.getDate() + 6);
        } else {
          let remaining = 4;
          while (remaining > 0) {
            weekEnd.setDate(weekEnd.getDate() + 1);
            if (weekEnd.getDay() !== 0 && weekEnd.getDay() !== 6) remaining--;
          }
        }
        return `${weekStart.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })} - ${weekEnd.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' })}`;
      }
      case 'month':
      default:
        return currentDate.toLocaleDateString(undefined, options);
    }
  };

  private buildSearchIndexText(appointment: IEvent): string {
    const title = (appointment.title || '').toLowerCase();
    const location = (appointment.location || '').toLowerCase();
    return `${title} ${location}`.trim();
  }

  private renderCalendarView(): React.ReactElement {
    const { appointments, currentDate, currentView, previousView } = this.state;
    const { settings } = this.props;

    const preferredStartMinutes = settings.userPreferredStartMinutes ?? settings.preferredStartMinutes;
    const visibleHourCount = settings.userVisibleHourCount ?? settings.visibleHourCount;

    // While search is active keep rendering the last calendar layout so it
    // stays mounted and appears instantly when the user clears the query.
    const displayView = currentView === 'search' ? previousView : currentView;

    const calendarViewProps = {
      appointments,
      currentDate,
      preferredStartMinutes,
      visibleHourCount,
      slotDurationMinutes: settings.slotDurationMinutes,
      showWeekends: settings.showWeekends
    };

    switch (displayView) {
      case 'day':
        return <DayView {...calendarViewProps} />;
      case 'week':
        return <WeekView {...calendarViewProps} />;
      case 'month':
      default:
        return <MonthView appointments={appointments} currentDate={currentDate} />;
    }
  }

  private handleSearch = (query: string): void => {
    if (this.searchDebounceTimer) {
      clearTimeout(this.searchDebounceTimer);
      this.searchDebounceTimer = null;
    }

    const trimmedQuery = (query || '').trim();

    // Empty: always switch back to calendar immediately.
    if (!trimmedQuery) {
      this.applySearchQuery(query);
      return;
    }

    // >3 chars: apply immediately (enough specificity to be fast).
    if (trimmedQuery.length > 3) {
      this.applySearchQuery(query);
      return;
    }

    // 1-3 chars: debounce to avoid expensive renders on every keystroke.
    this.searchDebounceTimer = window.setTimeout(() => {
      this.searchDebounceTimer = null;
      this.applySearchQuery(query);
    }, 100);
  };

  private applySearchQuery(query: string): void {
    const trimmedQuery = (query || '').trim();

    this.setState(prevState => {
      if (trimmedQuery && prevState.currentView !== 'search') {
        // Switch to search view when search text is entered.
        return { searchQuery: query, currentView: 'search' };
      }

      if (!trimmedQuery && prevState.currentView === 'search') {
        // Switch back to previous view when search is cleared.
        return { searchQuery: '', currentView: prevState.previousView };
      }

      return { searchQuery: query, currentView: prevState.currentView };
    });
  }

  public render(): React.ReactElement<IMyCalendarsProps> {
    const { currentView, isLoading, searchQuery, appointments } = this.state;
    const isSearchMode = currentView === 'search';

    // Compute filtered appointments only when the search overlay is active.
    let filteredAppointments: IEvent[] = [];
    if (isSearchMode) {
      const trimmedQuery = (searchQuery || '').trim();
      if (trimmedQuery) {
        const query = trimmedQuery.toLowerCase();
        filteredAppointments = appointments.filter(apt => {
          const indexed = apt as IndexedEvent;
          const searchText = indexed.searchIndexText || this.buildSearchIndexText(apt);
          return searchText.includes(query);
        });
      }
    }

    return (
      <div className={styles.myCalendars}>
        <CommandBar
          items={this.getCommandBarItems()}
          farItems={this.getFarCommandBarItems()}
          className={styles.commandBar}
          styles={{
            root: { padding: 0 },
            primarySet: { padding: 0 },
            secondarySet: { padding: 0 }
          }}
        />
        {!isSearchMode && (
          <CalendarToolbar
            currentDate={this.state.currentDate}
            currentView={this.state.previousView}
            dateRangeText={this.getDateRangeText()}
            onToday={() => this.handleDateChange(new Date())}
            onNavigate={this.navigateDate}
            onDateChange={this.handleDateChange}
            onViewChange={this.handleViewChange}
          />
        )}
        <div className={styles.calendarContainer}>
          {/*
            Keep the calendar mounted at all times (hidden via CSS when searching).
            This prevents the remount cost when clearing a search query and preserves
            the active view's scroll position.
          */}
          <div style={{ display: isSearchMode ? 'none' : 'block' }}>
            {this.renderCalendarView()}
          </div>
          {isSearchMode && (
            <SearchResultsView
              appointments={filteredAppointments}
              isLoading={isLoading}
              searchQuery={searchQuery}
            />
          )}
        </div>
        <SettingsPanel
          isOpen={this.state.isSettingsPanelOpen}
          onDismiss={() => this.setState({ isSettingsPanelOpen: false })}
          settings={this.props.settings}
          onSave={(settings) => {
            this.props.onSettingsChange(settings);
            this.setState({ isSettingsPanelOpen: false });
          }}
          onReset={() => {
            if (this.props.onResetSettings) {
              this.props.onResetSettings();
            }
          }}
          httpClient={this.props.context.httpClient}
          graphClient={this.state.graphClient}
        />
      </div>
    );
  }
}
