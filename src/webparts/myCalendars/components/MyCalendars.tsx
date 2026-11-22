import * as React from 'react';
import type { IMyCalendarsProps } from './IMyCalendarsProps';
import type { ICalendarEvent as IEvent } from '../models/ICalendarEvent';
import type { CalendarViewType, ICalendarSettings } from '../models/ICalendarSettings';
import { CalendarRangeCoordinator, calendarMonth, eventIntersectsRange } from '../services/CalendarRangeCoordinator';
import { type CalendarCacheServiceKey, type ICalendarEventCacheConfiguration, normalizeCacheDuration } from '../services/CalendarEventCache';
import { startCalendarSources } from '../services/CalendarSources';
import { IndexedCalendarEventCache } from '../services/IndexedCalendarEventCache';
import { addLocalDays, getMonthVisibleDays, getWeekVisibleDays, startOfLocalDay } from './views/calendarUtils';
import { DayView } from './views/DayView';
import { WeekView } from './views/WeekView';
import { MonthView } from './views/MonthView';
import { SearchResultsView } from './views/SearchResultsView';
import { CalendarToolbar } from './CalendarToolbar';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Callout } from '@fluentui/react/lib/Callout';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Text } from '@fluentui/react/lib/Text';
import { CommandBarButton } from '@fluentui/react/lib/Button';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { UserSettingsPanel } from './UserSettingsPanel';
import { getSourceIconName, getSourceTypeDisplayName } from '../utils/sourceIconHelper';
import { formatCalendarDate } from './views/calendarFormatting';
import * as strings from 'MyCalendarsWebPartStrings';
import { formatLocalizedString } from '../utils/localization';

type ServiceKey = CalendarCacheServiceKey;
type ServiceStatus = 'loading' | 'ready' | 'error';
type IndexedEvent = IEvent & { searchIndexText?: string };
type ToolbarLayout = 'wide' | 'compact' | 'stacked';

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

const serviceErrorLabels: Record<ServiceKey, string | undefined> = {
  exchange: strings.ExchangeLoadErrorLabel,
  ics: undefined,
  sharepoint: strings.SharePointLoadErrorLabel,
  planner: strings.PlannerLoadErrorLabel,
  teamsShifts: strings.TeamsShiftsLoadErrorLabel,
  unifiedGroup: strings.GroupsTeamsLoadErrorLabel
};

const CalendarDay = React.memo(DayView);
const CalendarWeek = React.memo(WeekView);
const CalendarMonth = React.memo(MonthView);

const styles = mergeStyleSets({
  myCalendars: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    color: 'var(--bodyText, #323130)'
  },
  topToolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: 8,
    flexShrink: 0,
    borderBottom: '1px solid var(--neutralLight, #edebe9)',
    paddingTop: 8,
    paddingBottom: 10
  },
  toolbarSearch: {
    flex: '1 1 0',
    minWidth: 0
  },
  toolbarActions: {
    display: 'flex',
    flexShrink: 0,
    alignItems: 'center',
    selectors: {
      '.ms-Button': { height: 44 }
    }
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
  settingsWriteError?: string;
  searchQuery: string;
  appliedSearchQuery: string;
  toolbarLayout: ToolbarLayout;
  visibleLoadErrors: string[];
  loadingSources: Record<ServiceKey, ServiceStatus>;
  loadErrors: Record<ServiceKey, string | undefined>;
  isLoadingStatusOpen: boolean;
  showRefreshButton: boolean;
}

export default class MyCalendars extends React.Component<IMyCalendarsProps, IMyCalendarsState> {
  private searchDebounceTimer: number | null = null;
  private topToolbarRef = React.createRef<HTMLDivElement>();
  private toolbarResizeObserver: ResizeObserver | undefined;
  private loadingStatusWrapperRef = React.createRef<HTMLDivElement>();
  private refreshingAdminSources = false;
  private refreshTimer: number | null = null;
  private coordinator?: CalendarRangeCoordinator;
  private loadingFingerprint?: string;
  private appointmentCache?: IndexedCalendarEventCache;
  private updateFrame?: number;
  private readonly prepared = new WeakMap<IEvent, IndexedEvent>();
  private visibleAppointments: IEvent[] = [];

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
      appliedSearchQuery: '',
      toolbarLayout: 'compact',
      visibleLoadErrors: [],
      loadingSources: { ...defaultLoadingSources },
      loadErrors: { ...defaultLoadErrors },
      isLoadingStatusOpen: false,
      showRefreshButton: true
    };
  }

  public componentDidMount(): void {
    this.settingsUiMounted = true;
    if (this.topToolbarRef.current) {
      this.toolbarResizeObserver = new ResizeObserver(entries => {
        const width = entries[0]?.contentRect.width;
        if (width === undefined) return;
        const toolbarLayout: ToolbarLayout = width >= 600 ? 'wide' : width >= 360 ? 'compact' : 'stacked';
        if (toolbarLayout !== this.state.toolbarLayout) this.setState({ toolbarLayout });
      });
      this.toolbarResizeObserver.observe(this.topToolbarRef.current);
    }
    this.loadAppointments().catch(err => console.error('Failed to load appointments:', err));
  }

  private getLoadingSettingsFingerprint(settings: ICalendarSettings): string {
    const loadingSettings = { ...settings } as unknown as Record<string, unknown>;
    ['defaultView', 'preferredStartMinutes', 'visibleHourCount', 'slotDurationMinutes', 'showWeekends', 'adminShowWeekends', 'userShowWeekends', 'userPreferredStartMinutes', 'userVisibleHourCount', 'maxConcurrentGraphRequests']
      .forEach(key => delete loadingSettings[key]);
    return JSON.stringify(loadingSettings);
  }

  private getCacheConfiguration(settings: ICalendarSettings): ICalendarEventCacheConfiguration {
    return {
      tenantId: this.props.tenantId,
      userId: this.props.userId,
      webPartInstanceId: this.props.webPartInstanceId,
      configSignature: this.getCacheSignature(settings)
    };
  }

  private getCacheSignature(settings: ICalendarSettings): string {
    const loadingSettings = { ...settings } as unknown as Record<string, unknown>;
    [
      'defaultView', 'preferredStartMinutes', 'visibleHourCount', 'slotDurationMinutes', 'showWeekends',
      'adminShowWeekends', 'userShowWeekends', 'userPreferredStartMinutes', 'userVisibleHourCount',
      'enableCache', 'cacheDurationMinutes', 'availableAdminIcsCatalogItems', 'maxConcurrentGraphRequests'
    ].forEach(key => delete loadingSettings[key]);
    return JSON.stringify(loadingSettings);
  }

  public componentDidUpdate(prevProps: IMyCalendarsProps): void {
    if (this.refreshingAdminSources) return;
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
        if ((this.loadingFingerprint || this.getLoadingSettingsFingerprint(prevProps.settings)) !== this.getLoadingSettingsFingerprint(this.props.settings)) {
          this.loadAppointments().catch(err => console.error('Failed to load appointments:', err));
        } else if (!defaultViewChanged && prevProps.settings.showWeekends !== this.props.settings.showWeekends) {
          this.ensureVisibleRange().catch(err => console.error('Failed to load visible range:', err));
        }
      }
    } catch {
      // Fallback to sources change detection
      if (prevProps.settings.sources !== this.props.settings.sources) {
        this.loadAppointments().catch(err => console.error('Failed to load appointments:', err));
      }
    }
  }

  public componentWillUnmount(): void {
    this.settingsUiMounted = false;
    this.toolbarResizeObserver?.disconnect();
    this.coordinator?.dispose();
    this.appointmentCache?.close();
    if (this.updateFrame !== undefined) window.cancelAnimationFrame(this.updateFrame);
    if (this.searchDebounceTimer) {
      clearTimeout(this.searchDebounceTimer);
      this.searchDebounceTimer = null;
    }

    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer);
      this.refreshTimer = null;
    }
  }

  private getVisibleDays(date: Date, view: CalendarViewType, showWeekends = this.props.settings.showWeekends): Date[] {
    if (view === 'month') return getMonthVisibleDays(date);
    if (view === 'week') return getWeekVisibleDays(date, showWeekends);
    return [startOfLocalDay(date)];
  }

  private getVisibleMonths(date: Date, view: CalendarViewType, showWeekends?: boolean): number[] {
    return Array.from(new Set([calendarMonth(date), ...this.getVisibleDays(date, view, showWeekends).map(calendarMonth)]));
  }

  private async ensureVisibleRange(date = this.state.currentDate, view = this.state.previousView): Promise<void> {
    this.coordinator?.setVisible(this.getVisibleMonths(date, view), true);
  }

  private async loadAppointments(force = false, settings = this.props.settings): Promise<void> {
    const graph = this.props.graph;
    if (!graph || !this.settingsUiMounted) return;
    const previous = this.coordinator;
    previous?.dispose();
    this.appointmentCache?.close();
    if (this.refreshTimer !== null) window.clearTimeout(this.refreshTimer);
    this.refreshTimer = null;
    const enableCache = settings.enableCache;
    const cache = new IndexedCalendarEventCache(this.getCacheConfiguration(settings), enableCache);
    this.appointmentCache = cache;
    const coordinator = new CalendarRangeCoordinator({
      today: new Date(),
      cacheDurationMs: normalizeCacheDuration(settings.cacheDurationMinutes) * 60000,
      cache: enableCache ? cache : undefined,
      force,
      previous: force && enableCache ? previous : undefined,
      onChange: () => this.queueCalendarUpdate()
    });
    this.coordinator = coordinator;
    this.loadingFingerprint = this.getLoadingSettingsFingerprint(settings);
    this.setState({
      appointments: force && enableCache ? this.state.appointments : [],
      isLoading: true,
      loadingSources: { ...defaultLoadingSources },
      loadErrors: { ...defaultLoadErrors },
      visibleLoadErrors: [],
      showRefreshButton: false,
      isLoadingStatusOpen: false
    });
    coordinator.setVisible(this.getVisibleMonths(this.state.currentDate, this.state.previousView, settings.showWeekends), force);
    startCalendarSources(coordinator, graph, settings, this.props.tenantId);
    coordinator.warmVisibleCache().catch(err => console.error('Failed to read visible appointments from the cache:', err));
  }

  private prepare(event: IEvent): IndexedEvent {
    const cached = this.prepared.get(event);
    if (cached) return cached;
    const prepared = {
      ...event,
      colorHex: event.colorHex || '#0078d4',
      sourceIconName: getSourceIconName(event.sourceType, event.sourceIconName),
      sourceDisplayName: event.sourceDisplayName || getSourceTypeDisplayName(event.sourceType, event.sourceIconName),
      searchIndexText: this.buildSearchIndexText(event)
    };
    this.prepared.set(event, prepared);
    return prepared;
  }

  private queueCalendarUpdate(): void {
    if (!this.settingsUiMounted || this.updateFrame !== undefined) return;
    this.updateFrame = window.requestAnimationFrame(() => {
      this.updateFrame = undefined;
      const coordinator = this.coordinator;
      if (!coordinator) return;
      const status = coordinator.status();
      const loadingSources = { ...defaultLoadingSources };
      const loadErrors = { ...defaultLoadErrors };
      const visibleLoadErrors: string[] = [];
      status.pending.forEach(service => { loadingSources[service] = 'loading'; });
      status.failures.forEach(service => {
        loadingSources[service] = 'error';
        loadErrors[service] = serviceErrorLabels[service];
      });
      status.visibleFailures.forEach(service => {
        const message = serviceErrorLabels[service];
        if (message) visibleLoadErrors.push(message);
      });
      const exchangeErrors = this.props.settings.adminExchangeDiscoveryErrors || [];
      if (exchangeErrors.length) {
        loadingSources.exchange = 'error';
        loadErrors.exchange = [strings.ExchangeLoadErrorLabel, ...exchangeErrors].join(' — ');
        const index = visibleLoadErrors.indexOf(strings.ExchangeLoadErrorLabel);
        if (index >= 0) visibleLoadErrors[index] = loadErrors.exchange;
        else visibleLoadErrors.push(loadErrors.exchange);
      }
      const appointments = coordinator.appointments().map(event => this.prepare(event));
      if (status.pending.size > 0 || this.refreshingAdminSources) {
        if (this.refreshTimer !== null) window.clearTimeout(this.refreshTimer);
        this.refreshTimer = null;
      } else if (!this.state.showRefreshButton && this.refreshTimer === null) {
        this.refreshTimer = window.setTimeout(() => {
          this.refreshTimer = null;
          if (this.coordinator === coordinator) this.setState({ showRefreshButton: true });
        }, 2000);
      }
      this.setState(prev => this.coordinator === coordinator ? {
        appointments: this.sameItems(prev.appointments, appointments) ? prev.appointments : appointments,
        isLoading: status.background,
        loadingSources,
        loadErrors,
        visibleLoadErrors,
        showRefreshButton: status.pending.size === 0 && !this.refreshingAdminSources && prev.showRefreshButton
      } : null);
    });
  }

  private sameItems<T>(a: T[], b: T[]): boolean {
    return a.length === b.length && a.every((item, index) => item === b[index]);
  }

  private handleDateChange = (date: Date): void => {
    this.setState({ currentDate: date }, () => {
      if (this.state.currentDate === date) this.ensureVisibleRange(date).catch(err => console.error('Failed to load visible range:', err));
    });
  };

  private settingsUiMounted = false;
  private viewSavePending = false;

  private handleViewChange = async (view: CalendarViewType): Promise<void> => {
    if (view === 'search' || this.viewSavePending || this.props.isSettingsWritePending) return;
    this.viewSavePending = true;
    this.setState({ currentView: view, previousView: view, settingsWriteError: undefined }, () => {
      this.ensureVisibleRange(this.state.currentDate, view).catch(err => console.error('Failed to load visible range:', err));
    });
    try {
      await this.props.onDefaultViewChange(view);
    } catch (error) {
      console.error('Failed to save the personal default calendar view.', error);
      const confirmedView = this.props.settings.defaultView;
      if (this.settingsUiMounted) this.setState(prev => ({
        currentView: prev.currentView === 'search' ? 'search' : confirmedView,
        previousView: confirmedView,
        settingsWriteError: strings.UserSettingsSaveErrorLabel
      }), () => { this.ensureVisibleRange().catch(err => console.error('Failed to restore visible range:', err)); });
    } finally {
      this.viewSavePending = false;
    }
  };

  private handleSettingsDismiss = (): void => {
    if (this.props.isSettingsWritePending) return;
    this.props.onCancelSettingsPreview();
    this.setState({ isSettingsPanelOpen: false });
  };

  private handleManualRefresh = async (): Promise<void> => {
    if (this.refreshingAdminSources) return;
    this.refreshingAdminSources = true;
    this.setState({ showRefreshButton: false });
    try {
      const settings = await this.props.onRefreshAdminSources?.();
      await this.loadAppointments(true, settings || this.props.settings);
    } finally {
      this.refreshingAdminSources = false;
      this.queueCalendarUpdate();
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

  private renderToolbarButton(iconName: string, label: string, onClick: () => void, disabled = false): React.ReactElement {
    const iconOnly = this.state.toolbarLayout !== 'wide';
    return (
      <TooltipHost content={iconOnly ? label : undefined} setAriaDescribedBy={false}>
        <CommandBarButton
          text={iconOnly ? undefined : label}
          ariaLabel={label}
          iconProps={{ iconName }}
          onClick={onClick} disabled={disabled}
        />
      </TooltipHost>
    );
  }

  private openSettingsPanel = (): void => {
    if (this.props.isSettingsWritePending || this.viewSavePending) return;
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
      exchange: strings.ServiceExchangeLabel,
      ics: strings.ServiceIcsLabel,
      sharepoint: strings.ServiceSharePointLabel,
      planner: strings.ServicePlannerLabel,
      teamsShifts: strings.ServiceTeamsShiftsLabel,
      unifiedGroup: strings.ServiceGroupsTeamsLabel
    };
    const hasLoading = enabledServices.reduce((acc, service) => acc || loadingSources[service] === 'loading', false);
    const buttonLabel = hasLoading ? strings.ShowLoadingStatusLabel : strings.ShowLoadingSummaryLabel;

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
            <Text variant="medium" className={styles.loadingStatusTitle}>{strings.LoadingStatusLabel}</Text>
            <div className={styles.loadingStatusList}>
              {enabledServices.map(service => {
                const status = loadingSources[service] ?? 'ready';
                const statusLabel = status === 'loading' ? strings.LoadingLabel : status === 'error' ? strings.ErrorLabel : strings.ReadyLabel;
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
    const { currentDate, currentView, appliedSearchQuery } = this.state;
    
    if (currentView === 'search') {
      return appliedSearchQuery ? formatLocalizedString(strings.SearchResultsForLabel, appliedSearchQuery) : strings.SearchLabel;
    }
    
    const options: Intl.DateTimeFormatOptions = { month: 'long', year: 'numeric' };

    switch (currentView) {
      case 'day':
        return formatCalendarDate(currentDate, { ...options, day: 'numeric' }, this.props.locale);
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
        return `${formatCalendarDate(weekStart, { month: 'short', day: 'numeric' }, this.props.locale)} - ${formatCalendarDate(weekEnd, { month: 'short', day: 'numeric', year: 'numeric' }, this.props.locale)}`;
      }
      case 'month':
      default:
        return formatCalendarDate(currentDate, options, this.props.locale);
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

    const days = this.getVisibleDays(currentDate, displayView);
    const visible = appointments.filter(event => eventIntersectsRange(event, {
      start: days[0], end: addLocalDays(days[days.length - 1], 1)
    }));
    if (!this.sameItems(this.visibleAppointments, visible)) this.visibleAppointments = visible;
    const calendarViewProps = {
      appointments: this.visibleAppointments,
      currentDate,
      preferredStartMinutes,
      visibleHourCount,
      slotDurationMinutes: settings.slotDurationMinutes,
      showWeekends: settings.showWeekends,
      locale: this.props.locale
    };

    switch (displayView) {
      case 'day':
        return <CalendarDay {...calendarViewProps} />;
      case 'week':
        return <CalendarWeek {...calendarViewProps} />;
      case 'month':
      default:
        return <CalendarMonth appointments={this.visibleAppointments} currentDate={currentDate} locale={this.props.locale} />;
    }
  }

  private handleSearch = (query: string): void => {
    if (this.searchDebounceTimer) {
      clearTimeout(this.searchDebounceTimer);
      this.searchDebounceTimer = null;
    }

    // Keep controlled input immediate; debounce only the applied result query.
    this.setState({ searchQuery: query });
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
        return { appliedSearchQuery: query, currentView: 'search' };
      }

      if (!trimmedQuery && prevState.currentView === 'search') {
        // Switch back to previous view when search is cleared.
        return { appliedSearchQuery: '', currentView: prevState.previousView };
      }

      return { appliedSearchQuery: query, currentView: prevState.currentView };
    });
  }

  public render(): React.ReactElement<IMyCalendarsProps> {
    const { currentView, isLoading, searchQuery, appliedSearchQuery, appointments, toolbarLayout, showRefreshButton } = this.state;
    const isSearchMode = currentView === 'search';

    // Compute filtered appointments only when the search overlay is active.
    let filteredAppointments: IEvent[] = [];
    if (isSearchMode) {
      const trimmedQuery = appliedSearchQuery.trim();
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
        <div
          ref={this.topToolbarRef}
          className={styles.topToolbar}
          style={{ flexDirection: toolbarLayout === 'stacked' ? 'column' : 'row' }}
        >
          <div
            className={styles.toolbarSearch}
            style={toolbarLayout === 'stacked' ? { order: 1, width: '100%', flex: '0 0 auto' } : undefined}
          >
            <SearchBox
              value={searchQuery}
              placeholder={strings.SearchAppointmentsPlaceholder}
              ariaLabel={strings.SearchAppointmentsLabel}
              onChange={(_event, newValue) => this.handleSearch(newValue || '')}
              styles={{ root: { width: '100%' } }}
            />
          </div>
          <div
            className={styles.toolbarActions}
            style={toolbarLayout === 'stacked' ? { alignSelf: 'flex-end' } : undefined}
          >
            {this.renderToolbarButton('Settings', strings.SettingsLabel, this.openSettingsPanel, this.props.isSettingsWritePending)}
            {showRefreshButton
              ? this.renderToolbarButton('Refresh', strings.RefreshLabel, () => {
                this.handleManualRefresh().catch(err => console.error('Failed to refresh appointments:', err));
              })
              : this.renderLoadingStatusButton()}
          </div>
        </div>
        {!isSearchMode && (
          <CalendarToolbar
            isViewChangeDisabled={this.props.isSettingsWritePending}
            currentDate={this.state.currentDate}
            currentView={this.state.previousView}
            dateRangeText={this.getDateRangeText()}
            onToday={() => this.handleDateChange(new Date())}
            onNavigate={this.navigateDate}
            onDateChange={this.handleDateChange}
            onViewChange={this.handleViewChange}
          />
        )}
        {this.state.settingsWriteError && <MessageBar messageBarType={MessageBarType.error}>{this.state.settingsWriteError}</MessageBar>}
        {this.state.visibleLoadErrors.length > 0 && (
          <MessageBar messageBarType={MessageBarType.warning}>{this.state.visibleLoadErrors.join(' — ')}</MessageBar>
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
              searchQuery={appliedSearchQuery}
              locale={this.props.locale}
            />
          )}
        </div>
        <UserSettingsPanel
          isOpen={this.state.isSettingsPanelOpen}
          onDismiss={this.handleSettingsDismiss}
          settings={this.props.settings}
          onPreview={this.props.onPreviewSettings}
          onSave={this.props.onSettingsChange}
          onReset={this.props.onResetSettings}
          httpClient={this.props.context.httpClient}
          graphClient={this.props.settingsGraph}
          webAbsoluteUrl={this.props.context.pageContext.web.absoluteUrl}
          locale={this.props.locale}
        />
      </div>
    );
  }
}
