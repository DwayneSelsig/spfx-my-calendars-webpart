import * as React from 'react';
import styles from './MyCalendars.module.scss';
import type { IMyCalendarsProps } from './IMyCalendarsProps';
import { IAppointment } from '../models/IAppointment';
import { CalendarViewType } from '../models/ICalendarSettings';
import { IcsParser } from '../services/IcsParser';
import { ExchangeCalendarService } from '../services/ExchangeCalendarService';
import { SharePointCalendarService } from '../services/SharePointCalendarService';
import { PlannerTaskService } from '../services/PlannerTaskService';
import { TeamsShiftsService } from '../services/TeamsShiftsService';
import { UnifiedGroupCalendarService } from '../services/UnifiedGroupCalendarService';
import { DayView } from './views/DayView';
import { WeekView } from './views/WeekView';
import { MonthView } from './views/MonthView';
import { ScheduleView } from './views/ScheduleView';
import { SearchResultsView } from './views/SearchResultsView';
import { CommandBar, ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';
import { Callout } from '@fluentui/react/lib/Callout';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Text } from '@fluentui/react/lib/Text';
import { CommandBarButton } from '@fluentui/react/lib/Button';
import { SettingsPanel } from './SettingsPanel';
import { CalendarToolbar } from './CalendarToolbar';

// MSGraphClientV3 type - using any since @microsoft/sp-client-preview is not available
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type MSGraphClientV3 = any;

type ServiceKey = 'exchange' | 'ics' | 'sharepoint' | 'planner' | 'teamsShifts' | 'unifiedGroup';
type ServiceStatus = 'loading' | 'ready' | 'error';

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

interface IMyCalendarsState {
  appointments: IAppointment[];
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
  private debounceTimer: number | null = null;
  private activeLoadId = 0;
  private loadingStatusWrapperRef = React.createRef<HTMLDivElement>();
  private refreshTimer: number | null = null;

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
    this.loadAppointments().catch(err => console.error('Failed to load appointments:', err));
    // Resolve and store graphClient for use in SettingsPanel
    this.props.context.msGraphClientFactory.getClient('3')
      .then(client => this.setState({ graphClient: client }))
      .catch(err => console.error('Failed to initialize graph client:', err));
  }

  public componentDidUpdate(prevProps: IMyCalendarsProps): void {
    try {
      const prevJson = JSON.stringify(prevProps.settings);
      const currJson = JSON.stringify(this.props.settings);
      if (prevJson !== currJson) {
        this.loadAppointments().catch(err => console.error('Failed to load appointments:', err));
      }
    } catch {
      // Fallback to sources change detection
      if (prevProps.settings.sources !== this.props.settings.sources) {
        this.loadAppointments().catch(err => console.error('Failed to load appointments:', err));
      }
    }
  }

  private async loadAppointments(): Promise<void> {
    const loadId = ++this.activeLoadId;
    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer);
      this.refreshTimer = null;
    }
    const enabledServices = this.getEnabledServiceKeys();
    const initialLoadingSources = enabledServices.reduce((acc, key) => {
      acc[key] = 'loading';
      return acc;
    }, { ...defaultLoadingSources });

    this.setState({
      isLoading: true,
      appointments: [],
      loadingSources: initialLoadingSources,
      loadErrors: { ...defaultLoadErrors },
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
    const teamsShiftsService = new TeamsShiftsService(httpClient, graphClient);
    teamsShiftsService.setGraphClient(graphClient);
    const unifiedGroupService = new UnifiedGroupCalendarService(httpClient, graphClient);
    unifiedGroupService.setGraphClient(graphClient);
    
    // Calculate date range for filtering (current month ± 3 months)
    const today = new Date();
    const startDate = new Date(today.getFullYear(), today.getMonth() - 3, 1);
    const endDate = new Date(today.getFullYear(), today.getMonth() + 4, 0);
    
    const updateStatus = (service: ServiceKey, status: ServiceStatus, errorMessage?: string): void => {
      if (loadId !== this.activeLoadId) {
        return;
      }

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

    const appendAppointments = (appointments: IAppointment[]): void => {
      if (loadId !== this.activeLoadId || appointments.length === 0) {
        return;
      }

      this.setState(prev => ({
        appointments: [...prev.appointments, ...appointments]
      }));
    };

    const enabledSources = this.props.settings.sources.filter(source => source.isEnabled);

    const sourceGroups: Record<ServiceKey, typeof enabledSources> = {
      exchange: [],
      ics: enabledSources.filter(source => source.sourceType === 'ics'),
      sharepoint: enabledSources.filter(source => source.sourceType === 'sharepoint'),
      planner: enabledSources.filter(source => source.sourceType === 'planner'),
      teamsShifts: enabledSources.filter(source => source.sourceType === 'teamsShifts'),
      unifiedGroup: enabledSources.filter(source => source.sourceType === 'unifiedGroup')
    };

    const tasks: Array<Promise<void>> = [];

    // Load Exchange calendars (user's own calendars)
    if (enabledServices.indexOf('exchange') >= 0) {
      tasks.push((async () => {
        let hadError = false;

        try {
          const userCalendars = await exchangeService.getCalendars();
          const exchangeCalendarStates = this.props.settings.exchangeCalendarStates || {};
          const calendarPromises = userCalendars
            .filter(calendar => exchangeCalendarStates[calendar.id] !== false)
            .map(async calendar => {
              try {
                const events = await exchangeService.getCalendarEvents(
                  calendar.id,
                  startDate,
                  endDate,
                  undefined // current user
                );
                return events.map(event => ({
                  ...event,
                  sourceId: `exchange_${calendar.id}`,
                  color: calendar.hexColor,
                  sourceType: 'exchange' as const,
                  showSourceLogo: this.props.settings.exchangeShowSourceLogo ?? true
                }));
              } catch (error) {
                hadError = true;
                console.error(`Failed to load Exchange calendar ${calendar.name}:`, error);
                return [] as IAppointment[];
              }
            });

          const appointmentGroups = await Promise.all(calendarPromises);
          const flattenedAppointments = appointmentGroups.reduce((acc, group) => acc.concat(group), [] as IAppointment[]);
          appendAppointments(flattenedAppointments);
        } catch (error) {
          hadError = true;
          console.error('Failed to load user Exchange calendars:', error);
        }

        updateStatus('exchange', hadError ? 'error' : 'ready', hadError ? 'One or more Exchange calendars failed.' : undefined);
      })());
    }

    if (sourceGroups.ics.length > 0) {
      tasks.push((async () => {
        let hadError = false;
        const appointmentsBySource = await Promise.all(sourceGroups.ics.map(async source => {
          try {
            let appointments: IAppointment[] = [];
            if (source.rawContent) {
              appointments = IcsParser.parseRawContent(source.rawContent, source.id, source.color);
            } else if (source.url) {
              appointments = await IcsParser.fetchAndParse(source.url, source.id, source.color, httpClient);
            }

            return appointments.map(apt => ({
              ...apt,
              sourceType: 'ics' as const,
              showSourceLogo: source.showSourceLogo ?? true
            }));
          } catch (error) {
            hadError = true;
            console.error(`Failed to load ICS calendar ${source.name}:`, error);
            return [] as IAppointment[];
          }
        }));

        const flattenedAppointments = appointmentsBySource.reduce((acc, group) => acc.concat(group), [] as IAppointment[]);
        appendAppointments(flattenedAppointments);
        updateStatus('ics', hadError ? 'error' : 'ready', hadError ? 'One or more ICS sources failed.' : undefined);
      })());
    }

    if (sourceGroups.sharepoint.length > 0) {
      tasks.push((async () => {
        let hadError = false;
        const appointmentsBySource = await Promise.all(sourceGroups.sharepoint.map(async source => {
          try {
            if (source.sharePointSiteId && source.sharePointListId) {
              const items = await sharePointService.getListEvents(
                source.sharePointSiteId,
                source.sharePointListId,
                startDate,
                endDate,
                source.sharePointFieldMapping
              );
              return items.map(item => ({
                ...item,
                sourceId: source.id,
                color: source.color,
                sourceType: 'sharepoint' as const,
                showSourceLogo: this.props.settings.sharePointShowSourceLogo ?? true
              }));
            }
          } catch (error) {
            hadError = true;
            console.error(`Failed to load SharePoint calendar ${source.name}:`, error);
          }

          return [] as IAppointment[];
        }));

        const flattenedAppointments = appointmentsBySource.reduce((acc, group) => acc.concat(group), [] as IAppointment[]);
        appendAppointments(flattenedAppointments);
        updateStatus('sharepoint', hadError ? 'error' : 'ready', hadError ? 'One or more SharePoint calendars failed.' : undefined);
      })());
    }

    if (sourceGroups.planner.length > 0) {
      tasks.push((async () => {
        let hadError = false;
        const appointmentsBySource = await Promise.all(sourceGroups.planner.map(async source => {
          try {
            if (source.plannerPlanId) {
              return await plannerService.getTasks(
                source.plannerPlanId,
                startDate,
                endDate,
                source.plannerAssignedToMeOnly ?? false,
                source.showCompletedTasks ?? true,
                source,
                this.props.settings.plannerShowSourceLogo ?? true
              );
            }
          } catch (error) {
            hadError = true;
            console.error(`Failed to load Planner tasks ${source.name}:`, error);
          }

          return [] as IAppointment[];
        }));

        const flattenedAppointments = appointmentsBySource.reduce((acc, group) => acc.concat(group), [] as IAppointment[]);
        appendAppointments(flattenedAppointments);
        updateStatus('planner', hadError ? 'error' : 'ready', hadError ? 'One or more Planner sources failed.' : undefined);
      })());
    }

    if (sourceGroups.teamsShifts.length > 0) {
      tasks.push((async () => {
        let hadError = false;
        const appointmentsBySource = await Promise.all(sourceGroups.teamsShifts.map(async source => {
          try {
            return await teamsShiftsService.getShiftsForJoinedTeams(
              startDate,
              endDate,
              source,
              source.showSourceLogo ?? true
            );
          } catch (error) {
            hadError = true;
            console.error(`Failed to load Teams shifts ${source.name}:`, error);
            return [] as IAppointment[];
          }
        }));

        const flattenedAppointments = appointmentsBySource.reduce((acc, group) => acc.concat(group), [] as IAppointment[]);
        appendAppointments(flattenedAppointments);
        updateStatus('teamsShifts', hadError ? 'error' : 'ready', hadError ? 'One or more Teams Shifts sources failed.' : undefined);
      })());
    }

    if (sourceGroups.unifiedGroup.length > 0) {
      tasks.push((async () => {
        let hadError = false;
        let joinedTeamIds: Set<string> = new Set();

        try {
          joinedTeamIds = await unifiedGroupService.getJoinedTeamIds();
        } catch (error) {
          hadError = true;
          console.error('Failed to load joined Teams for group calendars:', error);
        }

        const appointmentsBySource = await Promise.all(sourceGroups.unifiedGroup.map(async source => {
          try {
            if (!source.groupId) {
              return [] as IAppointment[];
            }

            const iconName = joinedTeamIds.has(source.groupId) ? 'TeamsLogo' : 'Group';
            const events = await unifiedGroupService.getGroupEvents(source.groupId, startDate, endDate);
            return events.map(event => ({
              ...event,
              sourceId: source.id,
              color: source.color,
              sourceType: 'unifiedGroup' as const,
              showSourceLogo: source.showSourceLogo ?? true,
              sourceIconName: iconName
            }));
          } catch (error) {
            hadError = true;
            console.error(`Failed to load group calendar ${source.name}:`, error);
            return [] as IAppointment[];
          }
        }));

        const flattenedAppointments = appointmentsBySource.reduce((acc, group) => acc.concat(group), [] as IAppointment[]);
        appendAppointments(flattenedAppointments);
        updateStatus('unifiedGroup', hadError ? 'error' : 'ready', hadError ? 'One or more group calendars failed.' : undefined);
      })());
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
    this.setState({ currentDate: date });
  };

  private handleViewChange = (view: CalendarViewType): void => {
    if (view !== 'search') {
      this.setState({ currentView: view, previousView: view });
      
      const newSettings = { ...this.props.settings, defaultView: view };
      this.debouncedSettingsChange(newSettings);
    }
  };

  private debouncedSettingsChange = (settings: import('../models/ICalendarSettings').ICalendarSettings): void => {
    if (this.debounceTimer) {
      clearTimeout(this.debounceTimer);
    }
    
    this.debounceTimer = window.setTimeout(() => {
      this.props.onSettingsChange(settings);
    }, 500);
  };

  private navigateDate = (direction: 'prev' | 'next'): void => {
    const { currentDate, currentView } = this.state;
    const newDate = new Date(currentDate);

    switch (currentView) {
      case 'day':
      case 'schedule':
        newDate.setDate(newDate.getDate() + (direction === 'next' ? 1 : -1));
        break;
      case 'week':
        newDate.setDate(newDate.getDate() + (direction === 'next' ? 7 : -7));
        break;
      case 'month':
        newDate.setMonth(newDate.getMonth() + (direction === 'next' ? 1 : -1));
        break;
    }

    this.setState({ currentDate: newDate });
  };

  private getCommandBarItems = (): ICommandBarItemProps[] => {
    const { currentView } = this.state;
    const isSearchView = currentView === 'search';

    return [
      {
        key: 'today',
        text: 'Today',
        iconProps: { iconName: 'GotoToday' },
        disabled: isSearchView,
        onClick: () => this.setState({ currentDate: new Date() })
      },
      {
        key: 'prev',
        iconProps: { iconName: 'ChevronLeft' },
        disabled: isSearchView,
        onClick: () => this.navigateDate('prev')
      },
      {
        key: 'next',
        iconProps: { iconName: 'ChevronRight' },
        disabled: isSearchView,
        onClick: () => this.navigateDate('next')
      },
      {
        key: 'dateRange',
        text: this.getDateRangeText(),
        disabled: true
      }
    ];
  };

  private getFarCommandBarItems = (): ICommandBarItemProps[] => {
    const { currentView, previousView, showRefreshButton } = this.state;
    const displayView = currentView === 'search' ? previousView : currentView;

    const refreshItem: ICommandBarItemProps = {
      key: 'refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: () => {
        this.loadAppointments().catch(err => console.error('Failed to load appointments:', err));
      }
    };

    const loadingStatusItem: ICommandBarItemProps = {
      key: 'loadingStatus',
      iconOnly: true,
      onRender: () => this.renderLoadingStatusButton()
    };

    return [
      {
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
      },
      {
        key: 'settings',
        iconProps: { iconName: 'Settings' },
        onClick: () => this.setState({ isSettingsPanelOpen: true })
      },
      showRefreshButton ? refreshItem : loadingStatusItem
    ];
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
      case 'schedule':
        return currentDate.toLocaleDateString(undefined, { ...options, day: 'numeric' });
      case 'week': {
        const weekStart = new Date(currentDate);
        const day = weekStart.getDay();
        const diff = day === 0 ? -6 : 1 - day;
        weekStart.setDate(weekStart.getDate() + diff);
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekStart.getDate() + 6);
        return `${weekStart.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })} - ${weekEnd.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' })}`;
      }
      case 'month':
      default:
        return currentDate.toLocaleDateString(undefined, options);
    }
  };

  private renderView(): React.ReactElement {
    const { appointments, currentDate, currentView, isLoading, searchQuery } = this.state;
    const { settings } = this.props;

    // Filter appointments based on search query
    let filteredAppointments = appointments;
    const trimmedQuery = (searchQuery || '').trim();
    if (trimmedQuery) {
      const query = trimmedQuery.toLowerCase();
      filteredAppointments = appointments.filter(apt =>
        apt.title.toLowerCase().includes(query) ||
        (apt.location && apt.location.toLowerCase().includes(query)) ||
        (apt.organizer && apt.organizer.toLowerCase().includes(query))
      );
    }

    // If in search view, show search results
    if (currentView === 'search') {
      return (
        <SearchResultsView
          appointments={filteredAppointments}
          isLoading={isLoading}
          searchQuery={searchQuery}
        />
      );
    }

    // Use user-level overrides if set, otherwise fall back to webpart-level settings
    const startHour = settings.userStartHour !== undefined ? settings.userStartHour : settings.startHour;
    const endHour = settings.userEndHour !== undefined ? settings.userEndHour : settings.endHour;

    const viewProps = {
      appointments: filteredAppointments,
      currentDate,
      onDateChange: this.handleDateChange,
      isLoading,
      startHour: startHour,
      endHour: endHour,
      showWeekends: settings.showWeekends,
      slotDuration: settings.slotDuration
    };

    switch (currentView) {
      case 'day':
        return <DayView {...viewProps} />;
      case 'week':
        return <WeekView {...viewProps} />;
      case 'month':
        return <MonthView {...viewProps} />;
      case 'schedule':
        return <ScheduleView {...viewProps} />;
      default:
        return <MonthView {...viewProps} />;
    }
  }

  private handleSearch = (query: string): void => {
    const trimmedQuery = (query || '').trim();
    
    if (trimmedQuery && this.state.currentView !== 'search') {
      // Switch to search view when search text is entered
      this.setState({ searchQuery: query, currentView: 'search' });
    } else if (!trimmedQuery && this.state.currentView === 'search') {
      // Switch back to previous view when search is cleared
      this.setState({ searchQuery: '', currentView: this.state.previousView });
    } else {
      // Just update search query
      this.setState({ searchQuery: query });
    }
  };

  public render(): React.ReactElement<IMyCalendarsProps> {
    const { currentDate, currentView, previousView } = this.state;
    const displayView = (currentView === 'search' ? previousView : currentView) as 'day' | 'week' | 'month' | 'schedule';

    return (
      <div className={styles.myCalendars}>
        <CalendarToolbar
          currentDate={currentDate}
          viewType={displayView}
          onToday={() => this.setState({ currentDate: new Date() })}
          onDateChange={this.handleDateChange}
          onNavigate={this.navigateDate}
          onSearch={this.handleSearch}
          dateRangeText={this.getDateRangeText()}
        />
        <CommandBar
          items={[]}
          farItems={this.getFarCommandBarItems()}
          className={styles.commandBar}
        />
        <div className={styles.calendarContainer}>
          {this.renderView()}
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