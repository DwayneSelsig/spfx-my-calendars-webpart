import * as React from 'react';
import styles from './MyCalendars.module.scss';
import type { IMyCalendarsProps } from './IMyCalendarsProps';
import { IEvent } from '@pnp/spfx-controls-react/lib/controls/calendar/models/IEvents';
import { CalendarViewType } from '../models/ICalendarSettings';
import { ExchangeCalendarService } from '../services/ExchangeCalendarService';
import { SharePointCalendarService } from '../services/SharePointCalendarService';
import { PlannerTaskService } from '../services/PlannerTaskService';
import { TeamsShiftsService } from '../services/TeamsShiftsService';
import { UnifiedGroupCalendarService } from '../services/UnifiedGroupCalendarService';
import { DayView } from './views/DayView';
import { WeekView } from './views/WeekView';
import { MonthView } from './views/MonthView';
import { CommunityCalendarView } from './views/CommunityCalendarView';
import { SearchResultsView } from './views/SearchResultsView';
import { CommandBar, ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';
import { Callout } from '@fluentui/react/lib/Callout';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Text } from '@fluentui/react/lib/Text';
import { CommandBarButton } from '@fluentui/react/lib/Button';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { SettingsPanel } from './SettingsPanel';
//import { CalendarToolbar } from './CalendarToolbar';
import type { MSGraphClientV3 } from '@microsoft/sp-http';

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
  usePnpCalendar: boolean;
}

export default class MyCalendars extends React.Component<IMyCalendarsProps, IMyCalendarsState> {
  private debounceTimer: number | null = null;
  private searchDebounceTimer: number | null = null;
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
      showRefreshButton: true,
      usePnpCalendar: true
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

  public componentWillUnmount(): void {
    if (this.debounceTimer) {
      clearTimeout(this.debounceTimer);
      this.debounceTimer = null;
    }

    if (this.searchDebounceTimer) {
      clearTimeout(this.searchDebounceTimer);
      this.searchDebounceTimer = null;
    }

    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer);
      this.refreshTimer = null;
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

    const appendAppointments = (appointments: IEvent[]): void => {
      if (loadId !== this.activeLoadId || appointments.length === 0) {
        return;
      }

      const normalizedAppointments: IndexedEvent[] = appointments.map(apt => ({
        ...apt,
        colorHex: apt.colorHex || '#0078d4',
        // Build a reusable lowercase search index per event to avoid repeated string work while typing.
        searchIndexText: this.buildSearchIndexText(apt)
      }));

      this.setState(prev => ({
        appointments: [...prev.appointments, ...normalizedAppointments]
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
        } catch (error) {
          hadError = true;
          console.error('Failed to load user Exchange calendars:', error);
        }

        updateStatus('exchange', hadError ? 'error' : 'ready', hadError ? 'One or more Exchange calendars failed.' : undefined);
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

    if (this.props.settings.plannerShowAllCalendars) {
      tasks.push((async () => {
        let hadError = false;

        try {
          const plans = await plannerService.getUserPlans();
          const autoPlannerSources = plans.map(plan => ({
            id: `auto_planner_${plan.id}`,
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

          const appointmentsBySource = await Promise.all(autoPlannerSources.map(async source => {
            try {
              if (!source.plannerPlanId) {
                return [] as IEvent[];
              }

              return await plannerService.getTasks(
                source.plannerPlanId,
                startDate,
                endDate,
                source.plannerAssignedToMeOnly ?? false,
                source.showCompletedTasks ?? true,
                source,
                this.props.settings.plannerShowSourceLogo ?? true
              );
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
          console.error('Failed to load Planner plans for auto mode:', error);
        }

        updateStatus('planner', hadError ? 'error' : 'ready', hadError ? 'One or more Planner sources failed.' : undefined);
      })());
    } else if (sourceGroups.planner.length > 0) {
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

          return [] as IEvent[];
        }));

        const flattenedAppointments = appointmentsBySource.reduce<IEvent[]>((acc, group) => acc.concat(group), []);
        appendAppointments(flattenedAppointments);
        updateStatus('planner', hadError ? 'error' : 'ready', hadError ? 'One or more Planner sources failed.' : undefined);
      })());
    }

    if (this.props.settings.teamsShiftsShowAllCalendars) {
      tasks.push((async () => {
        let hadError = false;
        const autoSource = {
          id: 'auto_teamsShifts',
          sourceType: 'teamsShifts' as const,
          name: 'Teams Shifts',
          color: this.props.settings.organizationPrimaryColor || '#4a4fbe',
          isEnabled: true,
          showSourceLogo: this.props.settings.teamsShiftsShowSourceLogo ?? true
        };

        try {
          const events = await teamsShiftsService.getShiftsForJoinedTeams(
            startDate,
            endDate,
            autoSource,
            this.props.settings.teamsShiftsShowSourceLogo ?? true
          );
          appendAppointments(events);
        } catch (error) {
          hadError = true;
          console.error('Failed to load Teams shifts for auto mode:', error);
        }

        updateStatus('teamsShifts', hadError ? 'error' : 'ready', hadError ? 'One or more Teams Shifts sources failed.' : undefined);
      })());
    } else if (sourceGroups.teamsShifts.length > 0) {
      tasks.push((async () => {
        let hadError = false;
        const appointmentsBySource = await Promise.all(sourceGroups.teamsShifts.map(async source => {
          try {
            return await teamsShiftsService.getShiftsForJoinedTeams(
              startDate,
              endDate,
              source,
              this.props.settings.teamsShiftsShowSourceLogo ?? true
            );
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

    if (this.props.settings.unifiedGroupShowAllCalendars) {
      tasks.push((async () => {
        let hadError = false;

        try {
          const [groups, joinedTeamIds] = await Promise.all([
            unifiedGroupService.getUnifiedGroups(),
            unifiedGroupService.getJoinedTeamIds()
          ]);

          const appointmentsByGroup = await Promise.all(groups.map(async group => {
            try {
              const iconName = joinedTeamIds.has(group.id) ? 'TeamsLogo' : 'Group';
              const events = await unifiedGroupService.getGroupEvents(group.id, startDate, endDate);
              return events.map(event => ({
                ...event,
                sourceId: `auto_unifiedGroup_${group.id}`,
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
          console.error('Failed to load group calendars for auto mode:', error);
        }

        updateStatus('unifiedGroup', hadError ? 'error' : 'ready', hadError ? 'One or more group calendars failed.' : undefined);
      })());
    } else if (sourceGroups.unifiedGroup.length > 0) {
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
              return [] as IEvent[];
            }

            const iconName = joinedTeamIds.has(source.groupId) ? 'TeamsLogo' : 'Group';
            const events = await unifiedGroupService.getGroupEvents(source.groupId, startDate, endDate);
            return events.map(event => ({
              ...event,
              sourceId: source.id,
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
    const searchItem: ICommandBarItemProps = {
      key: 'search',
      onRender: () => (
        <SearchBox
          placeholder="Search appointments..."
          onChange={(_event, newValue) => this.handleSearch(newValue || '')}
          style={{
            width: 250,
            flexShrink: 0
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
        this.loadAppointments().catch(err => console.error('Failed to load appointments:', err));
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

  private buildSearchIndexText(appointment: IEvent): string {
    const title = (appointment.title || '').toLowerCase();
    const location = (appointment.location || '').toLowerCase();
    return `${title} ${location}`.trim();
  }

  private renderCalendarView(): React.ReactElement {
    const { appointments, currentDate, currentView, previousView, isLoading, usePnpCalendar } = this.state;
    const { settings } = this.props;

    const startHour = settings.userStartHour !== undefined ? settings.userStartHour : settings.startHour;
    const endHour = settings.userEndHour !== undefined ? settings.userEndHour : settings.endHour;

    // While search is active keep rendering the last calendar layout so it
    // stays mounted and appears instantly when the user clears the query.
    const displayView = currentView === 'search' ? previousView : currentView;

    const calendarViewProps = {
      appointments,
      currentDate,
      onDateChange: this.handleDateChange,
      isLoading,
      startHour,
      endHour,
      showWeekends: settings.showWeekends
    };

    const scheduleViewProps = {
      ...calendarViewProps,
      slotDuration: settings.slotDuration
    };

    switch (displayView) {
      case 'day':
        return usePnpCalendar
          ? <CommunityCalendarView {...calendarViewProps} viewType="day" />
          : <DayView {...scheduleViewProps} />;
      case 'week':
        return usePnpCalendar
          ? <CommunityCalendarView {...calendarViewProps} viewType="week" />
          : <WeekView {...scheduleViewProps} />;
      case 'month':
      default:
        return usePnpCalendar
          ? <CommunityCalendarView {...calendarViewProps} viewType="month" />
          : <MonthView {...scheduleViewProps} />;
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
        />
        <div className={styles.calendarContainer}>
          {/*
            Keep the calendar mounted at all times (hidden via CSS when searching).
            This prevents the remount cost when clearing a search query, which was
            causing noticeable lag before the PnP calendar became interactive again.
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
