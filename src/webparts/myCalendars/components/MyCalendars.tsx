import * as React from 'react';
import styles from './MyCalendars.module.scss';
import type { IMyCalendarsProps } from './IMyCalendarsProps';
import { IAppointment } from '../models/IAppointment';
import { CalendarViewType } from '../models/ICalendarSettings';
import { IcsParser } from '../services/IcsParser';
import { ExchangeCalendarService } from '../services/ExchangeCalendarService';
import { SharePointCalendarService } from '../services/SharePointCalendarService';
import { DayView } from './views/DayView';
import { WeekView } from './views/WeekView';
import { MonthView } from './views/MonthView';
import { ScheduleView } from './views/ScheduleView';
import { SearchResultsView } from './views/SearchResultsView';
import { CommandBar, ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';
import { SettingsPanel } from './SettingsPanel';
import { CalendarToolbar } from './CalendarToolbar';

interface IMyCalendarsState {
  appointments: IAppointment[];
  currentDate: Date;
  currentView: CalendarViewType;
  previousView: CalendarViewType;
  isLoading: boolean;
  isSettingsPanelOpen: boolean;
  searchQuery: string;
}

export default class MyCalendars extends React.Component<IMyCalendarsProps, IMyCalendarsState> {
  private debounceTimer: number | null = null;

  constructor(props: IMyCalendarsProps) {
    super(props);
    
    this.state = {
      appointments: [],
      currentDate: new Date(),
      currentView: props.settings.defaultView,
      previousView: props.settings.defaultView,
      isLoading: false,
      isSettingsPanelOpen: false,
      searchQuery: ''
    };
  }

  public componentDidMount(): void {
    this.loadAppointments().catch(err => console.error('Failed to load appointments:', err));
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
    this.setState({ isLoading: true });

    const allAppointments: IAppointment[] = [];
    const httpClient = this.props.context.httpClient;
    const graphClientPromise = this.props.context.msGraphClientFactory.getClient('3');
    const graphClient = await graphClientPromise;
    const exchangeService = new ExchangeCalendarService(httpClient, graphClient);
    exchangeService.setGraphClient(graphClient);
    const sharePointService = new SharePointCalendarService(httpClient, graphClient);
    sharePointService.setGraphClient(graphClient);
    
    // Calculate date range for filtering (current month ± 3 months)
    const today = new Date();
    const startDate = new Date(today.getFullYear(), today.getMonth() - 3, 1);
    const endDate = new Date(today.getFullYear(), today.getMonth() + 4, 0);
    
    // Load Exchange calendars (user's own calendars)
    try {
      const userCalendars = await exchangeService.getCalendars();
      const exchangeCalendarStates = this.props.settings.exchangeCalendarStates || {};
      
      for (const calendar of userCalendars) {
        // Check if calendar is enabled (default to true if not set)
        const isEnabled = exchangeCalendarStates[calendar.id] !== false;
        
        if (isEnabled) {
          try {
            const events = await exchangeService.getCalendarEvents(
              calendar.id,
              startDate,
              endDate,
              undefined // current user
            );
            const appointments = events.map(event => ({
              ...event,
              sourceId: `exchange_${calendar.id}`,
              color: calendar.hexColor
            }));
            allAppointments.push(...appointments);
          } catch (error) {
            console.error(`Failed to load Exchange calendar ${calendar.name}:`, error);
          }
        }
      }
    } catch (error) {
      console.error('Failed to load user Exchange calendars:', error);
    }
    
    // Load from manually added sources (SharePoint and ICS)
    for (const source of this.props.settings.sources) {
      if (source.isEnabled) {
        try {
          let appointments: IAppointment[] = [];
          
          if (source.sourceType === 'ics') {
            // Load from raw content if available, otherwise fetch from URL
            if (source.rawContent) {
              appointments = IcsParser.parseRawContent(source.rawContent, source.id, source.color);
            } else if (source.url) {
              appointments = await IcsParser.fetchAndParse(source.url, source.id, source.color, httpClient);
            }
          } else if (source.sourceType === 'sharepoint') {
            // Load from SharePoint list
            try {
              if (source.sharePointSiteId && source.sharePointListId) {
                const items = await sharePointService.getListEvents(
                  source.sharePointSiteId,
                  source.sharePointListId,
                  startDate,
                  endDate,
                  source.sharePointFieldMapping
                );
                appointments = items.map(item => ({
                  ...item,
                  sourceId: source.id,
                  color: source.color
                }));
              }
            } catch (error) {
              console.error(`Failed to load SharePoint calendar ${source.name}:`, error);
            }
          }
          
          allAppointments.push(...appointments);
        } catch (error) {
          console.error(`Failed to load calendar: ${source.name}`, error);
        }
      }
    }

    this.setState({
      appointments: allAppointments,
      isLoading: false
    });
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
    const { currentView, previousView } = this.state;
    const displayView = currentView === 'search' ? previousView : currentView;

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
      {
        key: 'refresh',
        iconProps: { iconName: 'Refresh' },
        onClick: () => {
          this.loadAppointments().catch(err => console.error('Failed to load appointments:', err));
        }
      }
    ];
  };

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
          graphClient={this.props.context.msGraphClientFactory.getClient('3')}
        />
      </div>
    );
  }
}