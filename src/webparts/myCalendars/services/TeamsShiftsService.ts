import { HttpClient, type MSGraphClientV3 } from '@microsoft/sp-http';
import { IEvent } from '@pnp/spfx-controls-react/lib/controls/calendar/models/IEvents';
import { ICalendarSource } from '../models/ICalendarSettings';

interface IGraphTeam {
  id: string;
  displayName: string;
}

interface IGraphShiftActivity {
  displayName?: string;
}

interface IGraphShiftItem {
  displayName?: string;
  notes?: string;
  startDateTime?: string;
  endDateTime?: string;
  theme?: string;
  activities?: IGraphShiftActivity[];
}

interface IGraphShift {
  id: string;
  sharedShift?: IGraphShiftItem;
  draftShift?: IGraphShiftItem;
  userId?: string;
}

/**
 * Service to retrieve Teams Shifts via Microsoft Graph API
 * Requires Team.ReadBasic.All and Schedule.Read.All permissions
 */
export class TeamsShiftsService {
  private httpClient: HttpClient;
  private graphClient: MSGraphClientV3 | undefined;

  constructor(httpClient: HttpClient, graphClient?: MSGraphClientV3) {
    this.httpClient = httpClient;
    this.graphClient = graphClient;
  }

  /**
   * Set the GraphClient (for authentication)
   */
  public setGraphClient(client: MSGraphClientV3): void {
    this.graphClient = client;
  }

  /**
   * Get shifts for all joined Teams, filtered to teams that actually have shifts
   */
  public async getShiftsForJoinedTeams(
    startDate: Date,
    endDate: Date,
    source: ICalendarSource,
    showSourceLogo: boolean = true
  ): Promise<IEvent[]> {
    if (!this.graphClient) {
      console.error('GraphClient not initialized');
      return [];
    }

    const teams = await this.getJoinedTeams();
    if (teams.length === 0) {
      return [];
    }

    const allAppointments: IEvent[] = [];

    for (const team of teams) {
      const teamShifts = await this.getTeamShifts(team.id, startDate, endDate);
      const mappedShifts = teamShifts.reduce((acc: IEvent[], shift: IGraphShift) => {
        acc.push(...this.mapShiftToAppointments(shift, team, source, showSourceLogo));
        return acc;
      }, []);

      if (mappedShifts.length > 0) {
        allAppointments.push(...mappedShifts);
      }
    }

    return allAppointments;
  }

  private async getJoinedTeams(): Promise<IGraphTeam[]> {
    try {
      if (!this.graphClient) {
        console.error('GraphClient not initialized');
        return [];
      }

      const teams: IGraphTeam[] = [];
      let nextLink: string | undefined = '/me/joinedTeams';

      while (nextLink) {
        const request = nextLink.startsWith('http')
          ? this.graphClient.api(nextLink)
          : this.graphClient.api(nextLink).select('id,displayName');

        const response = await request.get();
        teams.push(...((response.value || []) as IGraphTeam[]));
        nextLink = response['@odata.nextLink'] as string | undefined;
      }

      return teams;
    } catch (error) {
      console.error('Error fetching joined Teams:', error);
      return [];
    }
  }

  private async getTeamShifts(teamId: string, startDate: Date, endDate: Date): Promise<IGraphShift[]> {
    try {
      if (!this.graphClient) {
        console.error('GraphClient not initialized');
        return [];
      }

      const shifts: IGraphShift[] = [];
      let nextLink: string | undefined = `/teams/${teamId}/schedule/shifts`;

      while (nextLink) {
        const request = nextLink.startsWith('http')
          ? this.graphClient.api(nextLink)
          : this.graphClient.api(nextLink);

        const response = await request.get();
        shifts.push(...((response.value || []) as IGraphShift[]));
        nextLink = response['@odata.nextLink'] as string | undefined;
      }

      return this.filterShiftsByDate(shifts, startDate, endDate);
    } catch (error) {
      const statusCode = (error as { statusCode?: number }).statusCode;
      const message = (error as { message?: string }).message || '';
      const isNotFound = statusCode === 404 || message.includes('TeamNotFound');

      if (!isNotFound) {
        console.error(`Error fetching shifts for team ${teamId}:`, error);
      }
      return [];
    }
  }

  private filterShiftsByDate(shifts: IGraphShift[], startDate: Date, endDate: Date): IGraphShift[] {
    return shifts.filter(shift => {
      const sharedStart = shift.sharedShift?.startDateTime ? new Date(shift.sharedShift.startDateTime) : undefined;
      const sharedEnd = shift.sharedShift?.endDateTime ? new Date(shift.sharedShift.endDateTime) : undefined;
      const draftStart = shift.draftShift?.startDateTime ? new Date(shift.draftShift.startDateTime) : undefined;
      const draftEnd = shift.draftShift?.endDateTime ? new Date(shift.draftShift.endDateTime) : undefined;

      const sharedInRange = sharedStart && sharedEnd && sharedStart >= startDate && sharedEnd <= endDate;
      const draftInRange = draftStart && draftEnd && draftStart >= startDate && draftEnd <= endDate;

      return !!sharedInRange || !!draftInRange;
    });
  }

  private mapShiftToAppointments(
    shift: IGraphShift,
    team: IGraphTeam,
    source: ICalendarSource,
    showSourceLogo: boolean
  ): IEvent[] {
    const appointments: IEvent[] = [];

    if (shift.draftShift) {
      const draftAppointment = this.createAppointmentFromShiftItem(
        shift.draftShift,
        team,
        source,
        showSourceLogo,
        'draft',
        true
      );
      if (draftAppointment) {
        appointments.push(draftAppointment);
      }
      return appointments;
    }

    const sharedAppointment = this.createAppointmentFromShiftItem(
      shift.sharedShift,
      team,
      source,
      showSourceLogo,
      'shared',
      false
    );
    if (sharedAppointment) {
      appointments.push(sharedAppointment);
    }

    return appointments;
  }

  private createAppointmentFromShiftItem(
    shiftItem: IGraphShiftItem | undefined,
    team: IGraphTeam,
    source: ICalendarSource,
    showSourceLogo: boolean,
    suffix: 'shared' | 'draft',
    isDraft: boolean
  ): IEvent | null {
    if (!shiftItem?.startDateTime || !shiftItem?.endDateTime) {
      return null;
    }

    const startDate = new Date(shiftItem.startDateTime);
    const endDate = new Date(shiftItem.endDateTime);

    let normalizedStart = startDate;
    let normalizedEnd = endDate;
    if (normalizedStart > normalizedEnd) {
      [normalizedStart, normalizedEnd] = [normalizedEnd, normalizedStart];
    }

    const durationHours = (normalizedEnd.getTime() - normalizedStart.getTime()) / (1000 * 60 * 60);
    const isAllDay = durationHours >= 12;

    const description = this.buildDescription(shiftItem);

    return {
      id: `shift_${team.id}_${suffix}_${normalizedStart.getTime()}_${normalizedEnd.getTime()}`,
      title: team.displayName,
      description,
      location: team.displayName,
      start: normalizedStart.toISOString(),
      end: normalizedEnd.toISOString(),
      isFullDay: isAllDay,
      sourceId: source.id,
      colorHex: this.mapThemeToColor(shiftItem.theme) || source.color,
      sourceType: 'teamsShifts',
      showSourceLogo,
      isDraft
    };
  }

  private buildDescription(shared: IGraphShiftItem): string | undefined {
    const notes = shared.notes?.trim();
    const activities = (shared.activities || [])
      .map(activity => activity.displayName)
      .filter((name): name is string => !!name && name.trim().length > 0);

    const activityText = activities.length > 0 ? `Activities: ${activities.join(', ')}` : '';

    if (notes && activityText) {
      return `${notes}\n${activityText}`;
    }

    return notes || activityText || undefined;
  }

  private mapThemeToColor(theme?: string): string | undefined {
    if (!theme) {
      return undefined;
    }

    const normalized = theme.toLowerCase().replace(/\s+/g, '');
    const map: Record<string, string> = {
      white: '#ffffff',
      blue: '#0078d4',
      darkblue: '#004578',
      green: '#107c10',
      darkgreen: '#0b6a0b',
      orange: '#d83b01',
      pink: '#e3008c',
      darkpink: '#b4006f',
      purple: '#5c2d91',
      darkpurple: '#4b1d6e',
      teal: '#038387',
      yellow: '#ffb900',
      darkyellow: '#c19c00',
      gray: '#8a8886'
    };

    return map[normalized];
  }
}

