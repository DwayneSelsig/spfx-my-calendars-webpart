import { HttpClient } from '@microsoft/sp-http';
import { IAppointment } from '../models/IAppointment';

// MSGraphClientV3 type - using any since @microsoft/sp-client-preview is not available
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type MSGraphClientV3 = any;

export interface IUnifiedGroupItem {
  id: string;
  displayName: string;
  isTeam: boolean;
}

interface IGraphGroup {
  id: string;
  displayName?: string;
  groupTypes?: string[];
}

interface IGraphEvent {
  id: string;
  subject: string;
  bodyPreview: string;
  start: { dateTime?: string; date?: string };
  end: { dateTime?: string; date?: string };
  isReminderOn: boolean;
  location?: { displayName: string };
  organizer?: { emailAddress: { name?: string } };
  attendees: IGraphEventAttendee[];
}

interface IGraphEventAttendee {
  type: string;
  emailAddress: { name?: string; address?: string };
}

/**
 * Service to interact with M365 Group/Team calendars via Microsoft Graph API
 * Requires Group.Read.All and Calendars.Read permissions
 */
export class UnifiedGroupCalendarService {
  private httpClient: HttpClient;
  private graphClient: MSGraphClientV3 | null = null;

  constructor(httpClient: HttpClient, graphClient?: MSGraphClientV3) {
    this.httpClient = httpClient;
    this.graphClient = graphClient || null;
  }

  /**
   * Set the GraphClient (for authentication)
   */
  public setGraphClient(client: MSGraphClientV3): void {
    this.graphClient = client;
  }

  /**
   * Get Unified groups the current user is a member of
   */
  public async getUnifiedGroups(): Promise<IUnifiedGroupItem[]> {
    if (!this.graphClient) {
      console.error('GraphClient not initialized');
      return [];
    }

    try {
      const groups: IGraphGroup[] = [];
      let nextLink: string | undefined = '/me/memberOf/microsoft.graph.group';

      while (nextLink) {
        const request = nextLink.startsWith('http')
          ? this.graphClient.api(nextLink)
          : this.graphClient.api(nextLink).select('id,displayName,groupTypes');

        const response = await request.get();
        groups.push(...((response.value || []) as IGraphGroup[]));
        nextLink = response['@odata.nextLink'] as string | undefined;
      }

      return groups
        .filter(group => (group.groupTypes || []).indexOf('Unified') >= 0)
        .map(group => ({
          id: group.id,
          displayName: group.displayName || 'Unnamed Group',
          isTeam: false
        }));
    } catch (error) {
      console.error('Error fetching Unified groups:', error);
      return [];
    }
  }

  /**
   * Get joined Teams IDs for the current user
   */
  public async getJoinedTeamIds(): Promise<Set<string>> {
    if (!this.graphClient) {
      console.error('GraphClient not initialized');
      return new Set();
    }

    try {
      const teamIds: string[] = [];
      let nextLink: string | undefined = '/me/joinedTeams?$select=id';

      while (nextLink) {
        const request = nextLink.startsWith('http')
          ? this.graphClient.api(nextLink)
          : this.graphClient.api(nextLink);

        const response = await request.get();
        const ids = (response.value || []).map((team: { id?: string }) => team.id).filter(Boolean) as string[];
        teamIds.push(...ids);
        nextLink = response['@odata.nextLink'] as string | undefined;
      }

      return new Set(teamIds);
    } catch (error) {
      console.error('Error fetching joined Teams:', error);
      return new Set();
    }
  }

  /**
   * Get calendar events for a specific group
   */
  public async getGroupEvents(groupId: string, startDate: Date, endDate: Date): Promise<IAppointment[]> {
    if (!this.graphClient) {
      console.error('GraphClient not initialized');
      return [];
    }

    try {
      const startISO = startDate.toISOString();
      const endISO = endDate.toISOString();

      const data = await this.graphClient
        .api(`/groups/${groupId}/events`)
        .query({
          $filter: `start/dateTime ge '${startISO}' and end/dateTime le '${endISO}'`,
          $top: 500
        })
        .get();

      return (data.value || []).map((event: IGraphEvent) => this.mapGraphEventToAppointment(event));
    } catch (error) {
      console.error('Error fetching group calendar events:', error);
      return [];
    }
  }

  private mapGraphEventToAppointment(graphEvent: IGraphEvent): IAppointment {
    const start = new Date(graphEvent.start.dateTime || graphEvent.start.date || new Date());
    const end = new Date(graphEvent.end.dateTime || graphEvent.end.date || new Date());

    return {
      id: graphEvent.id,
      title: graphEvent.subject || 'Untitled',
      description: graphEvent.bodyPreview || '',
      location: graphEvent.location?.displayName || undefined,
      startDate: start,
      endDate: end,
      isAllDay: graphEvent.isReminderOn === false && !graphEvent.start.dateTime,
      sourceId: '',
      color: '',
      organizer: graphEvent.organizer?.emailAddress?.name || undefined,
      attendees: (graphEvent.attendees || [])
        .filter((att: IGraphEventAttendee) => att.type !== 'organizer')
        .map((att: IGraphEventAttendee) => att.emailAddress?.name || att.emailAddress?.address || '')
        .filter(Boolean)
    };
  }
}
