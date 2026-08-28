import { HttpClient } from '@microsoft/sp-http';
import type { ICalendarEvent as IEvent } from '../models/ICalendarEvent';
import { UserHelper } from '../utils/userHelper';

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
  isAllDay?: boolean;
  isOnlineMeeting?: boolean;
  onlineMeeting?: { joinUrl?: string };
  webLink?: string;
  location?: { displayName: string };
  organizer?: { emailAddress: { name?: string; address?: string } };
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
      throw new Error('GraphClient not initialized');
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
      throw error;
    }
  }

  /**
   * Get joined Teams IDs for the current user
   */
  public async getJoinedTeamIds(): Promise<Set<string>> {
    if (!this.graphClient) {
      throw new Error('GraphClient not initialized');
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
      throw error;
    }
  }

  /**
   * Get calendar events for a specific group
   */
  public async getGroupEvents(groupId: string, startDate: Date, endDate: Date): Promise<IEvent[]> {
    if (!this.graphClient) {
      console.error('GraphClient not initialized');
      return [];
    }

    try {
      const startISO = startDate.toISOString();
      const endISO = endDate.toISOString();
      const mailboxSettings = await UserHelper.getCurrentUserMailboxSettings(this.graphClient);
      console.log('[UnifiedGroupCalendarService] Timezone debug', {
        groupId,
        graphTimeZone: mailboxSettings?.timeZone || null,
        appliedTimeZone: 'UTC',
        startISO,
        endISO
      });

      const data = await this.graphClient
        .api(`/groups/${groupId}/calendarView`)
        .header('Prefer', 'outlook.timezone="UTC"')
        .query({
          startDateTime: startISO,
          endDateTime: endISO,
          $select: 'id,subject,bodyPreview,start,end,isReminderOn,isAllDay,isOnlineMeeting,onlineMeeting,webLink,location,organizer,attendees',
          $top: 500
        })
        .get();

      // Get current user email for isOrganizer determination
      const currentUserEmail = await UserHelper.getCurrentUserEmail(this.graphClient);

      return Promise.all((data.value || []).map((event: IGraphEvent) => this.mapGraphEventToAppointment(event, currentUserEmail)));
    } catch (error) {
      console.error('Error fetching group calendar events:', error);
      throw error;
    }
  }

  private async mapGraphEventToAppointment(graphEvent: IGraphEvent, currentUserEmail: string): Promise<IEvent> {
    const startISO = UnifiedGroupCalendarService.toSafeISOString(graphEvent.start.dateTime, graphEvent.start.date, graphEvent.isAllDay);
    const endISO = UnifiedGroupCalendarService.toSafeISOString(graphEvent.end.dateTime, graphEvent.end.date, graphEvent.isAllDay);
    const organizerEmail = graphEvent.organizer?.emailAddress?.address;

    // Map attendees excluding the organizer
    const mappedAttendees = (graphEvent.attendees || [])
      .filter((att: IGraphEventAttendee) => att.type !== 'organizer')
      .map((att: IGraphEventAttendee) => ({
        id: att.emailAddress?.address || '',
        name: att.emailAddress?.name || att.emailAddress?.address || '',
        email: att.emailAddress?.address || ''
      }))
      .filter(att => att.id);

    return {
      id: graphEvent.id,
      title: graphEvent.subject || 'Untitled',
      description: graphEvent.bodyPreview || '',
      location: graphEvent.location?.displayName || undefined,
      start: startISO,
      end: endISO,
      isFullDay: graphEvent.isAllDay || false,
      sourceId: '',
      color: undefined,
      isOrganizer: UserHelper.isEventOrganizer(organizerEmail, currentUserEmail),
      organizer: {
        name: graphEvent.organizer?.emailAddress?.name,
        email: organizerEmail
      },
      attendees: mappedAttendees,
      isOnlineMeeting: graphEvent.isOnlineMeeting || false,
      joinUrl: graphEvent.onlineMeeting?.joinUrl || undefined,
      webLink: graphEvent.webLink || undefined
    };
  }

  /**
   * Normalises a Graph date value to a valid ISO string.
   * All-day dates become local midnights so Graph's exclusive end date stays
   * on the intended local calendar boundary. Timed values are requested in UTC.
   */
  private static toSafeISOString(dateTime: string | undefined, dateOnly: string | undefined, isAllDay?: boolean): string {
    const raw = dateTime || dateOnly;
    if (!raw) {
      throw new Error('Group event has no usable date.');
    }
    if (isAllDay || !raw.includes('T')) {
      const parts = raw.substring(0, 10).split('-').map(Number);
      return new Date(parts[0], parts[1] - 1, parts[2]).toISOString();
    }
    const normalized = /Z$|[+-]\d{2}:\d{2}$/.test(raw) ? raw : `${raw}Z`;
    const d = new Date(normalized);
    if (isNaN(d.getTime())) {
      throw new Error(`Group event has an invalid date: ${raw}`);
    }
    return d.toISOString();
  }
}

