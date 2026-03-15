import { HttpClient } from '@microsoft/sp-http';
import { GraphFI } from "@pnp/graph";
import { IEvent } from '@pnp/spfx-controls-react/lib/controls/calendar/models/IEvents';
import { UserHelper } from '../utils/userHelper';

// MSGraphClientV3 type - using any since @microsoft/sp-client-preview is not available
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type MSGraphClientV3 = any;

export interface IExchangeCalendar {
  id: string;
  name: string;
  hexColor: string;
  isDefaultCalendar: boolean;
  color: string;
  canViewPrivateItems: boolean;
}

interface IGraphCalendar {
  id: string;
  name: string;
  hexColor?: string;
  isDefaultCalendar?: boolean;
  color?: string;
  canViewPrivateItems?: boolean;
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
 * Service to interact with Exchange calendars via Microsoft Graph API
 * Requires Calendars.Read and Calendars.Read.Shared permissions
 */
export class ExchangeCalendarService {
  private readonly GRAPH_API_URL = 'https://graph.microsoft.com/v1.0';
  private httpClient: HttpClient;
  private graphClient: MSGraphClientV3 | null = null;
  private graph: GraphFI | null = null;

  // Color mapping based on Outlook calendar colors
  private readonly COLOR_MAP: Record<string, string> = {
    'lightBlue': '#4ECFFF',
    'lightGreen': '#52D979',
    'lightOrange': '#FFC670',
    'lightGray': '#BEBEBE',
    'lightYellow': '#FFEB70',
    'lightTeal': '#4FBBBA',
    'lightPink': '#FF8FA6',
    'lightBrown': '#B4946E',
    'lightRed': '#FF7A6C',
    'maxColor': '#FF7C80',
    'auto': '#0078D4' // Outlook blue for default calendar
  };

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

  public setGraph(graph: GraphFI): void {
    this.graph = graph;
  }

  /**
   * Get the hex color for a calendar
   * If hexColor is empty and calendar is default: use Outlook blue
   * Otherwise map the color property to hex
   */
  private getCalendarColor(calendar: IGraphCalendar): string {
    // If hexColor is provided and not empty, use it
    if (calendar.hexColor && calendar.hexColor.trim() !== '') {
      return calendar.hexColor;
    }

    // If it's the default calendar and no hexColor, use Outlook blue
    if (calendar.isDefaultCalendar === true) {
      return '#0078D4'; // Outlook blue
    }

    // Map the color property to hex (but not 'auto' for non-default calendars)
    const colorKey = calendar.color || 'auto';
    if (colorKey === 'auto' && !calendar.isDefaultCalendar) {
      // For non-default calendars with 'auto', return a fallback color
      return '#8764B8'; // Purple
    }
    return this.COLOR_MAP[colorKey] || '#8764B8';
  }

  /**
   * Get all calendars for a specific mailbox
   * @param mailbox - User email or UPN; undefined = current user
   */
  public async getCalendars(mailbox?: string): Promise<IExchangeCalendar[]> {
    try {
      if (!this.graphClient) {
        console.error('GraphClient not initialized');
        return [];
      }

      console.log('Fetching Exchange calendars...');

      const endpoint = mailbox
        ? `/users/${encodeURIComponent(mailbox)}/calendars`
        : '/me/calendars';

      const data = await this.graphClient
        .api(endpoint)
        .query({
          $select: 'id,name,hexColor,isDefaultCalendar,color,canViewPrivateItems'
        })
        .get();

      console.log('Exchange calendars response:', data);

      return (data.value || []).map((calendar: IGraphCalendar) => ({
        id: calendar.id,
        name: calendar.name,
        hexColor: this.getCalendarColor(calendar),
        isDefaultCalendar: calendar.isDefaultCalendar || false,
        color: calendar.color || 'auto',
        canViewPrivateItems: calendar.canViewPrivateItems || false
      }));
    } catch (error) {
      console.error('Error fetching Exchange calendars:', error);
      return [];
    }
  }

  /**
   * Get events from a specific calendar
   * @param calendarId - Exchange calendar ID (e.g., 'calendar', or specific calendar id)
   * @param mailbox - User email or UPN; undefined = current user
   * @param startDate - Start of date range
   * @param endDate - End of date range
   */
  public async getCalendarEvents(
    calendarId: string,
    startDate: Date,
    endDate: Date,
    mailbox?: string
  ): Promise<IEvent[]> {
    try {
      if (!this.graphClient) {
        console.error('GraphClient not initialized');
        return [];
      }

      const endpoint = mailbox
        ? `/users/${encodeURIComponent(mailbox)}/calendars/${calendarId}/events`
        : `/me/calendars/${calendarId}/events`;

      // Filter by date range
      const startISO = startDate.toISOString();
      const endISO = endDate.toISOString();

      const mailboxSettings = await UserHelper.getCurrentUserMailboxSettings(this.graphClient);
      const userTimeZone = mailboxSettings?.timeZone || 'UTC';
      console.log('[ExchangeCalendarService] Timezone debug', {
        calendarId,
        mailbox: mailbox || 'me',
        graphTimeZone: mailboxSettings?.timeZone || null,
        appliedTimeZone: userTimeZone,
        startISO,
        endISO
      });

      const data = await this.graphClient
        .api(endpoint)
        .header('Prefer', `outlook.timezone="${userTimeZone}"`)
        .query({
          $filter: `start/dateTime ge '${startISO}' and end/dateTime le '${endISO}'`,
          $select: 'id,subject,bodyPreview,start,end,isReminderOn,isAllDay,isOnlineMeeting,webLink,location,organizer,attendees',
          $top: 500
        })
        .get();

      // Get current user email for isOrganizer determination
      const currentUserEmail = await UserHelper.getCurrentUserEmail(this.graphClient);

      return (data.value || []).map((event: IGraphEvent) => this.mapGraphEventToAppointment(event, currentUserEmail));
    } catch (error) {
      console.error('Error fetching Exchange calendar events:', error);
      return [];
    }
  }

  /**
   * Map Microsoft Graph event to IEvent
   */
  private mapGraphEventToAppointment(graphEvent: IGraphEvent, currentUserEmail: string): IEvent {
    const startISO = graphEvent.start.dateTime || graphEvent.start.date || new Date().toISOString();
    const endISO = graphEvent.end.dateTime || graphEvent.end.date || new Date().toISOString();
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
      start: startISO,
      end: endISO,
      isFullDay: graphEvent.isAllDay || false,
      description: graphEvent.bodyPreview || '',
      location: graphEvent.location?.displayName || undefined,
      isOrganizer: UserHelper.isEventOrganizer(organizerEmail, currentUserEmail),
      attendees: mappedAttendees,
      sourceId: '', // Will be set by caller
      color: undefined, // Will be set by caller
      // Extended properties
      isOnlineMeeting: graphEvent.isOnlineMeeting || false,
      webLink: graphEvent.webLink || undefined
    };
  }

  /**
   * Resolve a mailbox by email or UPN to verify it exists
   * @param mailbox - User email or UPN
   */
  public async resolveMailbox(mailbox: string): Promise<boolean> {
    try {
      // Try to fetch user profile to verify the mailbox exists
      const url = `${this.GRAPH_API_URL}/users/${encodeURIComponent(mailbox)}?$select=id,userPrincipalName`;
      const response = await this.httpClient.get(url, HttpClient.configurations.v1);
      return response.ok;
    } catch (error) {
      console.error('Error resolving mailbox:', error);
      return false;
    }
  }
}

