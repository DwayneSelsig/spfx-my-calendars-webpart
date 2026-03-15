import { HttpClient, type MSGraphClientV3 } from '@microsoft/sp-http';
import { IEvent } from '@pnp/spfx-controls-react/lib/controls/calendar/models/IEvents';
import { ISharePointFieldMapping } from '../models/ICalendarSettings';

export interface ISharePointSite {
  id: string;
  name: string;
  url: string;
}

export interface ISharePointList {
  id: string;
  name: string;
  webUrl: string;
}

interface IGraphSite {
  id: string;
  displayName?: string;
  name?: string;
  webUrl: string;
}

interface IGraphList {
  id: string;
  displayName?: string;
  name?: string;
  webUrl: string;
  list?: {
    template?: string;
    hidden?: boolean;
  };
}

interface IGraphListItemWithFields {
  id: string;
  fields?: Record<string, unknown>;
}

// Alias for backwards compatibility
type IGraphListItem = IGraphListItemWithFields;

/**
 * Service to interact with SharePoint calendars via Microsoft Graph API
 * Requires Sites.Read.All permission
 */
export class SharePointCalendarService {
  private readonly GRAPH_API_URL = 'https://graph.microsoft.com/v1.0';
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
   * Get accessible SharePoint sites (joined sites + followed sites)
   */
  public async getAccessibleSites(): Promise<ISharePointSite[]> {
    try {
      if (!this.graphClient) {
        console.error('GraphClient not initialized');
        return [];
      }

      console.log('Fetching accessible SharePoint sites...');
      
      // Use wildcard search to get all accessible sites
      const data = await this.graphClient
        .api('/sites')
        .query({
          search: '*',
          $select: 'id,displayName,webUrl',
          $top: 999
        })
        .get();

      console.log('Response data:', data);
      
      const sites = (data.value || []).map((site: IGraphSite) => ({
        id: site.id,
        name: site.displayName || site.name || 'Unnamed Site',
        url: site.webUrl
      }));

      console.log('Found accessible sites:', sites);
      return sites;
    } catch (error) {
      console.error('Error fetching accessible sites:', error);
      return [];
    }
  }

  /**
   * Search sites by name filter
   */
  public async searchSites(filter: string): Promise<ISharePointSite[]> {
    if (!filter || filter.trim().length === 0) {
      return this.getAccessibleSites();
    }

    if (!this.graphClient) {
      console.error('GraphClient not initialized');
      return [];
    }

    try {
      console.log('Searching sites for:', filter);
      
      const data = await this.graphClient
        .api('/sites')
        .query({
          search: filter,
          $select: 'id,displayName,webUrl',
          $top: 999
        })
        .get();

      const searchResults = (data.value || []).map((site: IGraphSite) => ({
        id: site.id,
        name: site.displayName || site.name || 'Unnamed Site',
        url: site.webUrl
      }));

      console.log('Search results for "' + filter + '":', searchResults);
      return searchResults;
    } catch (error) {
      console.error('Error searching sites:', error);
      // Fallback to accessible sites on error
      return this.getAccessibleSites();
    }
  }

  /**
   * Get calendar-capable lists from a site
   * Includes classic "Events" lists and modern list templates
   */
  public async getCalendarLists(siteId: string): Promise<ISharePointList[]> {
    if (!this.graphClient) {
      console.error('GraphClient not initialized');
      return [];
    }

    try {
      // Query lists and filter client-side for calendar-capable templates
      const data = await this.graphClient
        .api(`/sites/${siteId}/lists`)
        .query({
          $select: 'id,displayName,name,webUrl,list',
          $top: 200
        })
        .get();

      return (data.value || [])
        .filter((list: IGraphList) => {
          const template = list.list?.template;
          const isHidden = list.list?.hidden === true;
          if (isHidden) {
            return false;
          }
          // classic calendar = events, modern calendar often still genericList but named Calendar
          const isClassic = template === 'events';
          const isModern = template === 'genericList' 
          return isClassic || isModern;
        })
        .map((list: IGraphList) => ({
          id: list.id,
          name: list.displayName || list.name || 'Unnamed List',
          webUrl: list.webUrl
        }));
    } catch (error) {
      console.error('Error fetching calendar lists:', error);
      return [];
    }
  }

  /**
   * Get events from a SharePoint list (classic Events list)
   * Maps standard SharePoint columns: Title, EventDate, EndDate, Location, Description
   */
  public async getListEvents(
    siteId: string,
    listId: string,
    startDate: Date,
    endDate: Date,
    fieldMapping?: ISharePointFieldMapping
  ): Promise<IEvent[]> {
    if (!this.graphClient) {
      console.error('GraphClient not initialized');
      return [];
    }

    try {
      // Fetch all items - need to expand fields to get the actual field values
      const data = await this.graphClient
        .api(`/sites/${siteId}/lists/${listId}/items`)
        .expand('fields')
        .get();

      // Auto-detect field mapping from first item if not provided
      const effectiveMapping = fieldMapping || (data.value?.[0] ? this.detectFieldMapping(data.value[0]) : {});

      return (data.value || [])
        .map((item: IGraphListItem) => this.mapListItemToAppointment(item, startDate, endDate, effectiveMapping))
        .filter((apt: IEvent | null): apt is IEvent => apt !== null);
    } catch (error) {
      console.error('Error fetching list events:', error);
      return [];
    }
  }

  /**
   * Detect field mapping from first item by checking which fields exist
   */
  private detectFieldMapping(firstItem: IGraphListItemWithFields): ISharePointFieldMapping {
    if (!firstItem?.fields) {
      return {};
    }

    const fields = firstItem.fields as Record<string, unknown>;
    const fieldNames = Object.keys(fields);

    const mapping: ISharePointFieldMapping = {
      titleField: this.findField(fieldNames, ['Title', 'EventTitle', 'Subject']),
      startDateField: this.findField(fieldNames, ['EventDate', 'StartDate', 'StartDateTime', 'Start']),
      endDateField: this.findField(fieldNames, ['EndDate', 'DueDate', 'EndDateTime', 'End']),
      descriptionField: this.findField(fieldNames, ['Description', 'Body', 'Notes']),
      locationField: this.findField(fieldNames, ['Location', 'Place', 'Room']),
      allDayField: this.findField(fieldNames, ['fAllDayEvent', 'AllDayEvent', 'IsAllDay'])
    };

    console.log('[SharePointCalendarService] Auto-detected field mapping:', mapping);
    return mapping;
  }

  /**
   * Find first matching field name (case-insensitive)
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private findField(availableFields: string[], candidates: string[]): string | undefined {
    for (const candidate of candidates) {
      const match = availableFields.find(f => f.toLowerCase() === candidate.toLowerCase());
      if (match) {
        return match;
      }
    }
    return undefined;
  }

  /**
   * Map SharePoint list item to IEvent
   * Supports custom field mappings for different SharePoint calendar list schemas
   */
  private mapListItemToAppointment(item: IGraphListItem, startDate?: Date, endDate?: Date, fieldMapping?: ISharePointFieldMapping): IEvent | null {
    const fields = (item.fields || {}) as Record<string, unknown>;
    
    // Apply field mapping
    const titleFieldName = fieldMapping?.titleField || 'Title';
    const startDateFieldName = fieldMapping?.startDateField || 'EventDate';
    const endDateFieldName = fieldMapping?.endDateField || 'EndDate';
    const descriptionFieldName = fieldMapping?.descriptionField || 'Description';
    const locationFieldName = fieldMapping?.locationField || 'Location';
    const allDayFieldName = fieldMapping?.allDayField || 'fAllDayEvent';
    
    const title = fields[titleFieldName] as string | undefined;
    if (!title) {
      // Log available fields for debugging first time only
      if (!fieldMapping) {
        console.log(`[SharePoint] Item ${item.id}: Available fields:`, Object.keys(fields));
      }
      return null; // Skip items without title
    }

    const startDateValue = fields[startDateFieldName] as string | undefined;
    const endDateValue = fields[endDateFieldName] as string | undefined;
    const eventStart = startDateValue ? new Date(startDateValue) : new Date();
    const eventEnd = endDateValue ? new Date(endDateValue) : (startDateValue ? new Date(startDateValue) : new Date());
    const isAllDay = (fields[allDayFieldName] as boolean | undefined) === true;

    // Client-side date range filtering
    if (startDate && endDate) {
      if (eventEnd < startDate || eventStart > endDate) {
        return null; // Event is outside the requested date range
      }
    }

    return {
      id: item.id,
      title,
      description: (fields[descriptionFieldName] as string | undefined) || '',
      location: (fields[locationFieldName] as string | undefined) || undefined,
      start: eventStart.toISOString(),
      end: eventEnd.toISOString(),
      isFullDay: isAllDay,
      sourceId: '', // Will be set by caller
      color: undefined, // Will be set by caller
      attendees: []
    };
  }
}

