import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { ICalendarEvent as IEvent } from '../models/ICalendarEvent';
import { ISharePointFieldMapping } from '../models/ICalendarSettings';
import * as strings from 'MyCalendarsWebPartStrings';

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
  webUrl?: string;
  fields?: Record<string, unknown>;
}

// Alias for backwards compatibility
type IGraphListItem = IGraphListItemWithFields;

interface INormalizedSharePointDates {
  start: Date;
  end: Date;
}

/**
 * Graph bases a SharePoint list item's webUrl on its FileRef. For classic
 * calendar lists that can point to an internal file such as `2_.000`, which
 * the browser downloads instead of displaying the event. Use the containing
 * list path and the item's stable list ID to address SharePoint's display form.
 */
export function getSharePointItemWebLink(itemWebUrl: string | undefined, itemId: string): string | undefined {
  const normalizedWebUrl = itemWebUrl?.trim();
  const normalizedItemId = itemId?.trim();
  if (!normalizedWebUrl || !normalizedItemId || !/^https?:\/\//i.test(normalizedWebUrl)) return undefined;

  const urlWithoutQueryOrFragment = normalizedWebUrl.split(/[?#]/, 1)[0];
  const lastPathSeparator = urlWithoutQueryOrFragment.lastIndexOf('/');
  if (lastPathSeparator < 'https://a'.length) return undefined;

  return `${urlWithoutQueryOrFragment.substring(0, lastPathSeparator + 1)}DispForm.aspx?ID=${encodeURIComponent(normalizedItemId)}`;
}

function parseSharePointCalendarDate(value: string): Date | undefined {
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(value);
  if (!match) return undefined;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const parsed = new Date(year, month - 1, day);
  return parsed.getFullYear() === year && parsed.getMonth() === month - 1 && parsed.getDate() === day
    ? parsed
    : undefined;
}

/**
 * SharePoint stores all-day calendar boundaries as calendar dates. EndDate is
 * inclusive, while the local event contract uses an exclusive end boundary.
 */
export function normalizeSharePointEventDates(
  startValue: string,
  endValue: string | undefined,
  isAllDay: boolean
): INormalizedSharePointDates | undefined {
  if (!isAllDay) {
    const start = new Date(startValue);
    const end = endValue ? new Date(endValue) : new Date(start);
    return Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) ? undefined : { start, end };
  }

  const start = parseSharePointCalendarDate(startValue);
  const inclusiveEnd = parseSharePointCalendarDate(endValue || startValue);
  if (!start || !inclusiveEnd) return undefined;
  const end = new Date(inclusiveEnd);
  end.setDate(end.getDate() + 1);
  return end > start ? { start, end } : undefined;
}

/**
 * Service to interact with SharePoint calendars via Microsoft Graph API
 * Requires Sites.Read.All permission
 */
export class SharePointCalendarService {
  private readonly GRAPH_API_URL = 'https://graph.microsoft.com/v1.0';
  private graphClient: MSGraphClientV3 | null = null;
  private readonly sitePromises = new Map<string, Promise<ISharePointSite | undefined>>();

  constructor(graphClient?: MSGraphClientV3) {
    this.graphClient = graphClient || null;
  }

  /**
   * Set the GraphClient (for authentication)
   */
  public setGraphClient(client: MSGraphClientV3): void {
    this.graphClient = client;
    this.sitePromises.clear();
  }

  public getSite(siteId: string): Promise<ISharePointSite | undefined> {
    const existing = this.sitePromises.get(siteId);
    if (existing) return existing;

    const request = this.loadSite(siteId);
    this.sitePromises.set(siteId, request);
    request.then(site => {
      if (!site && this.sitePromises.get(siteId) === request) this.sitePromises.delete(siteId);
    }).catch(() => {
      if (this.sitePromises.get(siteId) === request) this.sitePromises.delete(siteId);
    });
    return request;
  }

  private async loadSite(siteId: string): Promise<ISharePointSite | undefined> {
    if (!this.graphClient) return undefined;
    try {
      const site = await this.graphClient
        .api(`/sites/${siteId}`)
        .query({ $select: 'id,displayName,name,webUrl' })
        .get() as IGraphSite;
      const resolvedName = (site.displayName || site.name || '').trim();
      if (!resolvedName) return undefined;
      return {
        id: site.id || siteId,
        name: resolvedName,
        url: site.webUrl || ''
      };
    } catch (error) {
      console.warn(`Could not resolve SharePoint site ${siteId}; continuing without its display name.`, error);
      return undefined;
    }
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
        name: site.displayName || site.name || strings.UnnamedSiteLabel,
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
      throw new Error('GraphClient not initialized');
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
        name: site.displayName || site.name || strings.UnnamedSiteLabel,
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
          name: list.displayName || list.name || strings.UnnamedListLabel,
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
    fieldMapping?: ISharePointFieldMapping,
    siteName?: string
  ): Promise<IEvent[]> {
    if (!this.graphClient) {
      throw new Error('GraphClient not initialized');
    }

    try {
      const resolvedSiteNamePromise = siteName?.trim()
        ? Promise.resolve(siteName.trim())
        : this.getSite(siteId).then(site => site?.name || undefined);
      // Fetch all items - need to expand fields to get the actual field values
      const data = await this.graphClient
        .api(`/sites/${siteId}/lists/${listId}/items`)
        .select('id,webUrl')
        .expand('fields')
        .get();

      // Auto-detect field mapping from first item if not provided
      const effectiveMapping = fieldMapping || (data.value?.[0] ? this.detectFieldMapping(data.value[0]) : {});

      const resolvedSiteName = await resolvedSiteNamePromise;
      return (data.value || [])
        .map((item: IGraphListItem) => this.mapListItemToAppointment(item, startDate, endDate, effectiveMapping))
        .filter((apt: IEvent | null): apt is IEvent => apt !== null)
        .map((event: IEvent) => ({ ...event, sharePointSiteName: resolvedSiteName }));
    } catch (error) {
      console.error('Error fetching list events:', error);
      throw error;
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
    if (!startDateValue) return null;
    const isAllDay = (fields[allDayFieldName] as boolean | undefined) === true;
    const eventDates = normalizeSharePointEventDates(startDateValue, endDateValue, isAllDay);
    if (!eventDates) return null;
    const { start: eventStart, end: eventEnd } = eventDates;

    // Client-side date range filtering
    if (startDate && endDate) {
      if (eventEnd <= startDate || eventStart >= endDate) {
        return null; // Event is outside the requested date range
      }
    }

    return {
      id: item.id,
      title,
      description: (fields[descriptionFieldName] as string | undefined) || '',
      descriptionFormat: 'html',
      location: (fields[locationFieldName] as string | undefined) || undefined,
      start: eventStart.toISOString(),
      end: eventEnd.toISOString(),
      isFullDay: isAllDay,
      sourceId: '', // Will be set by caller
      color: undefined, // Will be set by caller
      attendees: [],
      webLink: getSharePointItemWebLink(item.webUrl, item.id)
    };
  }
}

