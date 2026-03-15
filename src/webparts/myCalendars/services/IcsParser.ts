import { IEvent } from '@pnp/spfx-controls-react/lib/controls/calendar/models/IEvents';
import { ICalendarSettings } from '../models/ICalendarSettings';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';

export class IcsParser {
  public static async fetchAndParse(
    url: string | undefined,
    sourceId: string,
    color: string,
    httpClient: HttpClient,
    settings?: ICalendarSettings
  ): Promise<IEvent[]> {
    if (!url) {
      return [];
    }
    
    try {
      // First attempt: direct fetch
      const directResponse: HttpClientResponse = await httpClient.get(url, HttpClient.configurations.v1);
      if (directResponse.ok) {
        const icsText = await directResponse.text();
        return this.parse(icsText, sourceId, color);
      }

      // If not OK, fall back using proxies in configured order
      const icsText = await this.fetchWithProxies(url, httpClient, settings);
      if (icsText) {
        return this.parse(icsText, sourceId, color);
      }
      throw new Error(`Failed to fetch ICS directly and via proxies: ${directResponse.statusText}`);
    } catch (error) {
      // Likely CORS/network error -> try proxies
      console.warn('Direct fetch failed, attempting proxies...', error);
      try {
        const icsText = await this.fetchWithProxies(url, httpClient, settings);
        if (icsText) {
          console.log('Successfully fetched via proxy');
          return this.parse(icsText, sourceId, color);
        }
      } catch (proxyError) {
        console.error('Proxy fetch also failed:', proxyError);
      }
      console.error('Error fetching ICS file:', error);
      return [];
    }
  }

  public static parseRawContent(icsText: string, sourceId: string, color: string): IEvent[] {
    return this.parse(icsText, sourceId, color);
  }

  public static parse(icsText: string, sourceId: string, color: string): IEvent[] {
    const appointments: IEvent[] = [];
    const lines = icsText.split(/\r\n|\n|\r/);
    
    let inEvent = false;
    let currentEvent: Partial<IEvent> = {};

    for (let i = 0; i < lines.length; i++) {
      let line = lines[i].trim();
      
      // Handle line continuation
      while (i + 1 < lines.length && (lines[i + 1].startsWith(' ') || lines[i + 1].startsWith('\t'))) {
        i++;
        line += lines[i].trim();
      }

      if (line === 'BEGIN:VEVENT') {
        inEvent = true;
        currentEvent = {
          sourceId,
          color: undefined,
          isFullDay: false
        };
      } else if (line === 'END:VEVENT' && inEvent) {
        if (currentEvent.id && currentEvent.title && currentEvent.start && currentEvent.end) {
          appointments.push(currentEvent as IEvent);
        }
        inEvent = false;
        currentEvent = {};
      } else if (inEvent) {
        const colonIndex = line.indexOf(':');
        if (colonIndex === -1) continue;

        const fullKey = line.substring(0, colonIndex);
        const value = line.substring(colonIndex + 1);

        const [key, ...params] = fullKey.split(';');

        switch (key) {
          case 'UID':
            currentEvent.id = value;
            break;
          case 'SUMMARY':
            currentEvent.title = this.unescapeValue(value);
            break;
          case 'DESCRIPTION':
            currentEvent.description = this.unescapeValue(value);
            break;
          case 'LOCATION':
            currentEvent.location = this.unescapeValue(value);
            break;
          case 'DTSTART':
            currentEvent.start = this.parseDate(value, params).toISOString();
            if (params.some(p => p.includes('VALUE=DATE'))) {
              currentEvent.isFullDay = true;
            }
            break;
          case 'DTEND':
            currentEvent.end = this.parseDate(value, params).toISOString();
            break;
          case 'ORGANIZER': {
            // IEvent doesn't have organizer field, skip for now
            break;
          }
        }
      }
    }

    return appointments;
  }

  private static parseDate(value: string, params: string[]): Date {
    // Check if it's a date-only value (no time)
    if (value.length === 8) {
      // YYYYMMDD format
      const year = parseInt(value.substring(0, 4), 10);
      const month = parseInt(value.substring(4, 6), 10) - 1;
      const day = parseInt(value.substring(6, 8), 10);
      return new Date(year, month, day);
    }

    // DateTime format: YYYYMMDDTHHmmss or YYYYMMDDTHHmmssZ
    const year = parseInt(value.substring(0, 4), 10);
    const month = parseInt(value.substring(4, 6), 10) - 1;
    const day = parseInt(value.substring(6, 8), 10);
    const hour = parseInt(value.substring(9, 11), 10);
    const minute = parseInt(value.substring(11, 13), 10);
    const second = parseInt(value.substring(13, 15), 10);

    // Check if UTC (ends with Z) or has TZID parameter
    if (value.endsWith('Z')) {
      return new Date(Date.UTC(year, month, day, hour, minute, second));
    }

    // For now, treat as local time
    // In production, you'd want to handle TZID properly
    return new Date(year, month, day, hour, minute, second);
  }

  private static unescapeValue(value: string): string {
    return value
      .replace(/\\n/g, '\n')
      .replace(/\\,/g, ',')
      .replace(/\\;/g, ';')
      .replace(/\\\\/g, '\\');
  }

  private static async fetchWithProxies(
    url: string,
    httpClient: HttpClient,
    settings?: ICalendarSettings
  ): Promise<string | null> {
    const enabledProxies: Array<'custom' | 'whateverorigin'> = [];
    if (settings?.useCustomProxy) enabledProxies.push('custom');
    if (settings?.useWhateverOrigin) enabledProxies.push('whateverorigin');

    console.debug('Enabled proxies:', enabledProxies);

    // Order by admin preference
    const order = [settings?.proxyPriority1, settings?.proxyPriority2]
      .filter(Boolean) as Array<'custom' | 'whateverorigin'>;
    const ordered = order.filter(p => enabledProxies.indexOf(p) !== -1);
    let proxiesToTry = ordered.length ? ordered : enabledProxies;
    
    // Defensive default: if none enabled, try custom first, then whateverorigin
    if (proxiesToTry.length === 0) {
      console.warn('No proxies enabled in settings; trying default order custom -> whateverorigin');
      proxiesToTry = ['custom', 'whateverorigin'];
    }

    console.info('Proxy order to try:', proxiesToTry);

    for (const p of proxiesToTry) {
      try {
        console.info(`Attempting proxy: ${p}`);
        if (p === 'custom' && settings?.customProxyUrl) {
          const proxyUrl = this.buildCustomProxyUrl(settings.customProxyUrl, url);
          console.info(`Custom proxy URL: ${proxyUrl}`);
          const resp = await httpClient.get(proxyUrl, HttpClient.configurations.v1);
          if (resp.ok) {
            console.info("ICS fetched via proxy: custom");
            return await resp.text();
          }
        } else if (p === 'whateverorigin') {
          const proxyUrl = `https://whateverorigin.org/get?url=${encodeURIComponent(url)}`;
          console.info(`WhateverOrigin URL: ${proxyUrl}`);
          const resp = await fetch(proxyUrl, { method: 'GET' });
          if (resp.ok) {
            const data = await resp.json();
            if (data && data.contents) {
              console.info("ICS fetched via proxy: whateverorigin");
              return data.contents as string;
            }
          }
        }
      } catch (e) {
        // Continue to next proxy
        // eslint-disable-next-line no-console
        console.warn(`Proxy '${p}' failed`, e);
      }
    }
    console.warn('All proxies exhausted, none succeeded');
    return null;
  }

  private static buildCustomProxyUrl(base: string, target: string): string {
    if (base.includes('{url}')) {
      return base.replace('{url}', encodeURIComponent(target));
    }
    if (base.endsWith('/')) {
      return `${base}${target}`;
    }
    // Default: append target as-is
    return `${base}${target}`;
  }
}

