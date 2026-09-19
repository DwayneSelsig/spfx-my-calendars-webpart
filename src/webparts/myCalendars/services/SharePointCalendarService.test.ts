import type { HttpClient } from '@microsoft/sp-http';
import { eventOccursOnDay, localDateKey } from '../components/views/calendarUtils';
import { normalizeSharePointEventDates, SharePointCalendarService } from './SharePointCalendarService';

describe('SharePointCalendarService event mapping', () => {
  it.each(['2026-03-29', '2026-10-25'])('keeps the all-day boundary at local midnight for %s', date => {
    const dates = normalizeSharePointEventDates(`${date}T00:00:00Z`, `${date}T23:59:00Z`, true);

    expect(dates).toBeDefined();
    expect(localDateKey(dates?.start as Date)).toBe(date);
    expect((dates?.start as Date).getHours()).toBe(0);
    expect((dates?.end as Date).getHours()).toBe(0);
  });

  it('converts an inclusive multi-day end into an exclusive local boundary', () => {
    const dates = normalizeSharePointEventDates('2026-03-28T00:00:00Z', '2026-03-29T23:59:00Z', true);
    expect(dates).toBeDefined();
    const event = {
      id: 'all-day',
      sourceId: 'sharepoint',
      title: 'Weekend',
      start: (dates?.start as Date).toISOString(),
      end: (dates?.end as Date).toISOString(),
      isFullDay: true
    };

    expect(eventOccursOnDay(event, new Date(2026, 2, 28))).toBe(true);
    expect(eventOccursOnDay(event, new Date(2026, 2, 29))).toBe(true);
    expect(eventOccursOnDay(event, new Date(2026, 2, 30))).toBe(false);
  });

  it('preserves timed values and marks SharePoint descriptions as HTML', async () => {
    const get = jest.fn().mockResolvedValue({
      value: [{
        id: 'timed',
        fields: {
          Title: 'Timed event',
          EventDate: '2026-09-10T08:00:00Z',
          EndDate: '2026-09-10T09:00:00Z',
          Description: '<p><strong>Details</strong></p>',
          fAllDayEvent: false
        }
      }]
    });
    const request = { expand: jest.fn().mockReturnThis(), get };
    const service = new SharePointCalendarService({} as HttpClient, { api: jest.fn().mockReturnValue(request) } as never);

    const events = await service.getListEvents(
      'site-id',
      'list-id',
      new Date('2026-09-01T00:00:00Z'),
      new Date('2026-10-01T00:00:00Z'),
      undefined,
      'Contoso Site'
    );

    expect(events[0]).toMatchObject({
      start: '2026-09-10T08:00:00.000Z',
      end: '2026-09-10T09:00:00.000Z',
      description: '<p><strong>Details</strong></p>',
      descriptionFormat: 'html',
      sharePointSiteName: 'Contoso Site'
    });
  });

  it('resolves a missing site name once and does not block events when resolution fails', async () => {
    const itemData = { value: [{ id: '1', fields: { Title: 'Event', EventDate: '2026-09-10T08:00:00Z', EndDate: '2026-09-10T09:00:00Z' } }] };
    const siteGet = jest.fn().mockResolvedValue({ id: 'site-id', displayName: 'Resolved Site', webUrl: 'https://contoso.sharepoint.com/sites/resolved' });
    const itemGet = jest.fn().mockResolvedValue(itemData);
    const graphClient = {
      api: jest.fn((endpoint: string) => endpoint === '/sites/site-id'
        ? { query: jest.fn().mockReturnThis(), get: siteGet }
        : { expand: jest.fn().mockReturnThis(), get: itemGet })
    };
    const service = new SharePointCalendarService({} as HttpClient, graphClient as never);

    const first = await service.getListEvents('site-id', 'list-id', new Date('2026-09-01'), new Date('2026-10-01'));
    const second = await service.getListEvents('site-id', 'list-id', new Date('2026-09-01'), new Date('2026-10-01'));
    expect(first[0].sharePointSiteName).toBe('Resolved Site');
    expect(second[0].sharePointSiteName).toBe('Resolved Site');
    expect(siteGet).toHaveBeenCalledTimes(1);

    const failingClient = {
      api: jest.fn((endpoint: string) => endpoint === '/sites/site-id'
        ? { query: jest.fn().mockReturnThis(), get: jest.fn().mockRejectedValue(new Error('denied')) }
        : { expand: jest.fn().mockReturnThis(), get: jest.fn().mockResolvedValue(itemData) })
    };
    const fallbackService = new SharePointCalendarService({} as HttpClient, failingClient as never);
    await expect(fallbackService.getListEvents('site-id', 'list-id', new Date('2026-09-01'), new Date('2026-10-01'))).resolves.toHaveLength(1);
  });
});
