jest.mock('@microsoft/sp-http', () => ({
  HttpClient: { configurations: { v1: {} } }
}));

import type { HttpClient } from '@microsoft/sp-http';
import { UserHelper } from '../utils/userHelper';
import { ExchangeCalendarService } from './ExchangeCalendarService';

describe('ExchangeCalendarService event mapping', () => {
  afterEach(() => jest.restoreAllMocks());

  it('uses the configured mailbox and only exposes organizer metadata for meetings', async () => {
    jest.spyOn(UserHelper, 'getCurrentUserMailboxSettings').mockResolvedValue(undefined);
    jest.spyOn(UserHelper, 'getCurrentUserEmail').mockResolvedValue('viewer@example.com');

    const get = jest.fn().mockResolvedValue({
      value: [
        {
          id: 'appointment',
          subject: 'Focus time',
          bodyPreview: 'Work without attendees',
          start: { dateTime: '2026-09-10T08:00:00Z' },
          end: { dateTime: '2026-09-10T09:00:00Z' },
          isReminderOn: false,
          organizer: { emailAddress: { name: 'Viewer', address: 'viewer@example.com' } },
          attendees: []
          ,showAs: 'free'
          ,responseStatus: { response: 'declined' }
        },
        {
          id: 'meeting',
          subject: 'Project meeting',
          bodyPreview: 'Meeting notes',
          start: { dateTime: '2026-09-10T10:00:00Z' },
          end: { dateTime: '2026-09-10T11:00:00Z' },
          isReminderOn: false,
          organizer: { emailAddress: { name: 'Viewer', address: 'viewer@example.com' } },
          attendees: [{ type: 'required', emailAddress: { name: 'Guest', address: 'guest@example.com' } }]
        }
      ]
    });
    const request = {
      header: jest.fn().mockReturnThis(),
      query: jest.fn().mockReturnThis(),
      get
    };
    const graphClient = { api: jest.fn().mockReturnValue(request) };
    const service = new ExchangeCalendarService({} as HttpClient, graphClient);

    const events = await service.getCalendarEvents(
      'calendar-id',
      new Date('2026-09-01T00:00:00Z'),
      new Date('2026-10-01T00:00:00Z'),
      'owner@example.com'
    );

    expect(graphClient.api).toHaveBeenCalledWith('/users/owner%40example.com/calendars/calendar-id/calendarView');
    expect(events[0].organizer).toBeUndefined();
    expect(events[0].isOrganizer).toBeUndefined();
    expect(events[0].showAs).toBe('free');
    expect(events[0].responseStatus).toBeUndefined();
    expect(events[1].organizer).toEqual({ name: 'Viewer', email: 'viewer@example.com' });
    expect(events[1].isOrganizer).toBe(true);
  });

  it('maps every supported current-user response and availability value', async () => {
    jest.spyOn(UserHelper, 'getCurrentUserMailboxSettings').mockResolvedValue(undefined);
    jest.spyOn(UserHelper, 'getCurrentUserEmail').mockResolvedValue('viewer@example.com');
    const responses = ['none', 'organizer', 'tentativelyAccepted', 'accepted', 'declined', 'notResponded'];
    const availability = ['free', 'tentative', 'busy', 'oof', 'workingElsewhere', 'unknown'];
    const get = jest.fn().mockResolvedValue({
      value: responses.map((response, index) => ({
        id: String(index), subject: response, bodyPreview: '',
        start: { dateTime: '2026-09-10T08:00:00Z' }, end: { dateTime: '2026-09-10T09:00:00Z' },
        isReminderOn: false, attendees: [], showAs: availability[index], responseStatus: { response }
      }))
    });
    const request = { header: jest.fn().mockReturnThis(), query: jest.fn().mockReturnThis(), get };
    const graphClient = { api: jest.fn().mockReturnValue(request) };
    const service = new ExchangeCalendarService({} as HttpClient, graphClient);

    const events = await service.getCalendarEvents('calendar-id', new Date('2026-09-01'), new Date('2026-10-01'));

    expect(events.map(event => event.responseStatus)).toEqual(responses);
    expect(events.map(event => event.showAs)).toEqual(availability);
  });
});
