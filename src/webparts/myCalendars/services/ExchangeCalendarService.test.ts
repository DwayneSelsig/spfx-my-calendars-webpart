import { UserHelper } from '../utils/userHelper';
import { ExchangeCalendarService } from './ExchangeCalendarService';
import type { GraphClient } from './GraphRequestExecutor';

describe('ExchangeCalendarService event mapping', () => {
  afterEach(() => jest.restoreAllMocks());

  it('publishes Exchange pages incrementally and returns each appointment once', async () => {
    jest.spyOn(UserHelper, 'getCurrentUserEmail').mockResolvedValue('viewer@example.com');
    const event = (id: string): object => ({ id, subject: id, start: { dateTime: '2026-09-10T08:00:00Z' },
      end: { dateTime: '2026-09-10T09:00:00Z' } });
    const continuation = 'https://graph.microsoft.com/v1.0/me/calendars/calendar/calendarView?$skiptoken=OPAQUE';
    const get = jest.fn().mockResolvedValueOnce({ value: [event('one')], '@odata.nextLink': continuation });
    const onPage = jest.fn();
    get.mockImplementationOnce(async () => {
      expect(onPage).toHaveBeenCalledWith([expect.objectContaining({ id: 'one' })]);
      return { value: [event('one'), event('two')] };
    });
    const request = { header: jest.fn().mockReturnThis(), query: jest.fn().mockReturnThis(), get };
    const api = jest.fn().mockReturnValue(request);
    const service = new ExchangeCalendarService({ api });
    const events = await service.getCalendarEvents('calendar', new Date('2026-09-01'), new Date('2026-10-01'), undefined, onPage);
    expect(events.map(item => item.id)).toEqual(['one', 'two']);
    expect(api.mock.calls.map(call => call[0])).toEqual(['/me/calendars/calendar/calendarView', continuation]);
    expect(request.header.mock.calls).toEqual(Array(2).fill(['Prefer', 'outlook.timezone="UTC"']));
  });

  it('uses the configured mailbox and only exposes organizer metadata for meetings', async () => {
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
    const graphClient = { api: jest.fn((endpoint: string) => endpoint === '/users' ? { query: jest.fn().mockReturnThis(), get: jest.fn().mockResolvedValue({ value: [{ id: 'owner-id', userPrincipalName: 'owner@example.com' }] }) } : request) };
    const service = new ExchangeCalendarService(graphClient as unknown as GraphClient);

    const events = await service.getCalendarEvents(
      'calendar-id',
      new Date('2026-09-01T00:00:00Z'),
      new Date('2026-10-01T00:00:00Z'),
      'owner@example.com'
    );

    expect(graphClient.api).toHaveBeenCalledWith('/users/owner-id/calendars/calendar-id/calendarView');
    expect(events[0].organizer).toBeUndefined();
    expect(events[0].isOrganizer).toBeUndefined();
    expect(events[0].showAs).toBe('free');
    expect(events[0].responseStatus).toBeUndefined();
    expect(events[1].organizer).toEqual({ name: 'Viewer', email: 'viewer@example.com' });
    expect(events[1].isOrganizer).toBe(true);
  });

  it('maps every supported current-user response and availability value', async () => {
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
    const service = new ExchangeCalendarService(graphClient);

    const events = await service.getCalendarEvents('calendar-id', new Date('2026-09-01'), new Date('2026-10-01'));

    expect(events.map(event => event.responseStatus)).toEqual(responses);
    expect(events.map(event => event.showAs)).toEqual(availability);
  });

  it('uses the localized fallback for events without a subject', async () => {
    jest.spyOn(UserHelper, 'getCurrentUserEmail').mockResolvedValue('viewer@example.com');
    const request = {
      header: jest.fn().mockReturnThis(),
      query: jest.fn().mockReturnThis(),
      get: jest.fn().mockResolvedValue({ value: [{
        id: 'fallback', subject: '', bodyPreview: '',
        start: { dateTime: '2026-09-10T08:00:00Z' }, end: { dateTime: '2026-09-10T09:00:00Z' },
        isReminderOn: false, attendees: []
      }] })
    };
    const service = new ExchangeCalendarService({ api: jest.fn().mockReturnValue(request) });

    const events = await service.getCalendarEvents('calendar-id', new Date('2026-09-01'), new Date('2026-10-01'));

    expect(events[0].title).toBe('Untitled');
  });
});
