import { ExchangeCalendarService, ExchangeRequestError, getExchangeDiscoveryErrorMessage } from './ExchangeCalendarService';
import { UserHelper } from '../utils/userHelper';

function graph(pages: unknown[]): { api: jest.Mock; request: { select: jest.Mock; query: jest.Mock; header: jest.Mock; get: jest.Mock } } {
  const request = { select: jest.fn().mockReturnThis(), query: jest.fn().mockReturnThis(), header: jest.fn().mockReturnThis(), get: jest.fn() };
  pages.forEach(page => request.get.mockResolvedValueOnce(page));
  return { api: jest.fn().mockReturnValue(request), request };
}

describe('Exchange identity and calendar access', () => {
  beforeEach(() => jest.spyOn(console, 'error').mockImplementation(() => undefined));
  afterEach(() => jest.restoreAllMocks());

  it.each(['owner@contoso.com', 'primary@contoso.com', "o'brien@contoso.com"])('resolves UPN or primary SMTP %s and discovers calendars via object ID', async address => {
    const client = graph([{ value: [{ id: 'owner-id', userPrincipalName: 'owner@contoso.com' }] }, { value: [{ id: 'calendar', name: 'Shared' }] }]);
    const calendars = await new ExchangeCalendarService(client).getCalendars(`  ${address}  `);
    expect(client.request.query.mock.calls[0][0].$filter).toBe(`userPrincipalName eq '${address.replace(/'/g, "''")}' or mail eq '${address.replace(/'/g, "''")}'`);
    expect(client.api).toHaveBeenCalledWith('/users/owner-id/calendars');
    expect(calendars[0].name).toBe('Shared');
  });

  it('uses a direct object-ID lookup and follows address resolution pages', async () => {
    const id = '12345678-1234-1234-1234-123456789012';
    const direct = graph([{ id, userPrincipalName: 'upn@contoso.com' }]);
    expect(await new ExchangeCalendarService(direct).resolveMailbox(` ${id} `)).toEqual({ id, userPrincipalName: 'upn@contoso.com' });
    expect(direct.api).toHaveBeenCalledWith(`/users/${id}`);
    const paged = graph([{ value: [], '@odata.nextLink': 'next' }, { value: [{ id: 'one', userPrincipalName: 'upn' }] }]);
    expect((await new ExchangeCalendarService(paged).resolveMailbox('primary@contoso.com')).id).toBe('one');
    expect(paged.api).toHaveBeenCalledWith('next');
  });

  it('rejects empty, missing and ambiguous identities instead of returning false or choosing arbitrarily', async () => {
    const empty = graph([]);
    await expect(new ExchangeCalendarService(empty).getCalendars('  ')).rejects.toMatchObject({ stage: 'identity', reason: 'empty' });
    expect(empty.api).not.toHaveBeenCalled();
    await expect(new ExchangeCalendarService(graph([{ value: [] }])).resolveMailbox('missing@contoso.com')).rejects.toMatchObject({ reason: 'notFound' });
    await expect(new ExchangeCalendarService(graph([
      { value: [{ id: 'a' }], '@odata.nextLink': 'next' }, { value: [{ id: 'b' }] }
    ])).resolveMailbox('ambiguous@contoso.com')).rejects.toMatchObject({ reason: 'ambiguous' });
  });

  it.each([403, 404])('does not label discovery %s as a missing mailbox after identity succeeds', async statusCode => {
    const original = { statusCode, code: 'ErrorItemNotFound' };
    const client = graph([{ value: [{ id: 'owner-id', userPrincipalName: 'upn' }] }]);
    client.request.get.mockRejectedValueOnce(original);
    try {
      await new ExchangeCalendarService(client).getCalendars('primary@contoso.com');
      throw new Error('Expected failure');
    } catch (error) {
      expect(error).toMatchObject({ stage: 'discovery', reason: 'request', statusCode, originalError: original });
      expect(getExchangeDiscoveryErrorMessage(error)).not.toBe(getExchangeDiscoveryErrorMessage(new ExchangeRequestError('identity', 'notFound')));
    }
  });

  it('preserves identity-stage technical failures and calendarView-stage 404 context', async () => {
    const failure = { statusCode: 503, code: 'ServiceUnavailable' };
    const client = graph([]);
    client.request.get.mockRejectedValue(failure);
    await expect(new ExchangeCalendarService(client).getCalendars('upn@contoso.com')).rejects.toMatchObject({ stage: 'identity', originalError: failure });
    const events = graph([{ value: [{ id: 'owner-id' }] }]);
    events.request.get.mockRejectedValueOnce({ statusCode: 404, code: 'ErrorItemNotFound' });
    await expect(new ExchangeCalendarService(events).getCalendarEvents('calendar/id+', new Date('2026-10-01'), new Date('2026-11-01'), ' primary@contoso.com ')).rejects.toMatchObject({ stage: 'events', statusCode: 404 });
    expect(events.api).toHaveBeenCalledWith('/users/owner-id/calendars/calendar%2Fid%2B/calendarView');
  });

  it('keeps undefined mailbox on /me and encodes calendar IDs without directory lookup', async () => {
    jest.spyOn(UserHelper, 'getCurrentUserEmail').mockResolvedValue('me@contoso.com');
    const client = graph([{ value: [] }, { value: [] }]);
    const service = new ExchangeCalendarService(client);
    expect(await service.getCalendars()).toEqual([]);
    expect(await service.getCalendarEvents('calendar/id+', new Date('2026-10-01'), new Date('2026-11-01'))).toEqual([]);
    expect(client.api.mock.calls.map(call => call[0])).toEqual(['/me/calendars', '/me/calendars/calendar%2Fid%2B/calendarView']);
  });
});
