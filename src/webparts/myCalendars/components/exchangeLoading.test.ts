import { loadExchangeSources } from './exchangeLoading';
import type { ICalendarEvent } from '../models/ICalendarEvent';

const event: ICalendarEvent = { id: 'event', sourceId: 'own', title: 'Meeting', start: '2026-10-05T08:00:00Z', end: '2026-10-05T09:00:00Z' };

describe('independent Exchange loading', () => {
  it('loads saved sources even if own calendar discovery fails', async () => {
    const loaded = jest.fn();
    const error = jest.fn();
    const result = await loadExchangeSources(async () => { throw new Error('discovery'); },
      [{ sourceId: 'saved', load: async () => [event] }], loaded, error);
    expect(result).toEqual({ events: [event], hadError: true });
    expect(loaded).toHaveBeenCalledWith('saved', [event]);
    expect(error).toHaveBeenCalledWith(new Error('discovery'));
  });

  it('keeps own and other saved results when calendarView fails and marks only successful loads', async () => {
    const loaded = jest.fn();
    const error = jest.fn();
    const result = await loadExchangeSources(async () => [{ sourceId: 'own', load: async () => [event] }], [
      { sourceId: 'limited', load: async () => { throw Object.assign(new Error('Calendar unavailable'), { statusCode: 404, code: 'ErrorItemNotFound' }); } },
      { sourceId: 'empty', load: async () => [] }
    ], loaded, error);
    expect(result).toEqual({ events: [event], hadError: true });
    expect(loaded.mock.calls.map(call => call[0]).sort()).toEqual(['empty', 'own']);
    expect(error).toHaveBeenCalledWith(expect.objectContaining({ statusCode: 404, code: 'ErrorItemNotFound' }), 'limited');
  });
});
