import { GraphRequestExecutor, forEachGraphPage, type GraphRequest } from './GraphRequestExecutor';
import { GraphRequestScheduler, GraphRequestScope } from './GraphRequestScheduler';

describe('scheduled Graph pagination', () => {
  it('releases capacity between pages, follows the whole URL and keeps required headers after promotion', async () => {
    const scheduler = new GraphRequestScheduler(2);
    const scope = new GraphRequestScope();
    const next = 'https://graph.microsoft.com/v1.0/me/events?$skiptoken=ABC%2B12%3D';
    const requests = [
      { value: ['first'], '@odata.nextLink': next },
      { value: [], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/events?$skiptoken=next' },
      { value: ['last'] }
    ];
    const builder = {
      header: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnThis(),
      get: jest.fn(async () => {
        expect(scheduler.counts.activeHigh).toBe(requests.length === 3 ? 0 : 1);
        return requests.shift();
      })
    };
    const api = jest.fn((_path: string) => builder as unknown as GraphRequest);
    const graph = new GraphRequestExecutor(async () => ({ api }), scheduler);
    const pages: string[][] = [];
    await forEachGraphPage<string>(graph.scoped(scope), '/me/events', page => {
      expect(scheduler.counts.activeNormal + scheduler.counts.activeHigh).toBe(0);
      pages.push(page);
      scope.setHigh(true);
    }, request => request.top(1), { Prefer: 'outlook.timezone="UTC"' });
    expect(pages).toEqual([['first'], [], ['last']]);
    expect(api.mock.calls.map(call => call[0])).toEqual(['/me/events', next, 'https://graph.microsoft.com/v1.0/me/events?$skiptoken=next']);
    expect(builder.top).toHaveBeenCalledTimes(1);
    expect(builder.header.mock.calls).toEqual(Array(3).fill(['Prefer', 'outlook.timezone="UTC"']));
  });

  it('propagates a failed continuation and retains pages delivered before it', async () => {
    const scheduler = new GraphRequestScheduler(10);
    const post = jest.fn(async () => {
      expect(scheduler.counts.activeNormal).toBe(1);
      return { responses: Array(20).fill({ status: 200 }) };
    });
    const get = jest.fn().mockResolvedValueOnce({ value: [1], '@odata.nextLink': '/next' }).mockRejectedValueOnce(new Error('next failed'));
    const graph = new GraphRequestExecutor(async () => ({ api: jest.fn().mockReturnValue({ get, post }) }), scheduler);
    await graph.api('/$batch').post({ requests: Array(20).fill({ method: 'GET' }) });
    const read = jest.fn();
    await expect(forEachGraphPage(graph, '/events', read)).rejects.toThrow('next failed');
    expect(read.mock.calls).toEqual([[[1]]]);
    expect(scheduler.counts.activeNormal).toBe(0);
  });
});
