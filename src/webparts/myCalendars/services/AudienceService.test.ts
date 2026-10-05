import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { AudienceService } from './AudienceService';

function client(pages: unknown[]): { graph: MSGraphClientV3; api: jest.Mock; request: { header: jest.Mock; query: jest.Mock; get: jest.Mock } } {
  const request = { header: jest.fn().mockReturnThis(), query: jest.fn().mockReturnThis(), get: jest.fn() };
  pages.forEach(page => request.get.mockResolvedValueOnce(page));
  const api = jest.fn().mockReturnValue(request);
  return { graph: { api } as unknown as MSGraphClientV3, api, request };
}

describe('audience discovery', () => {
  afterEach(() => jest.restoreAllMocks());

  it('escapes prefix searches, pages all results and sorts locally without orderby/count', async () => {
    const { graph, api, request } = client([
      { value: [{ id: 'z', displayName: 'Zebra' }, { id: 'b', displayName: 'Alpha' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/groups?$skiptoken=next' },
      { value: [{ id: 'a', displayName: 'Alpha' }, { id: 'z', displayName: 'Zebra' }] }
    ]);
    const results = await new AudienceService(graph).getSecurityGroups(" O'Brien ");
    expect(request.query).toHaveBeenCalledWith({
      $select: 'id,displayName', $filter: "mailEnabled eq false and securityEnabled eq true and startswith(displayName,'O''Brien')", $top: 50
    });
    expect(api).toHaveBeenCalledWith('https://graph.microsoft.com/v1.0/groups?$skiptoken=next');
    expect(request.header).toHaveBeenCalledTimes(2);
    expect(request.header).toHaveBeenCalledWith('ConsistencyLevel', 'eventual');
    expect(results.map(group => group.id)).toEqual(['a', 'b', 'z']);
  });

  it('distinguishes successful empty results from first-page and later-page API failures', async () => {
    jest.spyOn(console, 'error').mockImplementation(() => undefined);
    const empty = client([{ value: [] }]);
    expect(await new AudienceService(empty.graph).getSecurityGroups(' ')).toEqual([]);
    expect(empty.request.query.mock.calls[0][0].$filter).toBe('mailEnabled eq false and securityEnabled eq true');
    const failed = client([]);
    const failure = { statusCode: 403, code: 'Authorization_RequestDenied' };
    failed.request.get.mockRejectedValue(failure);
    await expect(new AudienceService(failed.graph).getSecurityGroups()).rejects.toBe(failure);
    const later = client([{ value: [{ id: 'a' }], '@odata.nextLink': 'next' }]);
    later.request.get.mockRejectedValueOnce(failure);
    await expect(new AudienceService(later.graph).getSecurityGroups()).rejects.toBe(failure);
  });

  it('preserves transitive matching, batching, OR inputs and fail-closed membership failures', async () => {
    jest.spyOn(console, 'error').mockImplementation(() => undefined);
    const setItem = jest.fn();
    const storage = { getItem: jest.fn().mockReturnValue(null), setItem };
    const originalStorage = Object.getOwnPropertyDescriptor(globalThis, 'sessionStorage');
    Object.defineProperty(globalThis, 'sessionStorage', { value: storage, configurable: true });
    try {
      const post = jest.fn().mockResolvedValueOnce({ value: ['group0'] }).mockRejectedValueOnce(new Error('403'));
      const graph = { api: jest.fn().mockReturnValue({ post }) } as unknown as MSGraphClientV3;
      const ids = Array.from({ length: 21 }, (_, index) => `group${index}`);
      expect(await new AudienceService(graph).getMatchingGroupIds(ids)).toEqual(new Set(['group0']));
      expect(post.mock.calls[0][0].groupIds).toHaveLength(20);
      expect(post.mock.calls[1][0].groupIds).toEqual(['group20']);
      const cache = JSON.parse(setItem.mock.calls[0][1]);
      expect(cache.group0.isMember).toBe(true);
      expect(cache.group1.isMember).toBe(false);
      expect(cache.group20).toBeUndefined();
    } finally {
      if (originalStorage) Object.defineProperty(globalThis, 'sessionStorage', originalStorage);
      else delete (globalThis as unknown as { sessionStorage?: unknown }).sessionStorage;
    }
  });
});
