import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { AudienceService } from './AudienceService';

function client(pages: unknown[]): { graph: MSGraphClientV3; api: jest.Mock; request: { header: jest.Mock; query: jest.Mock; get: jest.Mock } } {
  const request = { header: jest.fn().mockReturnThis(), query: jest.fn().mockReturnThis(), get: jest.fn() };
  pages.forEach(page => request.get.mockResolvedValueOnce(page));
  const api = jest.fn().mockReturnValue(request);
  return { graph: { api } as unknown as MSGraphClientV3, api, request };
}

describe('audience discovery', () => {
  it('retains supported group types and excludes distribution groups', async () => {
    const { graph } = client([{ value: [
      { id: 'm365', displayName: 'Microsoft 365', groupTypes: ['Unified'], mailEnabled: true, securityEnabled: false },
      { id: 'security', displayName: 'Security', groupTypes: [], mailEnabled: false, securityEnabled: true },
      { id: 'mailsecurity', displayName: 'Mail security', groupTypes: [], mailEnabled: true, securityEnabled: true },
      { id: 'distribution', displayName: 'Distribution', groupTypes: [], mailEnabled: true, securityEnabled: false },
      { id: 'dynamicDistribution', displayName: 'Dynamic distribution', groupTypes: ['DynamicMembership'], mailEnabled: true, securityEnabled: false }
    ] }]);
    expect((await new AudienceService(graph).getAudienceGroups()).map(group => [group.id, group.groupType])).toEqual([
      ['mailsecurity', 'mailEnabledSecurity'], ['m365', 'microsoft365'], ['security', 'security']
    ]);
  });

  afterEach(() => jest.restoreAllMocks());

  it('escapes prefix searches, pages all results and sorts locally without orderby', async () => {
    const { graph, api, request } = client([
      { value: [{ id: 'z', displayName: 'Zebra', securityEnabled: true, mailEnabled: false }, { id: 'b', displayName: 'Alpha', securityEnabled: true, mailEnabled: false }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/groups?$skiptoken=next' },
      { value: [{ id: 'a', displayName: 'Alpha', securityEnabled: true, mailEnabled: false }, { id: 'z', displayName: 'Zebra', securityEnabled: true, mailEnabled: false }] }
    ]);
    const results = await new AudienceService(graph).getAudienceGroups(" O'Brien ");
    expect(request.query).toHaveBeenCalledWith({
      $select: 'id,displayName,groupTypes,mailEnabled,securityEnabled', $filter: "(securityEnabled eq true or groupTypes/any(c:c eq 'Unified')) and startswith(displayName,'O''Brien')", $top: 50, $count: 'true'
    });
    expect(api).toHaveBeenCalledWith('https://graph.microsoft.com/v1.0/groups?$skiptoken=next');
    expect(request.header).toHaveBeenCalledTimes(2);
    expect(request.header).toHaveBeenCalledWith('ConsistencyLevel', 'eventual');
    expect(results.map(group => group.id)).toEqual(['a', 'b', 'z']);
  });

  it('distinguishes successful empty results from first-page and later-page API failures', async () => {
    jest.spyOn(console, 'error').mockImplementation(() => undefined);
    const empty = client([{ value: [] }]);
    expect(await new AudienceService(empty.graph).getAudienceGroups(' ')).toEqual([]);
    expect(empty.request.query.mock.calls[0][0].$filter).toBe("(securityEnabled eq true or groupTypes/any(c:c eq 'Unified'))");
    const failed = client([]);
    const failure = { statusCode: 403, code: 'Authorization_RequestDenied' };
    failed.request.get.mockRejectedValue(failure);
    await expect(new AudienceService(failed.graph).getAudienceGroups()).rejects.toBe(failure);
    const later = client([{ value: [{ id: 'a' }], '@odata.nextLink': 'next' }]);
    later.request.get.mockRejectedValueOnce(failure);
    await expect(new AudienceService(later.graph).getAudienceGroups()).rejects.toBe(failure);
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
