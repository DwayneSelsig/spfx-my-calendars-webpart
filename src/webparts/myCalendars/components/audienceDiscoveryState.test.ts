import { updateAudienceDiscovery, type IAudienceDiscoveryState } from './audienceDiscoveryState';

describe('audience results and selection separation', () => {
  it('retains selected IDs/names across filtered/empty results, edited input and API failures', () => {
    const original: IAudienceDiscoveryState = {
      securityGroups: [{ id: 'old', displayName: 'Old result' }], securityGroupsLoading: false,
      selectedAudienceGroups: { hidden: 'Selected outside search' }
    };
    let state = updateAudienceDiscovery(original, { type: 'start' });
    expect(state.securityGroups).toEqual([]);
    expect(state.securityGroupsLoaded).toBe(false);
    state = updateAudienceDiscovery(state, { type: 'success', groups: [{ id: 'new', displayName: 'New result' }] });
    state = updateAudienceDiscovery(state, { type: 'finish' });
    expect(state.selectedAudienceGroups).toEqual(original.selectedAudienceGroups);
    state = updateAudienceDiscovery(state, { type: 'input' });
    expect(state.securityGroupsLoading).toBe(false);
    expect(state.securityGroupsLoaded).toBe(false);
    state = updateAudienceDiscovery(state, { type: 'start' });
    state = updateAudienceDiscovery(state, { type: 'error', message: 'Graph failed' });
    state = updateAudienceDiscovery(state, { type: 'finish' });
    expect(state.securityGroupsError).toBe('Graph failed');
    expect(state.securityGroups).toEqual([]);
    expect(state.selectedAudienceGroups).toEqual(original.selectedAudienceGroups);
    state = updateAudienceDiscovery(state, { type: 'start' });
    state = updateAudienceDiscovery(state, { type: 'success', groups: [] });
    state = updateAudienceDiscovery(state, { type: 'finish' });
    expect(state.securityGroupsError).toBeUndefined();
    expect(state.securityGroupsLoaded).toBe(true);
    expect(state.selectedAudienceGroups).toEqual(original.selectedAudienceGroups);
    expect(original.securityGroups).toHaveLength(1);
  });
});
