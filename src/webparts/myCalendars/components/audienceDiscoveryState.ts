import type { IEntraSecurityGroup } from '../services/AudienceService';

export interface IAudienceDiscoveryState {
  securityGroups: IEntraSecurityGroup[];
  securityGroupsLoading: boolean;
  securityGroupsError?: string;
  securityGroupsLoaded?: boolean;
  selectedAudienceGroups: Record<string, string>;
}

type AudienceDiscoveryEvent =
  | { type: 'start' | 'input' }
  | { type: 'success'; groups: IEntraSecurityGroup[] }
  | { type: 'error'; message: string }
  | { type: 'finish' };

/** Discovery owns results/status, never the separate group selection. */
export function updateAudienceDiscovery(state: IAudienceDiscoveryState, event: AudienceDiscoveryEvent): IAudienceDiscoveryState {
  switch (event.type) {
    case 'start':
    case 'input':
      return { ...state, securityGroups: [], securityGroupsLoading: event.type === 'start', securityGroupsError: undefined, securityGroupsLoaded: false };
    case 'success':
      return { ...state, securityGroups: event.groups, securityGroupsLoaded: true };
    case 'error':
      return { ...state, securityGroups: [], securityGroupsError: event.message, securityGroupsLoaded: false };
    case 'finish':
      return { ...state, securityGroupsLoading: false };
  }
}
