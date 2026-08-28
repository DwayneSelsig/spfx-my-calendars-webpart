import type { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IEntraSecurityGroup {
  id: string;
  displayName: string;
}

interface IGroupMembershipCacheEntry {
  expiresAt: number;
  isMember: boolean;
}

export class AudienceService {
  private readonly MEMBERSHIP_CACHE_KEY = 'myCalendarsAudienceMembershipCache';
  private readonly MEMBERSHIP_CACHE_DURATION_MS = 5 * 60 * 1000;
  private readonly CHECK_MEMBER_GROUPS_BATCH_SIZE = 20;
  private graphClient: MSGraphClientV3;

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  public async getSecurityGroups(searchText?: string): Promise<IEntraSecurityGroup[]> {
    const normalizedSearch = (searchText || '').trim().replace(/'/g, "''");
    const filterSegments = ['mailEnabled eq false', 'securityEnabled eq true'];
    if (normalizedSearch) {
      filterSegments.push(`startswith(displayName,'${normalizedSearch}')`);
    }

    try {
      const data = await this.graphClient
        .api('/groups')
        .header('ConsistencyLevel', 'eventual')
        .query({
          $select: 'id,displayName',
          $filter: filterSegments.join(' and '),
          $orderby: 'displayName',
          $top: 50
        })
        .get();

      return (data.value || [])
        .map((item: { id?: string; displayName?: string }) => ({
          id: item.id || '',
          displayName: item.displayName || 'Unnamed group'
        }))
        .filter((item: IEntraSecurityGroup) => !!item.id);
    } catch (error) {
      console.error('Failed to load Entra security groups:', error);
      return [];
    }
  }

  public async getMatchingGroupIds(groupIds: string[]): Promise<Set<string>> {
    const cleanedGroupIds = Array.from(new Set(groupIds.filter(Boolean)));
    if (cleanedGroupIds.length === 0) {
      return new Set();
    }

    const cached = this.getMembershipCache();
    const now = Date.now();
    const matchedGroupIds = new Set<string>();
    const uncachedGroupIds: string[] = [];

    cleanedGroupIds.forEach(groupId => {
      const entry = cached[groupId];
      if (entry && entry.expiresAt > now) {
        if (entry.isMember) {
          matchedGroupIds.add(groupId);
        }
        return;
      }
      uncachedGroupIds.push(groupId);
    });

    for (let index = 0; index < uncachedGroupIds.length; index += this.CHECK_MEMBER_GROUPS_BATCH_SIZE) {
      const batch = uncachedGroupIds.slice(index, index + this.CHECK_MEMBER_GROUPS_BATCH_SIZE);
      try {
        const response = await this.graphClient
          .api('/me/checkMemberGroups')
          .post({ groupIds: batch });

        const batchMatches = new Set<string>((response.value || []) as string[]);
        batch.forEach(groupId => {
          const isMember = batchMatches.has(groupId);
          cached[groupId] = {
            isMember,
            expiresAt: now + this.MEMBERSHIP_CACHE_DURATION_MS
          };

          if (isMember) {
            matchedGroupIds.add(groupId);
          }
        });
      } catch (error) {
        console.error('Failed to evaluate current user group memberships:', error);
      }
    }

    this.setMembershipCache(cached);
    return matchedGroupIds;
  }

  private getMembershipCache(): Record<string, IGroupMembershipCacheEntry> {
    try {
      const raw = sessionStorage.getItem(this.MEMBERSHIP_CACHE_KEY);
      if (!raw) {
        return {};
      }

      const parsed = JSON.parse(raw) as Record<string, IGroupMembershipCacheEntry>;
      return parsed && typeof parsed === 'object' ? parsed : {};
    } catch (error) {
      console.error('Failed to read audience cache:', error);
      return {};
    }
  }

  private setMembershipCache(cache: Record<string, IGroupMembershipCacheEntry>): void {
    try {
      sessionStorage.setItem(this.MEMBERSHIP_CACHE_KEY, JSON.stringify(cache));
    } catch (error) {
      console.error('Failed to write audience cache:', error);
    }
  }
}
