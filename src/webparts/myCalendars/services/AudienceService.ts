import type { MSGraphClientV3 } from '@microsoft/sp-http';
import * as strings from 'MyCalendarsWebPartStrings';

export interface IEntraAudienceGroup {
  id: string;
  displayName: string;
  groupType?: import('../models/ICalendarSettings').AudienceGroupType;
}

/** Legacy type name retained for existing discovery-state consumers. */
export type IEntraSecurityGroup = IEntraAudienceGroup;

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

  public async getAudienceGroups(searchText?: string): Promise<IEntraSecurityGroup[]> {
    const normalizedSearch = (searchText || '').trim().replace(/'/g, "''");
    const filterSegments = ["(securityEnabled eq true or groupTypes/any(c:c eq 'Unified'))"];
    if (normalizedSearch) {
      filterSegments.push(`startswith(displayName,'${normalizedSearch}')`);
    }

    try {
      let data = await this.graphClient
        .api('/groups')
        .header('ConsistencyLevel', 'eventual')
        .query({ $select: 'id,displayName,groupTypes,mailEnabled,securityEnabled', $filter: filterSegments.join(' and '), $top: 50, $count: 'true' })
        .get();
      const groups = new Map<string, IEntraSecurityGroup>();
      for (;;) {
        (data.value || []).forEach((item: { id?: string; displayName?: string; groupTypes?: string[]; securityEnabled?: boolean; mailEnabled?: boolean }) => {
          const groupType = (item.groupTypes || []).indexOf('Unified') >= 0 ? 'microsoft365'
            : item.securityEnabled === true ? item.mailEnabled === true ? 'mailEnabledSecurity' : 'security' : undefined;
          if (item.id && groupType) groups.set(item.id, { id: item.id, displayName: item.displayName || strings.UnnamedSecurityGroupLabel, groupType });
        });
        if (!data['@odata.nextLink']) break;
        data = await this.graphClient.api(data['@odata.nextLink']).header('ConsistencyLevel', 'eventual').get();
      }
      return Array.from(groups.values()).sort((a, b) => a.displayName.localeCompare(b.displayName) || a.id.localeCompare(b.id));
    } catch (error) {
      console.error('Failed to load Entra security groups:', error);
      throw error;
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
