import * as strings from 'MyCalendarsWebPartStrings';
import type { ICalendarSettings, ICalendarSource } from '../models/ICalendarSettings';
import { UserHelper } from '../utils/userHelper';
import { type CalendarRangeCoordinator, type CalendarRangeSource } from './CalendarRangeCoordinator';
import { isAutomaticExchangeCalendarAssigned } from './CalendarSettingsService';
import { ExchangeCalendarService, type IResolvedMailbox } from './ExchangeCalendarService';
import type { GraphClient, GraphRequestExecutor } from './GraphRequestExecutor';
import { GraphRequestScope } from './GraphRequestScheduler';
import { PlannerTaskService } from './PlannerTaskService';
import { SharePointCalendarService } from './SharePointCalendarService';
import { TeamsShiftsService } from './TeamsShiftsService';
import { UnifiedGroupCalendarService } from './UnifiedGroupCalendarService';

function sharedLookup<T>(graph: GraphRequestExecutor, load: (client: GraphClient) => Promise<T>, reusable: (value: T) => boolean = () => true): (scope: GraphRequestScope) => Promise<T> {
  let pending: Promise<T> | undefined;
  let settled = false;
  let scope: GraphRequestScope;
  const users = new Map<GraphRequestScope, () => void>();
  const sync = (): void => {
    const active = Array.from(users.keys()).filter(user => !user.aborted);
    if (!active.length) scope.cancel();
    else scope.setHigh(active.some(user => user.high));
  };
  return user => {
    if (settled && pending) return pending;
    if (!pending || scope.aborted) {
      users.forEach(unsubscribe => unsubscribe());
      users.clear();
      const loadScope = new GraphRequestScope();
      scope = loadScope;
      const finish = (keep: boolean): void => {
        if (scope !== loadScope) return;
        settled = keep;
        if (!keep) pending = undefined;
        users.forEach(unsubscribe => unsubscribe());
        users.clear();
      };
      pending = Promise.resolve().then(() => load(graph.scoped(loadScope))).then(value => {
        finish(reusable(value));
        return value;
      }, error => {
        finish(false);
        throw error;
      });
    }
    if (!users.has(user)) users.set(user, user.subscribe(sync));
    sync();
    return pending;
  };
}

export function startCalendarSources(
  coordinator: CalendarRangeCoordinator, graph: GraphRequestExecutor, settings: ICalendarSettings, tenantId?: string
): void {
  const enabled = settings.sources.filter(source => source.isEnabled);
  const enabledIds = new Set(enabled.filter(source =>
    !(source.sourceType === 'planner' && settings.plannerShowAllCalendars) &&
    !(source.sourceType === 'unifiedGroup' && settings.unifiedGroupShowAllCalendars) &&
    !(source.sourceType === 'teamsShifts' && settings.teamsShiftsShowAllCalendars)
  ).map(source => source.id));
  coordinator.restrictExisting(id => {
    if (id.startsWith('exchange_')) {
      const calendarId = id.substring('exchange_'.length);
      return settings.exchangeCalendarStates?.[calendarId] !== false && !isAutomaticExchangeCalendarAssigned(settings, calendarId);
    }
    if (id.startsWith('auto_planner_')) return settings.plannerShowAllCalendars;
    if (id.startsWith('auto_unifiedGroup_')) return settings.unifiedGroupShowAllCalendars;
    if (id === 'auto_teamsShifts') return settings.teamsShiftsShowAllCalendars;
    return enabledIds.has(id);
  });
  const email = sharedLookup(graph, client => UserHelper.getCurrentUserEmail(client), Boolean);
  const joinedTeams = sharedLookup(graph, client => new TeamsShiftsService(client).getJoinedTeams());
  const userId = sharedLookup(graph, async client => {
    try {
      const user: { id?: string } = await client.api('/me').select('id').get();
      return user.id;
    } catch { return undefined; }
  }, Boolean);
  const mailboxes = new Map<string, (scope: GraphRequestScope) => Promise<IResolvedMailbox>>();
  const siteNames = new Map<string, (scope: GraphRequestScope) => Promise<string | undefined>>();
  const resolveMailbox = (mailbox: string, scope: GraphRequestScope): Promise<IResolvedMailbox> => {
    const key = mailbox.trim().toLowerCase();
    let lookup = mailboxes.get(key);
    if (!lookup) {
      lookup = sharedLookup(graph, client => new ExchangeCalendarService(client).resolveMailbox(mailbox));
      mailboxes.set(key, lookup);
    }
    return lookup(scope);
  };
  const lookupSiteName = (siteId: string, scope: GraphRequestScope): Promise<string | undefined> => {
    let lookup = siteNames.get(siteId);
    if (!lookup) {
      lookup = sharedLookup(graph, async client => (await new SharePointCalendarService(client).getSite(siteId))?.name, Boolean);
      siteNames.set(siteId, lookup);
    }
    return lookup(scope);
  };

  const exchange = (id: string, calendarId: string, name: string, color: string, showLogo: boolean, mailbox?: string): CalendarRangeSource => ({
    id, service: 'exchange',
    load: async (range, scope, onPage) => {
      if (!range) return;
      const service = new ExchangeCalendarService(graph.scoped(scope));
      await service.getCalendarEvents(calendarId, range.start, range.end, mailbox, events => {
        onPage(events.map(event => ({ ...event, sourceId: id, sourceDisplayName: name, colorHex: color,
          sourceType: 'exchange', showSourceLogo: showLogo })));
      }, address => resolveMailbox(address, scope), () => email(scope));
    }
  });

  const planner = (source: ICalendarSource): CalendarRangeSource => ({
    id: source.id, service: 'planner', collection: true,
    load: async (_range, scope, onPage) => {
      await new PlannerTaskService(graph.scoped(scope), tenantId).getTasks(source.plannerPlanId!, undefined, undefined,
        source.plannerAssignedToMeOnly ?? false, source.showCompletedTasks ?? true, source,
        source.showSourceLogo ?? settings.plannerShowSourceLogo ?? true, onPage, () => userId(scope));
    }
  });

  const group = (id: string, groupId: string, name: string, color: string, showLogo: boolean): CalendarRangeSource => ({
    id, service: 'unifiedGroup',
    load: async (range, scope, onPage) => {
      if (!range) return;
      const teams = joinedTeams(scope).then(items => items.some(team => team.id === groupId), () => false);
      const service = new UnifiedGroupCalendarService(graph.scoped(scope));
      const isTeam = await teams;
      await service.getGroupEvents(groupId, range.start, range.end, events => {
        onPage(events.map(event => ({ ...event, sourceId: id, sourceDisplayName: name, colorHex: color,
          sourceType: 'unifiedGroup', showSourceLogo: showLogo, sourceIconName: isTeam ? 'TeamsLogo' : 'Group' })));
      }, () => email(scope));
    }
  });

  const shifts = (source: ICalendarSource): CalendarRangeSource => ({
    id: source.id, service: 'teamsShifts', collection: true,
    load: async (_range, scope, onPage) => {
      await new TeamsShiftsService(graph.scoped(scope)).getShiftsForJoinedTeams(undefined, undefined, source,
        source.showSourceLogo ?? settings.teamsShiftsShowSourceLogo ?? true, onPage, () => joinedTeams(scope));
    }
  });

  coordinator.discover('exchange', async scope => {
    const calendars = await new ExchangeCalendarService(graph.scoped(scope)).getCalendars();
    return calendars.filter(calendar => !isAutomaticExchangeCalendarAssigned(settings, calendar.id)
      && settings.exchangeCalendarStates?.[calendar.id] !== false).map(calendar =>
      exchange(`exchange_${calendar.id}`, calendar.id, calendar.name, calendar.hexColor, settings.exchangeShowSourceLogo ?? true));
  });

  for (const source of enabled) {
    switch (source.sourceType) {
      case 'exchange':
        if (source.exchangeCalendarId) coordinator.register(exchange(source.id, source.exchangeCalendarId, source.name, source.color,
          source.showSourceLogo ?? settings.exchangeShowSourceLogo ?? true, source.exchangeMailbox));
        break;
      case 'sharepoint':
        if (source.sharePointSiteId && source.sharePointListId) coordinator.register({
          id: source.id, service: 'sharepoint', collection: true,
          load: async (_range, scope, onPage) => {
            await new SharePointCalendarService(graph.scoped(scope)).getListEvents(source.sharePointSiteId!, source.sharePointListId!,
              undefined, undefined, source.sharePointFieldMapping, source.sharePointSiteName, events => {
                onPage(events.map(event => ({ ...event, sourceId: source.id, sourceDisplayName: source.name,
                  colorHex: source.color, sourceType: 'sharepoint', showSourceLogo: source.showSourceLogo ?? settings.sharePointShowSourceLogo ?? true })));
              }, () => lookupSiteName(source.sharePointSiteId!, scope));
          }
        });
        break;
      case 'planner':
        if (!settings.plannerShowAllCalendars && source.plannerPlanId) coordinator.register(planner(source));
        break;
      case 'unifiedGroup':
        if (!settings.unifiedGroupShowAllCalendars && source.groupId) coordinator.register(group(source.id, source.groupId, source.name,
          source.color, source.showSourceLogo ?? settings.unifiedGroupShowSourceLogo ?? true));
        break;
      case 'teamsShifts':
        if (!settings.teamsShiftsShowAllCalendars) coordinator.register(shifts(source));
        break;
    }
  }

  if (settings.plannerShowAllCalendars) coordinator.discover('planner', async scope => {
    const plans = await new PlannerTaskService(graph.scoped(scope), tenantId).getUserPlans();
    return plans.map(plan => planner({ id: `auto_planner_${plan.id}`, origin: 'user', sourceType: 'planner', name: plan.title,
      color: settings.organizationPrimaryColor || '#0078d4', isEnabled: true, plannerPlanId: plan.id,
      plannerPlanTitle: plan.title, plannerAssignedToMeOnly: settings.plannerShowAllAssignedToMeOnly ?? false, showCompletedTasks: true }));
  });
  if (settings.unifiedGroupShowAllCalendars) coordinator.discover('unifiedGroup', async scope => {
    const [groups] = await Promise.all([
      new UnifiedGroupCalendarService(graph.scoped(scope)).getUnifiedGroups(),
      joinedTeams(scope)
    ]);
    return groups.map(item => group(`auto_unifiedGroup_${item.id}`, item.id, item.displayName,
      settings.organizationPrimaryColor || '#5b5fc7', settings.unifiedGroupShowSourceLogo ?? true));
  });
  else if (enabled.some(source => source.sourceType === 'unifiedGroup' && source.groupId)) coordinator.discover('unifiedGroup', async scope => {
    await joinedTeams(scope);
    return [];
  });
  if (settings.teamsShiftsShowAllCalendars) coordinator.register(shifts({
    id: 'auto_teamsShifts', origin: 'user', sourceType: 'teamsShifts', name: strings.TeamsShiftsLabel,
    color: settings.organizationPrimaryColor || '#4a4fbe', isEnabled: true
  }));
}
