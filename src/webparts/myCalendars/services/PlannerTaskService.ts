import { collectGraphPages, forEachGraphPage, type GraphClient } from './GraphRequestExecutor';
import { GraphRequestCancelled } from './GraphRequestScheduler';
import type { ICalendarEvent as IEvent } from '../models/ICalendarEvent';
import { ICalendarSource } from '../models/ICalendarSettings';
import * as strings from 'MyCalendarsWebPartStrings';
import { formatLocalizedString } from '../utils/localization';

export interface IPlannerPlan {
  id: string;
  title: string;
  owner?: string;
  createdDateTime?: string;
}

interface IGraphPlannerPlan {
  id: string;
  title: string;
  owner?: string;
  createdDateTime?: string;
}

interface IGraphPlannerTask {
  id: string;
  title: string;
  planId: string;
  bucketId?: string;
  percentComplete: number;
  startDateTime?: string;
  dueDateTime?: string;
  completedDateTime?: string;
  assignments?: Record<string, { assignedDateTime: string }>;
  hasDescription?: boolean;
  previewType?: string;
  referenceCount?: number;
  checklistItemCount?: number;
  activeChecklistItemCount?: number;
  conversationThreadId?: string;
  priority?: number;
}

/**
 * Service to interact with Microsoft Planner tasks via Microsoft Graph API
 * Requires Tasks.Read or Tasks.ReadWrite permissions
 */
export class PlannerTaskService {
  private readonly GRAPH_API_URL = 'https://graph.microsoft.com/v1.0';
  private graphClient: GraphClient | null = null;
  private currentUserId: string | null = null;
  private readonly tenantId?: string;

  constructor(graphClient?: GraphClient, tenantId?: string) {
    this.graphClient = graphClient || null;
    this.tenantId = tenantId?.trim() || undefined;
  }

  /**
   * Set the GraphClient (for authentication)
   */
  public setGraphClient(client: GraphClient): void {
    this.graphClient = client;
  }

  /**
   * Get the current user's ID (cached)
   */
  private async getCurrentUserId(): Promise<string | null> {
    if (this.currentUserId) {
      return this.currentUserId;
    }

    try {
      if (!this.graphClient) {
        return null;
      }

      const user = await this.graphClient
        .api('/me')
        .select('id')
        .get();

      this.currentUserId = user.id;
      return this.currentUserId;
    } catch (error) {
      console.error('Error fetching current user ID:', error);
      return null;
    }
  }

  /**
   * Get all Planner plans accessible to the current user
   * Fetches plans from all groups the user is a member of
   */
  public async getUserPlans(): Promise<IPlannerPlan[]> {
    try {
      if (!this.graphClient) {
        throw new Error('GraphClient not initialized');
      }

      const groups = await collectGraphPages<{ id: string }>(this.graphClient, '/me/memberOf/microsoft.graph.group',
        request => request.select('id,displayName').top(100));
      const allPlans: IPlannerPlan[] = [];
      for (const group of groups) {
        try {
          const plans = await collectGraphPages<IGraphPlannerPlan>(this.graphClient, `/groups/${group.id}/planner/plans`);
          allPlans.push(...plans.map(plan => ({ ...plan, owner: plan.owner || group.id })));
        } catch (error) {
          if (error instanceof GraphRequestCancelled) throw error;
          // Some groups do not have Planner enabled.
          console.debug('Could not read Planner plans for a group.');
        }
      }
      return allPlans;
    } catch (error) {
      console.error('Error fetching Planner plans:', error);
      throw error;
    }
  }

  /**
   * Get tasks from a specific Planner plan
   * @param planId - The Planner plan ID
   * @param startDate - Start of date range (optional, for client-side filtering)
   * @param endDate - End of date range (optional, for client-side filtering)
   * @param assignedToMeOnly - If true, only return tasks assigned to current user
   * @param showCompleted - If true, include completed tasks
   * @param source - The calendar source configuration
   * @param showSourceLogo - Whether to show source logo (service-level setting)
   */
  public async getTasks(
    planId: string,
    startDate: Date | undefined,
    endDate: Date | undefined,
    assignedToMeOnly: boolean = false,
    showCompleted: boolean = true,
    source: ICalendarSource,
    showSourceLogo: boolean = true,
    onPage?: (events: IEvent[]) => void,
    getUserId: () => Promise<string | undefined> = async () => (await this.getCurrentUserId()) || undefined
  ): Promise<IEvent[]> {
    try {
      if (!this.graphClient) {
        console.error('GraphClient not initialized');
        return [];
      }

      const events: IEvent[] = [];
      let user: Promise<string | undefined> | undefined;
      await forEachGraphPage<IGraphPlannerTask>(this.graphClient, `/planner/plans/${planId}/tasks`, async page => {
        if (assignedToMeOnly) user = user || getUserId();
        const userId = await user;
        const appointments = page
          .filter(task => (!userId || !!task.assignments && userId in task.assignments) && (showCompleted || task.percentComplete < 100))
          .map(task => this.mapPlannerTaskToAppointment(task, planId, source, showSourceLogo))
          .filter((event): event is IEvent => event !== null)
          .filter(event => (!startDate || Date.parse(event.start) >= startDate.getTime()) &&
            (!endDate || Date.parse(event.start) < endDate.getTime()));
        events.push(...appointments);
        onPage?.(appointments);
      });
      return events;
    } catch (error) {
      console.error('Error fetching Planner tasks:', error);
      throw error;
    }
  }

  /**
   * Map Planner task to IEvent
   * Returns null if the task has no usable date
   */
  private mapPlannerTaskToAppointment(task: IGraphPlannerTask, planId: string, source: ICalendarSource, showSourceLogo: boolean): IEvent | null {
    // Determine dates - use both if available, otherwise use whichever is available
    let startDate: Date | null = null;
    let endDate: Date | null = null;

    if (task.startDateTime) {
      startDate = new Date(task.startDateTime);
    }

    if (task.dueDateTime) {
      endDate = new Date(task.dueDateTime);
    }

    // If we have both, use them
    if (startDate && endDate) {
      // Ensure logical ordering
      if (startDate > endDate) {
        [startDate, endDate] = [endDate, startDate];
      }
    } else if (startDate && !endDate) {
      // Only start date: use it as both start and end
      endDate = new Date(startDate);
    } else if (!startDate && endDate) {
      // Only due date: use it as both start and end
      startDate = new Date(endDate);
    } else {
      // No dates at all - skip this task
      return null;
    }

    // Build description with progress info
    let description = '';
    if (task.percentComplete > 0) {
      description = formatLocalizedString(strings.ProgressLabel, task.percentComplete);
    }
    if (task.checklistItemCount && task.checklistItemCount > 0) {
      const completed = task.checklistItemCount - (task.activeChecklistItemCount || 0);
      description += (description ? '\n' : '') + formatLocalizedString(strings.ChecklistLabel, completed, task.checklistItemCount);
    }

    return {
      id: task.id,
      title: task.title || strings.UntitledTaskLabel,
      description: description || undefined,
      location: undefined, // Planner tasks don't have locations
      start: startDate.toISOString(),
      end: endDate.toISOString(),
      isFullDay: true, // Planner tasks are always all-day (no time component)
      sourceId: source.id,
      sourceDisplayName: source.name,
      colorHex: source.color,
      sourceType: 'planner',
      showSourceLogo: showSourceLogo,
      percentComplete: task.percentComplete,
      webLink: this.getTaskWebLink(planId, task.id)
    };
  }

  private getTaskWebLink(planId: string | undefined, taskId: string | undefined): string | undefined {
    const normalizedPlanId = planId?.trim();
    const normalizedTaskId = taskId?.trim();
    if (!normalizedPlanId || !normalizedTaskId || !this.tenantId) return undefined;

    return `https://planner.cloud.microsoft/webui/v1/plan/${encodeURIComponent(normalizedPlanId)}/task/${encodeURIComponent(normalizedTaskId)}?tid=${encodeURIComponent(this.tenantId)}`;
  }

  /**
   * Verify that a plan exists and is accessible
   * @param planId - The Planner plan ID
   */
  public async verifyPlanAccess(planId: string): Promise<boolean> {
    try {
      if (!this.graphClient) {
        return false;
      }

      await this.graphClient
        .api(`/planner/plans/${planId}`)
        .get();

      return true;
    } catch (error) {
      console.error('Error verifying plan access:', error);
      return false;
    }
  }
}

