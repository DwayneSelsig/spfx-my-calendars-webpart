import { HttpClient, type MSGraphClientV3 } from '@microsoft/sp-http';
import type { ICalendarEvent as IEvent } from '../models/ICalendarEvent';
import { ICalendarSource } from '../models/ICalendarSettings';

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
  private httpClient: HttpClient;
  private graphClient: MSGraphClientV3 | null = null;
  private currentUserId: string | null = null;

  constructor(httpClient: HttpClient, graphClient?: MSGraphClientV3) {
    this.httpClient = httpClient;
    this.graphClient = graphClient || null;
  }

  /**
   * Set the GraphClient (for authentication)
   */
  public setGraphClient(client: MSGraphClientV3): void {
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

      console.log('Fetching Planner plans via groups...');

      // Get all groups the user is a member of
      const groupsData = await this.graphClient
        .api('/me/memberOf/microsoft.graph.group')
        .select('id,displayName')
        .top(100)
        .get();

      console.log(`Found ${groupsData.value?.length || 0} groups`);

      const allPlans: IPlannerPlan[] = [];

      // For each group, try to get its planner plans
      for (const group of (groupsData.value || [])) {
        try {
          const plansData = await this.graphClient
            .api(`/groups/${group.id}/planner/plans`)
            .get();

          if (plansData.value && plansData.value.length > 0) {
            console.log(`Found ${plansData.value.length} plan(s) in group: ${group.displayName}`);
            
            const plans = plansData.value.map((plan: IGraphPlannerPlan) => ({
              id: plan.id,
              title: plan.title,
              owner: plan.owner || group.id,
              createdDateTime: plan.createdDateTime
            }));

            allPlans.push(...plans);
          }
        } catch (groupError) {
          // Some groups might not have Planner enabled, that's okay
          console.log(`No Planner found for group ${group.displayName}:`, groupError);
        }
      }

      console.log(`Total Planner plans found: ${allPlans.length}`);
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
    startDate: Date,
    endDate: Date,
    assignedToMeOnly: boolean = false,
    showCompleted: boolean = true,
    source: ICalendarSource,
    showSourceLogo: boolean = true
  ): Promise<IEvent[]> {
    try {
      if (!this.graphClient) {
        console.error('GraphClient not initialized');
        return [];
      }

      console.log('Fetching Planner tasks for plan:', planId);

      const data = await this.graphClient
        .api(`/planner/plans/${planId}/tasks`)
        .get();

      console.log('Planner tasks response:', data);

      let tasks = data.value || [];

      // Filter by assignment if requested
      if (assignedToMeOnly) {
        const userId = await this.getCurrentUserId();
        if (userId) {
          tasks = tasks.filter((task: IGraphPlannerTask) => {
            return task.assignments && userId in task.assignments;
          });
        }
      }

      // Filter by completion status if requested
      if (!showCompleted) {
        tasks = tasks.filter((task: IGraphPlannerTask) => task.percentComplete < 100);
      }

      // Map to appointments and filter by date range
      const appointments = tasks
        .map((task: IGraphPlannerTask) => this.mapPlannerTaskToAppointment(task, source, showSourceLogo))
        .filter((apt: IEvent | null): apt is IEvent => apt !== null);

      // Client-side date filtering (only include tasks with dates in range)
      return appointments.filter((apt: IEvent) => {
        const aptDate = new Date(apt.start);
        return aptDate >= startDate && aptDate <= endDate;
      });
    } catch (error) {
      console.error('Error fetching Planner tasks:', error);
      throw error;
    }
  }

  /**
   * Map Planner task to IEvent
   * Returns null if the task has no usable date
   */
  private mapPlannerTaskToAppointment(task: IGraphPlannerTask, source: ICalendarSource, showSourceLogo: boolean): IEvent | null {
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
      console.log('Skipping Planner task without dates:', task.title);
      return null;
    }

    // Build description with progress info
    let description = '';
    if (task.percentComplete > 0) {
      description = `Progress: ${task.percentComplete}%`;
    }
    if (task.checklistItemCount && task.checklistItemCount > 0) {
      const completed = task.checklistItemCount - (task.activeChecklistItemCount || 0);
      description += (description ? '\n' : '') + `Checklist: ${completed}/${task.checklistItemCount} completed`;
    }

    return {
      id: task.id,
      title: task.title || 'Untitled Task',
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
      percentComplete: task.percentComplete
    };
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

