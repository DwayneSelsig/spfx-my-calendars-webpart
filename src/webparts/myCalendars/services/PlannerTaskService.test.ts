import type { ICalendarSource } from '../models/ICalendarSettings';
import { PlannerTaskService } from './PlannerTaskService';

describe('PlannerTaskService localized descriptions', () => {
  it('uses localized fallback titles and composed progress/checklist descriptions', async () => {
    const request = {
      get: jest.fn().mockResolvedValue({ value: [{
        id: 'task/id',
        title: '',
        startDateTime: '2026-09-10T08:00:00Z',
        dueDateTime: '2026-09-10T09:00:00Z',
        percentComplete: 50,
        checklistItemCount: 2,
        activeChecklistItemCount: 1
      }] })
    };
    const source = {
      id: 'planner', name: 'Planner', color: '#0078d4', sourceType: 'planner', isEnabled: true,
      plannerPlanId: 'plan/id'
    } as ICalendarSource;
    const service = new PlannerTaskService({ api: jest.fn().mockReturnValue(request) } as never, 'tenant/id');

    const events = await service.getTasks('plan/id', new Date('2026-09-01'), new Date('2026-10-01'), false, true, source);

    expect(events[0]).toMatchObject({
      title: 'Untitled Task',
      description: 'Progress: 50%\nChecklist: 1/2 completed',
      webLink: 'https://planner.cloud.microsoft/webui/v1/plan/plan%2Fid/task/task%2Fid?tid=tenant%2Fid'
    });
  });

  it('omits the task link when the tenant ID is unavailable', async () => {
    const request = {
      get: jest.fn().mockResolvedValue({ value: [{
        id: 'task-id', title: 'Task', planId: 'plan-id', dueDateTime: '2026-09-10T09:00:00Z', percentComplete: 0
      }] })
    };
    const source = {
      id: 'planner', name: 'Planner', color: '#0078d4', sourceType: 'planner', isEnabled: true,
      plannerPlanId: 'plan-id'
    } as ICalendarSource;
    const service = new PlannerTaskService({ api: jest.fn().mockReturnValue(request) } as never);

    const events = await service.getTasks('plan-id', new Date('2026-09-01'), new Date('2026-10-01'), false, true, source);

    expect(events[0].webLink).toBeUndefined();
  });
});
