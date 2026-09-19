import type { ICalendarSource } from '../models/ICalendarSettings';
import { TeamsShiftsService } from './TeamsShiftsService';

describe('TeamsShiftsService localized descriptions', () => {
  it('uses the localized activities template', async () => {
    const joinedTeamsRequest = {
      select: jest.fn().mockReturnThis(),
      get: jest.fn().mockResolvedValue({ value: [{ id: 'team-id', displayName: 'Operations' }] })
    };
    const shiftsRequest = {
      get: jest.fn().mockResolvedValue({ value: [{
        id: 'shift-id',
        sharedShift: {
          displayName: 'Morning',
          startDateTime: '2026-09-10T08:00:00Z',
          endDateTime: '2026-09-10T10:00:00Z',
          activities: [{ displayName: 'Opening' }, { displayName: 'Handover' }]
        }
      }] })
    };
    const graphClient = {
      api: jest.fn((path: string) => path === '/me/joinedTeams' ? joinedTeamsRequest : shiftsRequest)
    };
    const source = { id: 'teams-shifts', name: 'Teams Shifts', color: '#4a4fbe', sourceType: 'teamsShifts', isEnabled: true } as ICalendarSource;
    const service = new TeamsShiftsService(graphClient as never);

    const events = await service.getShiftsForJoinedTeams(new Date('2026-09-01'), new Date('2026-10-01'), source);

    expect(events[0].description).toBe('Activities: Opening, Handover');
  });
});
