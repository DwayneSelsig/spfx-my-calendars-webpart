import { UserHelper } from '../utils/userHelper';
import { UnifiedGroupCalendarService } from './UnifiedGroupCalendarService';

describe('UnifiedGroupCalendarService event status mapping', () => {
  afterEach(() => jest.restoreAllMocks());

  it.each(['free', 'tentative', 'busy', 'oof', 'workingElsewhere', 'unknown'])('maps group availability %s without a personal response', async showAs => {
    jest.spyOn(UserHelper, 'getCurrentUserEmail').mockResolvedValue('viewer@example.com');
    const request = {
      header: jest.fn().mockReturnThis(), query: jest.fn().mockReturnThis(),
      get: jest.fn().mockResolvedValue({ value: [{
        id: 'event', subject: 'Group event', bodyPreview: '',
        start: { dateTime: '2026-09-10T08:00:00Z' }, end: { dateTime: '2026-09-10T09:00:00Z' },
        isReminderOn: false, attendees: [], showAs, responseStatus: { response: 'accepted' },
        webLink: 'https://outlook.office.com/calendar/item/group-event'
      }] })
    };
    const service = new UnifiedGroupCalendarService({ api: jest.fn().mockReturnValue(request) });
    const events = await service.getGroupEvents('group-id', new Date('2026-09-01'), new Date('2026-10-01'));
    expect(events[0].showAs).toBe(showAs);
    expect(events[0].responseStatus).toBeUndefined();
    expect(events[0].webLink).toBe('https://outlook.office.com/calendar/item/group-event');
  });
});
