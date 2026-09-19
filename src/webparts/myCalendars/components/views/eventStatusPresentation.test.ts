import type { CalendarAvailabilityStatus, ICalendarEvent } from '../../models/ICalendarEvent';
import { getEventStatusPresentation } from './eventStatusPresentation';

const baseEvent: ICalendarEvent = { id: '1', sourceId: 'source', title: 'Event', start: '2026-01-01T09:00:00Z', end: '2026-01-01T10:00:00Z' };

describe('event status presentation', () => {
  it.each<CalendarAvailabilityStatus>(['free', 'tentative', 'busy', 'oof', 'workingElsewhere', 'unknown'])('provides a non-color glyph for %s', showAs => {
    const result = getEventStatusPresentation({ ...baseEvent, showAs }, '#0078d4');
    expect(result.availabilityIconName).toBeTruthy();
    expect(result.cardStyle.background).toBeTruthy();
  });

  it('keeps response separate and visibly mutes declined events', () => {
    const result = getEventStatusPresentation({ ...baseEvent, showAs: 'busy', responseStatus: 'declined' }, '#0078d4');
    expect(result.availabilityIconName).toBe('CircleFill');
    expect(result.responseIconName).toBe('StatusCircleErrorX');
    expect(result.titleStyle.textDecoration).toBe('line-through');
  });

  it('uses the theme surface without a calendar tint for free events', () => {
    const result = getEventStatusPresentation({ ...baseEvent, showAs: 'free' }, '#0078d4');
    expect(result.cardStyle.background).toBe('transparent');
  });

  it('uses diagonal bands for tentative and cross-hatching for out of office', () => {
    const tentative = getEventStatusPresentation({ ...baseEvent, showAs: 'tentative' }, '#0078d4');
    const oof = getEventStatusPresentation({ ...baseEvent, showAs: 'oof' }, '#0078d4');

    expect(tentative.cardStyle.background).toContain('repeating-linear-gradient(135deg');
    expect(oof.cardStyle.background).toContain('repeating-linear-gradient(45deg');
    expect(oof.cardStyle.background).toContain('repeating-linear-gradient(-45deg');
  });

  it('leaves statusless events on their existing base treatment', () => {
    const result = getEventStatusPresentation(baseEvent, '#0078d4', 16);
    expect(result.availabilityIconName).toBeUndefined();
    expect(result.responseIconName).toBeUndefined();
    expect(result.cardStyle.borderTop).toBeUndefined();
  });
});
