import { formatAppointmentDuration } from './searchFormatting';

describe('search result formatting', () => {
  it('uses the localized all-day label for all-day appointments', () => {
    expect(formatAppointmentDuration(new Date('2026-01-01T00:00:00Z'), new Date('2026-01-02T00:00:00Z'), true)).toBe('All day');
  });

  it('formats timed appointment durations without hardcoded plural text', () => {
    expect(formatAppointmentDuration(new Date('2026-01-01T09:00:00Z'), new Date('2026-01-01T10:30:00Z'), false)).toBe('1 hr 30 min');
    expect(formatAppointmentDuration(new Date('2026-01-01T09:00:00Z'), new Date('2026-01-01T09:45:00Z'), false)).toBe('45 min');
  });
});
