import { formatLocalizedString } from './localization';

describe('localization formatters', () => {
  it('replaces all numbered placeholders', () => {
    expect(formatLocalizedString('Page {0} of {1}', 2, 4)).toBe('Page 2 of 4');
    expect(formatLocalizedString('Showing {0} of {1} sites', 20, 42)).toBe('Showing 20 of 42 sites');
    expect(formatLocalizedString('Admin defaults: {0} calendars, {1} ICS catalog items, default view: {2}', 3, 1, 'Week')).toBe('Admin defaults: 3 calendars, 1 ICS catalog items, default view: Week');
    expect(formatLocalizedString('No results found for "{0}"', 'planning')).toBe('No results found for "planning"');
    expect(formatLocalizedString('{0} minutes · Weekends: {1}', 30, 'shown')).toBe('30 minutes · Weekends: shown');
  });
});
