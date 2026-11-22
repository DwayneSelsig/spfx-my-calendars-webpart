import { defaultAdminWebPartSettings, defaultAllowedOverrides, defaultUserCalendarSettings } from '../models/ICalendarSettings';
import { importAdminWebPartSettingsJson, normalizeAdminWebPartSettings, normalizeUserCalendarSettings, resolveCalendarSettings, serializeAdminWebPartSettings } from './CalendarSettingsService';

describe('administrator configuration files', () => {
  const group = { groupId: 'group-one', displayName: 'Audience' };
  const configuration = {
    ...defaultAdminWebPartSettings,
    organizationPrimaryColor: '#123456', plannerShowAllCalendars: false, maxConcurrentGraphRequests: 17,
    assignedSources: [{
      assignmentId: 'assignment-one', adminSourceId: 'source-one',
      source: { name: 'Calendar', color: '#abcdef', isEnabled: true, sourceType: 'sharepoint' as const,
        sharePointSiteId: 'original-site', sharePointListId: 'original-list',
        sharePointFieldMapping: { titleField: 'Subject', startDateField: 'Start', endDateField: 'End' } },
      audienceGroups: [group], isMandatory: true, defaultEnabled: true,
      allowedOverrides: { ...defaultAllowedOverrides, color: false }
    }],
    exchangeMailboxAssignments: [{ assignmentId: 'mailbox-rule', mailboxId: 'mailbox-one', mailboxDisplayName: 'Mailbox',
      audienceGroups: [group], isMandatory: false, defaultEnabled: true, allowedOverrides: defaultAllowedOverrides,
      exceptions: [{ calendarId: 'calendar-one', excluded: true }, { calendarId: 'calendar-two', isMandatory: true, defaultEnabled: true }] }],
    icsCatalog: [{ adminIcsId: 'subscription-one', displayName: 'Holidays', icsUrl: 'https://example.com/holidays.ics', audienceGroups: [group] }]
  };

  it('round-trips the complete draft with stable references, hidden fields and policy', () => {
    const json = serializeAdminWebPartSettings(configuration);
    const imported = importAdminWebPartSettingsJson(json);
    expect(imported).toEqual(normalizeAdminWebPartSettings(configuration));
    expect(imported?.maxConcurrentGraphRequests).toBe(17);
    expect(serializeAdminWebPartSettings(imported!)).toBe(json);
  });

  it.each([2, 3, 4, 5, 6, 7])('accepts supported schema %s', schemaVersion => {
    const imported = importAdminWebPartSettingsJson(JSON.stringify({ ...defaultAdminWebPartSettings, schemaVersion, maxConcurrentGraphRequests: undefined }));
    expect(imported?.schemaVersion).toBe(7);
    expect(imported?.maxConcurrentGraphRequests).toBe(10);
  });

  it('keeps the scheduler limit under administrator control', () => {
    const personal = normalizeUserCalendarSettings({ ...defaultUserCalendarSettings, maxConcurrentGraphRequests: 20 });
    expect(personal).not.toHaveProperty('maxConcurrentGraphRequests');
    const settings = resolveCalendarSettings({ adminSettings: configuration, userSettings: personal!, matchedGroupIds: new Set<string>() });
    expect(settings.maxConcurrentGraphRequests).toBe(17);
  });

  it.each(['{', 'null', '[]', '{}', JSON.stringify(defaultUserCalendarSettings),
    JSON.stringify({ ...defaultAdminWebPartSettings, schemaVersion: 8 }),
    JSON.stringify({ ...defaultAdminWebPartSettings, schemaVersion: undefined }),
    JSON.stringify({ ...defaultAdminWebPartSettings, assignedSources: {} }),
    JSON.stringify({ ...defaultAdminWebPartSettings, icsCatalog: [{ adminIcsId: 'bad' }] })
  ])('rejects malformed or incompatible input %s', json => {
    expect(importAdminWebPartSettingsJson(json)).toBeUndefined();
  });

  it('rejects missing and partially invalid audiences instead of widening them', () => {
    const persisted = JSON.parse(serializeAdminWebPartSettings(configuration));
    persisted.assignedSources[0].audienceGroupIds.push('unknown');
    expect(importAdminWebPartSettingsJson(JSON.stringify(persisted))).toBeUndefined();
    const hydrated = { ...configuration, assignedSources: [{ ...configuration.assignedSources[0], audienceGroups: [group, {}] }] };
    expect(importAdminWebPartSettingsJson(JSON.stringify(hydrated))).toBeUndefined();
    const malformedIds = { ...configuration, assignedSources: [{ ...configuration.assignedSources[0], audienceGroupIds: 'group-one', audienceGroups: [] }] };
    expect(importAdminWebPartSettingsJson(JSON.stringify(malformedIds))).toBeUndefined();
  });

  it('rejects broken catalog definitions instead of discarding them', () => {
    const persisted = JSON.parse(serializeAdminWebPartSettings(configuration));
    persisted.sourceCatalog[0].source = {};
    expect(importAdminWebPartSettingsJson(JSON.stringify(persisted))).toBeUndefined();
  });

  it('rejects malformed mailbox exceptions instead of silently removing exclusions', () => {
    const persisted = JSON.parse(serializeAdminWebPartSettings(configuration));
    persisted.exchangeMailboxAssignments[0].exceptions.push({ excluded: true });
    expect(importAdminWebPartSettingsJson(JSON.stringify(persisted))).toBeUndefined();
  });
});
