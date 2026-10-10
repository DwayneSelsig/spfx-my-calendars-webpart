import { loadExchangeSources } from '../components/exchangeLoading';
import { defaultAdminWebPartSettings, defaultUserCalendarSettings, defaultAllowedOverrides, IAdminAssignedSource, IAdminWebPartSettings, ICalendarSourceBase, IExchangeMailboxAssignment } from '../models/ICalendarSettings';
import { addExchangeMailboxAssignment, addAdministratorAssignments, canonicalizeExchangeSourceIdentities, applyAdminSourceChanges, cleanAdminSourceOverrides, deriveUserCalendarSettings, getSourceIdentityKey, isAutomaticExchangeCalendarAssigned, normalizeAdminWebPartSettings, normalizeAssignmentPolicy, normalizeUserCalendarSettings, resolveCalendarSettings, serializeAdminWebPartSettings } from './CalendarSettingsService';

const source: ICalendarSourceBase = { sourceType: 'exchange', exchangeMailbox: 'mailbox', exchangeCalendarId: 'calendar-id', name: 'Calendar', color: '#0078d4', isEnabled: true };
function assignment(groupId: string, isMandatory = false, defaultEnabled = true): IAdminAssignedSource {
  return { assignmentId: 'assignment-' + groupId, adminSourceId: 'original-id', source: { ...source }, audienceGroups: groupId ? [{ groupId, displayName: groupId }] : [], isMandatory, defaultEnabled, allowedOverrides: { ...defaultAllowedOverrides } };
}
function admin(assignments: IAdminAssignedSource[]): IAdminWebPartSettings { return { ...defaultAdminWebPartSettings, assignedSources: assignments }; }
function rule(): IExchangeMailboxAssignment { return { assignmentId: 'rule', mailboxId: 'mailbox', mailboxDisplayName: 'mailbox@example.com', audienceGroups: [{ groupId: 'a', displayName: 'A' }], isMandatory: false, defaultEnabled: true, allowedOverrides: { ...defaultAllowedOverrides }, exceptions: [] }; }
const matched = new Set(['a', 'b']);

describe('administrator assignment policy', () => {
  it.each([[true, true, true], [false, true, true], [false, false, false], [true, false, true]])('normalizes mandatory=%s enabled=%s', (isMandatory, defaultEnabled, expected) => {
    expect(normalizeAssignmentPolicy({ isMandatory, defaultEnabled })).toEqual({ isMandatory, defaultEnabled: expected });
  });
  it.each([[true, true, false, true], [false, true, undefined, true], [false, false, undefined, false], [false, true, false, false], [false, false, true, true]])('resolves mandatory=%s default=%s override=%s', (mandatory, enabled, override, expected) => {
    const result = resolveCalendarSettings({ adminSettings: admin([assignment('a', mandatory, enabled)]), matchedGroupIds: matched,
      userSettings: { ...defaultUserCalendarSettings, adminSourceOverridesById: { 'original-id': { isEnabled: override } } } });
    expect(result.sources[0].isEnabled).toBe(expected);
  });
  it('applies Everyone without membership and fails closed for group targets', () => {
    const result = resolveCalendarSettings({ adminSettings: admin([assignment(''), { ...assignment('a'), adminSourceId: 'other', source: { ...source, exchangeCalendarId: 'other' } }]), matchedGroupIds: new Set(), userSettings: defaultUserCalendarSettings });
    expect(result.sources.map(item => item.id)).toEqual(['original-id']);
  });
  it.each([[true, true, false, true, true, true], [true, true, false, false, true, true], [false, true, false, false, false, true]])('deduplicates and combines applicable policies', (m1, d1, m2, d2, expectedMandatory, expectedDefault) => {
    const result = resolveCalendarSettings({ adminSettings: admin([assignment('a', m1, d1), assignment('b', m2, d2)]), matchedGroupIds: matched, userSettings: defaultUserCalendarSettings });
    expect(result.sources).toHaveLength(1);
    expect(result.sources[0]).toMatchObject({ isMandatory: expectedMandatory, defaultEnabled: expectedDefault, audienceGroupNames: ['a', 'b'] });
  });
  it('preserves explicit visibility equal to the default across unrelated saves and default changes', () => {
    const settings = admin([assignment('a')]);
    const effective = resolveCalendarSettings({ adminSettings: settings, matchedGroupIds: matched, userSettings: defaultUserCalendarSettings });
    effective.sources[0] = applyAdminSourceChanges(effective.sources[0], { isEnabled: true });
    const user = deriveUserCalendarSettings({ nextResolvedSettings: effective, adminSettings: settings, matchedGroupIds: matched });
    expect(user.adminSourceOverridesById['original-id'].isEnabled).toBe(true);
    const changed = admin([assignment('a', false, false)]);
    const next = resolveCalendarSettings({ adminSettings: changed, matchedGroupIds: matched, userSettings: user });
    expect(next.sources[0].isEnabled).toBe(true);
    const again = deriveUserCalendarSettings({ nextResolvedSettings: next, adminSettings: changed, matchedGroupIds: matched, existingUserSettings: user });
    expect(again.adminSourceOverridesById['original-id'].isEnabled).toBe(true);
    next.sources[0] = applyAdminSourceChanges(next.sources[0], { visibilityOverride: undefined });
    const reset = deriveUserCalendarSettings({ nextResolvedSettings: next, adminSettings: changed, matchedGroupIds: matched, existingUserSettings: user });
    expect(reset.adminSourceOverridesById['original-id']).toBeUndefined();
  });
  it('does not create overrides for untouched defaults', () => {
    const settings = admin([assignment('a')]);
    const effective = resolveCalendarSettings({ adminSettings: settings, matchedGroupIds: matched, userSettings: defaultUserCalendarSettings });
    expect(deriveUserCalendarSettings({ nextResolvedSettings: effective, adminSettings: settings, matchedGroupIds: matched }).adminSourceOverridesById).toEqual({});
  });
  it('cleans observed disallowed choices so they do not reappear when policy is relaxed', () => {
    const settings = admin([{ ...assignment('a', true), allowedOverrides: { ...defaultAllowedOverrides, color: false } }]);
    const user = { ...defaultUserCalendarSettings, adminSourceOverridesById: { 'original-id': { isEnabled: false, removed: true, color: '#ffffff', name: 'Personal' }, orphan: { isEnabled: true } } };
    const effective = resolveCalendarSettings({ adminSettings: settings, matchedGroupIds: matched, userSettings: user });
    expect(effective.sources[0]).toMatchObject({ isEnabled: true, color: source.color, name: 'Personal' });
    const cleaned = cleanAdminSourceOverrides(user, effective.applicableAdminSources || []);
    expect(cleaned.adminSourceOverridesById).toEqual({ 'original-id': { name: 'Personal' } });
    const relaxed = resolveCalendarSettings({ adminSettings: admin([assignment('a')]), matchedGroupIds: matched, userSettings: cleaned });
    expect(relaxed.sources[0]).toMatchObject({ isEnabled: true, color: source.color, name: 'Personal' });
  });
  it('rejects programmatic mandatory disabling and locked presentation changes', () => {
    const settings = admin([{ ...assignment('a', true), allowedOverrides: { ...defaultAllowedOverrides, name: false } }]);
    const effective = resolveCalendarSettings({ adminSettings: settings, matchedGroupIds: matched, userSettings: defaultUserCalendarSettings });
    expect(applyAdminSourceChanges(effective.sources[0], { name: 'Forbidden', isEnabled: false, color: '#ffffff' })).toMatchObject({ name: source.name, isEnabled: true, color: '#ffffff' });
    effective.sources = [];
    expect(deriveUserCalendarSettings({ nextResolvedSettings: effective, adminSettings: settings, matchedGroupIds: matched }).adminSourceOverridesById).toEqual({});
  });
  it('deduplicates personal sources without losing their persisted IDs', () => {
    const settings = admin([assignment('a')]);
    const user = { ...defaultUserCalendarSettings, personalSources: [{ ...source, userSourceId: 'personal-id' }] };
    const effective = resolveCalendarSettings({ adminSettings: settings, matchedGroupIds: matched, userSettings: user });
    expect(effective.sources).toHaveLength(1);
    expect(deriveUserCalendarSettings({ nextResolvedSettings: effective, adminSettings: settings, matchedGroupIds: matched, existingUserSettings: user }).personalSources[0].userSourceId).toBe('personal-id');
  });
});

describe('Exchange mailbox assignments', () => {
  it('assigns future discovered calendars only when all-calendar rules exist', () => {
    const discoveries = [{ mailboxId: 'mailbox', sources: [source, { ...source, exchangeCalendarId: 'future', name: 'Future' }] }];
    const explicit = resolveCalendarSettings({ adminSettings: admin([assignment('a')]), matchedGroupIds: matched, userSettings: defaultUserCalendarSettings, mailboxDiscoveries: discoveries });
    expect(explicit.sources).toHaveLength(1);
    const all = resolveCalendarSettings({ adminSettings: { ...admin([]), exchangeMailboxAssignments: [rule()] }, matchedGroupIds: matched, userSettings: defaultUserCalendarSettings, mailboxDiscoveries: discoveries });
    expect(all.sources).toHaveLength(2);
    expect(all.sources[1].defaultEnabled).toBe(true);
  });
  it('applies calendar exceptions locally and lets another group assign an excluded calendar', () => {
    const mailboxRule = { ...rule(), isMandatory: true, exceptions: [{ calendarId: 'calendar-id', excluded: true }, { calendarId: 'future', isMandatory: false, defaultEnabled: false }] };
    const result = resolveCalendarSettings({ adminSettings: { ...admin([assignment('b')]), exchangeMailboxAssignments: [mailboxRule] }, matchedGroupIds: matched, userSettings: defaultUserCalendarSettings,
      mailboxDiscoveries: [{ mailboxId: 'mailbox', sources: [source, { ...source, exchangeCalendarId: 'future' }] }] });
    expect(result.sources).toHaveLength(2);
    expect(result.sources[0].isMandatory).toBe(false);
    expect(result.sources[1]).toMatchObject({ isMandatory: false, isEnabled: false });
  });
  it('uses shared definitions and existing IDs when explicit and mailbox rules overlap', () => {
    const settings = { ...admin([assignment('a', false, false)]), exchangeMailboxAssignments: [{ ...rule(), isMandatory: true }] };
    const result = resolveCalendarSettings({ adminSettings: settings, matchedGroupIds: matched, userSettings: defaultUserCalendarSettings, mailboxDiscoveries: [{ mailboxId: 'mailbox', sources: [{ ...source, name: 'Renamed' }] }] });
    expect(result.sources).toHaveLength(1);
    expect(result.sources[0]).toMatchObject({ id: 'original-id', name: source.name, isMandatory: true });
  });
  it('keeps dynamic source identity stable when a calendar is renamed', () => {
    const settings = { ...admin([]), exchangeMailboxAssignments: [rule()] };
    const first = resolveCalendarSettings({ adminSettings: settings, matchedGroupIds: matched, userSettings: defaultUserCalendarSettings, mailboxDiscoveries: [{ mailboxId: 'mailbox', sources: [source] }] });
    const user = { ...defaultUserCalendarSettings, adminSourceOverridesById: { [first.sources[0].id]: { isEnabled: false } } };
    const renamed = resolveCalendarSettings({ adminSettings: settings, matchedGroupIds: matched, userSettings: user, mailboxDiscoveries: [{ mailboxId: 'mailbox', sources: [{ ...source, name: 'Renamed' }] }] });
    expect(renamed.sources[0]).toMatchObject({ id: first.sources[0].id, isEnabled: false });
  });
  it('surfaces discovery failures without treating them as an empty successful result or deleting overrides', () => {
    const settings = { ...admin([]), exchangeMailboxAssignments: [rule()] };
    const user = { ...defaultUserCalendarSettings, adminSourceOverridesById: { 'exchange|mailbox|calendar-id': { isEnabled: false } } };
    const failed = resolveCalendarSettings({ adminSettings: settings, matchedGroupIds: matched, userSettings: user, mailboxDiscoveries: [{ mailboxId: 'mailbox', sources: [], error: 'Denied' }] });
    expect(failed.adminExchangeDiscoveryErrors).toEqual(['mailbox@example.com: Denied']);
    const saved = deriveUserCalendarSettings({ nextResolvedSettings: failed, adminSettings: settings, matchedGroupIds: matched, existingUserSettings: user });
    expect(saved.adminSourceOverridesById).toEqual(user.adminSourceOverridesById);
  });
  it('suppresses automatic current-user copies, including disabled Available assignments', () => {
    const settings = admin([{ ...assignment('a', false, false), source: { ...source, exchangeMailbox: 'current-user' } }]);
    const effective = resolveCalendarSettings({ adminSettings: settings, matchedGroupIds: matched, userSettings: defaultUserCalendarSettings, currentUserMailboxId: 'current-user' });
    expect(isAutomaticExchangeCalendarAssigned(effective, 'calendar-id')).toBe(true);
    expect(isAutomaticExchangeCalendarAssigned(effective, 'other')).toBe(false);
  });
});

describe('schema 7 migration and persistence', () => {
  it.each([2, 3, 4, 5, 6])('migrates schema %s without changing IDs, default visibility or user choices', schemaVersion => {
    const legacyAssignment = { adminSourceId: 'original-id', source: { ...source, isEnabled: false }, audienceGroups: [{ groupId: 'a', displayName: 'A' }] };
    const settings = normalizeAdminWebPartSettings({ ...defaultAdminWebPartSettings, schemaVersion, assignedSources: [legacyAssignment] });
    expect(settings?.assignedSources[0]).toMatchObject({ adminSourceId: 'original-id', isMandatory: false, defaultEnabled: false, allowedOverrides: defaultAllowedOverrides });
    expect(settings?.exchangeMailboxAssignments).toEqual([]);
    expect(settings?.audienceGroups?.[0].groupType).toBeUndefined();
    expect(normalizeUserCalendarSettings({ ...defaultUserCalendarSettings, schemaVersion, adminSourceOverridesById: { 'original-id': { isEnabled: true, color: '#ffffff' } } })?.adminSourceOverridesById['original-id']).toEqual({ isEnabled: true, color: '#ffffff' });
  });
  it('serializes shared source/group definitions once and round-trips different group policies', () => {
    const settings = admin([assignment('a', true), assignment('b', false, false)]);
    const wire = JSON.parse(serializeAdminWebPartSettings(settings));
    expect(wire.sourceCatalog).toHaveLength(1);
    expect(wire.audienceGroups).toHaveLength(2);
    expect(wire.assignedSources[0].source).toBeUndefined();
    expect(wire.assignedSources[0].audienceGroups).toBeUndefined();
    expect(wire.assignedSources[0].audienceGroupIds).toEqual(['a']);
    const loaded = normalizeAdminWebPartSettings(wire);
    expect(loaded?.assignedSources.map(item => [item.isMandatory, item.defaultEnabled])).toEqual([[true, true], [false, false]]);
  });
  it('normalizes impossible policy combinations and excludes malformed targeted audiences', () => {
    expect(normalizeAdminWebPartSettings(admin([assignment('a', true, false)]))?.assignedSources[0].defaultEnabled).toBe(true);
    const malformed = normalizeAdminWebPartSettings({ ...admin([]), assignedSources: [{ ...assignment('a'), audienceGroups: [{ groupId: '', displayName: '' }] }] });
    expect(malformed?.assignedSources).toEqual([]);
  });
  it('preserves SharePoint IDs and mappings while policies vary by group', () => {
    const sharePoint: ICalendarSourceBase = { sourceType: 'sharepoint', name: 'Events', color: '#0078d4', isEnabled: true, sharePointSiteId: 'site-id', sharePointListId: 'list-id', sharePointFieldMapping: { titleField: 'Title', startDateField: 'Start' } };
    const settings = admin([{ ...assignment('a'), source: sharePoint }, { ...assignment('b', true), source: sharePoint }]);
    const loaded = normalizeAdminWebPartSettings(JSON.parse(serializeAdminWebPartSettings(settings)));
    if (!loaded) throw new Error('Expected settings');
    const result = resolveCalendarSettings({ adminSettings: loaded, matchedGroupIds: matched, userSettings: defaultUserCalendarSettings });
    expect(result.sources).toHaveLength(1);
    expect(result.sources[0]).toMatchObject({ sharePointSiteId: 'site-id', sharePointListId: 'list-id', sharePointFieldMapping: sharePoint.sharePointFieldMapping, isMandatory: true });
    expect(getSourceIdentityKey({ ...sharePoint, name: 'Renamed' })).toBe(getSourceIdentityKey(sharePoint));
  });
});

describe('effective assignment loading', () => {
  it('requests a calendar only once when two groups assign it', async () => {
    const effective = resolveCalendarSettings({ adminSettings: admin([assignment('a'), assignment('b', true)]), matchedGroupIds: matched, userSettings: defaultUserCalendarSettings });
    const load = jest.fn().mockResolvedValue([]);
    const onLoaded = jest.fn();
    await loadExchangeSources(async () => [], effective.sources.map(item => ({ sourceId: item.id, load })), onLoaded, jest.fn());
    expect(load).toHaveBeenCalledTimes(1);
    expect(onLoaded).toHaveBeenCalledTimes(1);
  });
});

describe('shared assignment editing', () => {
  it('keeps a protected logo tied to the administrator default rather than a personal type-wide setting', () => {
    const settings = admin([{ ...assignment('a'), allowedOverrides: { ...defaultAllowedOverrides, showSourceLogo: false } }]);
    const result = resolveCalendarSettings({ adminSettings: settings, matchedGroupIds: matched, userSettings: { ...defaultUserCalendarSettings, exchangeShowSourceLogo: false } });
    expect(result.exchangeShowSourceLogo).toBe(false);
    expect(result.sources[0].showSourceLogo).toBe(true);
  });
  it('stores individual additions under an all-calendar rule as exceptions', () => {
    const settings = { ...admin([]), exchangeMailboxAssignments: [rule()] };
    const edited = addAdministratorAssignments(settings, [{ source, policy: { isMandatory: false, defaultEnabled: false } }], rule().audienceGroups);
    expect(edited.assignedSources).toEqual([]);
    expect(edited.exchangeMailboxAssignments?.[0].exceptions).toEqual([{ calendarId: 'calendar-id', isMandatory: false, defaultEnabled: false }]);
  });
  it('transfers explicit calendars into all-calendar exceptions without changing shared IDs or presentation', () => {
    const edited = addExchangeMailboxAssignment(admin([assignment('a', false, false)]), rule());
    expect(edited.assignedSources).toEqual([]);
    expect(edited.sourceCatalog?.[0].adminSourceId).toBe('original-id');
    expect(edited.exchangeMailboxAssignments?.[0].exceptions).toEqual([{ calendarId: 'calendar-id', isMandatory: false, defaultEnabled: false }]);
    const result = resolveCalendarSettings({ adminSettings: edited, matchedGroupIds: matched, userSettings: defaultUserCalendarSettings, mailboxDiscoveries: [{ mailboxId: 'mailbox', sources: [source] }] });
    expect(result.sources[0]).toMatchObject({ id: 'original-id', isEnabled: false });
  });
  it('canonicalizes old mailbox aliases while retaining administrator and user override IDs', () => {
    const settings = admin([{ ...assignment('a'), source: { ...source, exchangeMailbox: 'Legacy@contoso.com' } }]);
    const canonical = canonicalizeExchangeSourceIdentities(settings, new Map([['legacy@contoso.com', 'mailbox']]));
    expect(canonical.assignedSources[0]).toMatchObject({ adminSourceId: 'original-id', source: { exchangeMailbox: 'mailbox' } });
    const result = resolveCalendarSettings({ adminSettings: canonical, matchedGroupIds: matched, userSettings: { ...defaultUserCalendarSettings, adminSourceOverridesById: { 'original-id': { isEnabled: false } } } });
    expect(result.sources[0].isEnabled).toBe(false);
  });
  it('retains a personal current-user source suppressed by an explicit current-user administrator assignment', () => {
    const settings = admin([assignment('a')]);
    const user = { ...defaultUserCalendarSettings, personalSources: [{ ...source, exchangeMailbox: undefined, userSourceId: 'personal' }] };
    const result = resolveCalendarSettings({ adminSettings: settings, matchedGroupIds: matched, userSettings: user, currentUserMailboxId: 'mailbox' });
    expect(result.sources).toHaveLength(1);
    expect(deriveUserCalendarSettings({ nextResolvedSettings: result, adminSettings: settings, matchedGroupIds: matched, existingUserSettings: user }).personalSources.map(item => item.userSourceId)).toEqual(['personal']);
  });
});

  it('derives saves against current policy when an old personal draft predates an administrator change', () => {
    const oldAdmin = admin([assignment('a')]);
    const draft = resolveCalendarSettings({ adminSettings: oldAdmin, matchedGroupIds: matched, userSettings: defaultUserCalendarSettings });
    draft.sources[0] = applyAdminSourceChanges(draft.sources[0], { isEnabled: false, color: '#ffffff' });
    const currentAdmin = admin([{ ...assignment('a', true), allowedOverrides: { ...defaultAllowedOverrides, color: false } }]);
    const saved = deriveUserCalendarSettings({ nextResolvedSettings: draft, adminSettings: currentAdmin, matchedGroupIds: matched, mailboxDiscoveries: [] });
    expect(saved.adminSourceOverridesById).toEqual({});
  });
