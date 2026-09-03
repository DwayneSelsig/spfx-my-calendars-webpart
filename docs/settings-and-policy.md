# Settings and Policy

Read only the section selected by [AGENTS.md](../AGENTS.md). Requirements are normative; statements about current interfaces, storage, and resolution are verified implementation facts. Related intent and known conflicts are recorded by ID in [Decisions, deviations, and open questions](decisions-deviations-and-open-questions.md).

## Settings layers

| Layer | Current representation | Persistence | Purpose |
| --- | --- | --- | --- |
| Hardcoded defaults | `defaultAdminWebPartSettings`, `defaultUserCalendarSettings`, `defaultCalendarSettings` | Source code | Complete fallback and normalization defaults |
| Administrator settings | `IAdminWebPartSettings` | SPFx properties `adminSettings` and `adminSettingsBackup` | Web-part defaults, audience assignments, source catalog, ICS catalog |
| Legacy administrator input | `ILegacyCalendarSettings` in property `settings` | SPFx property | Migration input only |
| Personal settings | `IUserCalendarSettings` | OneDrive App Folder JSON | Personal sources and minimal overrides |
| Legacy personal input | `ILegacyCalendarSettings` | OneDrive App Folder legacy JSON | Migration input only |
| Audience result | `Set<string>` of matched group IDs | Runtime plus five-minute session cache | Select applicable administrator entries |
| Effective runtime settings | `ICalendarSettings` | Not persisted as one object | Input to `MyCalendars` and personal draft editing |
| Theme-derived value | Current `palette.themePrimary` | Runtime only | Highest-precedence organization/source fallback color |
| Event/source caches | Coordinator maps, promises, event state | Runtime only | Avoid duplicate discovery and month loads |

`CalendarSettingsService` owns validation, normalization, migration, precedence, and persisted/effective conversion. `SettingsStorageService` owns OneDrive access only. `AudienceService` owns group discovery and membership evaluation only.

## Precedence and fields

**Read when:** changing defaults, effective settings, `resolveCalendarSettings`, personal override derivation, automatic source modes, scalar settings, or source-entry fields. Related records: DEC-009, DEC-014, DEC-015, DEV-003, and DEV-004.

### Effective precedence

The current implementation resolves settings in this order:

1. Hardcoded defaults fill missing administrator and personal fields during normalization.
2. Administrator settings load from current JSON, backup JSON, legacy administrator JSON, or defaults.
3. Audience membership filters administrator-assigned sources and ICS catalog entries.
4. Personal sources are appended after applicable administrator sources.
5. Supported personal values replace administrator scalar defaults.
6. Personal administrator-source overrides replace name, color, and enabled state or remove the source.
7. The current theme primary color replaces the administrator organization color at runtime when present.

Precedence is not uniform for every field:

- `preferredStartMinutes` and `visibleHourCount` remain administrator base values; optional personal counterparts are consumed by the renderer.
- `showWeekends` is resolved directly, while administrator and optional personal values remain available for reset-to-admin behavior.
- The toolbar always writes a personal `defaultView`; the settings panel has no separate default-view control.
- `slotDurationMinutes` has no personal override.
- `organizationPrimaryColor` resolves theme, then administrator, then hardcoded default.
- Automatic Exchange membership has no administrator on/off switch; personal states can hide individual discovered calendars.
- Automatic Planner, Unified Group, or Teams Shifts mode suppresses explicit configured sources of that type.

### Central scalar matrix

`Admin UI` and `User UI` describe the current interface. `Current override` is current code, not confirmed future policy.

| Setting | Scope/default | Admin UI | User UI/current override | Effective precedence | Used by | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| `schemaVersion` | All contracts; `4` | No | No | Normalizers emit current version | Settings service | Migration metadata |
| `defaultView` | Admin `month` | Yes | Toolbar; full override | Personal → admin → default | Web part/coordinator | Scalar, not source-lockable |
| `showWeekends` | Admin `true` | Yes | Optional personal override | Personal → admin → default | Week/range logic | Override can be removed |
| `preferredStartMinutes` | Admin `480` | Yes | Optional personal override | Renderer: personal → admin | Day/Week | Clamped and snapped to admin slot duration |
| `visibleHourCount` | Admin `10` | Yes | Optional personal override | Renderer: personal → admin | Day/Week | Normalized to 1–24 |
| `slotDurationMinutes` | Admin `30`; 15/30/60 | Yes | No | Admin → default | Day/Week | Administrator-only |
| `organizationPrimaryColor` | `#0078d4` | **No, deviation** | No | Theme → admin → default | Automatic colors | Confirmed administrator setting |
| `exchangeCalendarStates` | Absent IDs enabled | No | Per discovered calendar | Personal map only | Exchange | User-owned |
| Source-type `showSourceLogo` fields | Admin `true` | **No, deviation** | Yes | Personal → admin → default | Event decoration | Five source types |
| `plannerShowAllCalendars` | Admin `true` | **No, deviation** | Yes | Personal → admin → default | Planner auto mode | Suppresses explicit Planner sources |
| `plannerShowAllAssignedToMeOnly` | Admin `false` | **No, deviation** | Yes in auto mode | Personal → admin → default | Planner auto filter | Does not affect explicit sources |
| `unifiedGroupShowAllCalendars` | Admin `true` | **No, deviation** | Yes | Personal → admin → default | Group auto mode | Suppresses explicit sources |
| `teamsShiftsShowAllCalendars` | Admin `true` | **No, deviation** | Yes | Personal → admin → default | Shifts auto mode | Suppresses explicit sources |
| `assignedSources` | Admin list | Yes | Currently modifiable/removable | Audience → personal override | Resolver/coordinator | Confirmed policy is unenforced |
| `icsCatalog` | Admin list | Yes | Prefills subscription flow | Audience filter only | Settings panel | Not a runtime event source |
| `personalSources` | Personal list | No | Yes | Appended after admin sources | Resolver/coordinator | User-owned |
| `adminSourceOverridesById` | Personal map | No | Derived | Stable `adminSourceId` | Resolver/save derivation | Current shape lacks policy |
| `sources` | Effective list | Indirect | Personal management | Admin, then personal | Coordinator | Not persisted directly |
| `firstDayOfWeek` | Legacy only | No | No | Ignored in migration | None | Renderers use fixed rules |

Source-type logo fields are `exchangeShowSourceLogo`, `sharePointShowSourceLogo`, `plannerShowSourceLogo`, `unifiedGroupShowSourceLogo`, and `teamsShiftsShowSourceLogo`.

### Source-entry matrix

| Field | Applies to | Configuration/current override | Intended policy/runtime notes |
| --- | --- | --- | --- |
| `name`, `color` | All entries | Admin/user; currently overrideable for admin entries | Independently allow/deny |
| `isEnabled`, removal | All entries | Admin/user; currently overrideable | Optional allowed; mandatory denied by membership |
| `sourceType` | All entries | Selected at creation; not overrideable | Six model values; `ics` has no adapter |
| Exchange mailbox/calendar ID | Exchange | Creation only | Mailbox and calendar endpoint identity |
| SharePoint site/list IDs | SharePoint | Creation only | Required for loading |
| SharePoint field mapping | SharePoint | Creation flows; all-day selector not exposed | Required/optional column mapping; future option capability |
| Planner plan ID/title | Planner | Creation only | Plan ID required; title not used for retrieval |
| Planner assignment/completion filters | Planner | Creation; no current admin-source user override | Independently allow/deny in policy |
| `groupId` | Unified Group | Creation only | Group/Team calendar identity |
| `showSourceLogo` | Source entry | Collected by some flows | Only manual Exchange currently honors it; policy capability |
| `icsUrl` | Legacy ICS shape | Subscription UI/catalog | No event adapter; outside source-policy scope |
| `audienceGroups` | Admin source/catalog | Admin only; normalizer currently requires at least one | Separate audience contract |
| `adminSourceId` | Admin source | Generated; key for overrides | Stable policy/migration identity |
| `adminIcsId`, ICS `displayName` | Admin catalog | Admin only | Catalog deep-link identity/label |
| `userSourceId` | Personal source | Generated/migrated | User-owned runtime identity |
| `origin` | Effective source | Derived `admin`/`user` | Routes save derivation and badges |
| `audienceGroupNames` | Effective admin source | Derived | Displays assignment provenance |

## Administrator source policy

**Read when:** changing administrator-assigned event-source membership, disable/remove behavior, user override permissions, source locking, or stale override cleanup. Related records: DEC-005, DEC-013, DEC-015, DEV-001, DEV-004, and OQ-001.

This section defines confirmed product requirements. The policy is not implemented. It applies only to administrator-assigned event-source entries, not to scalar administrator defaults, automatic-source flags, personal sources, or ICS catalog entries.

### Two independent dimensions

1. Membership is `optional` or `mandatory`.
2. Allowed presentation and source-option overrides are controlled separately for each supported field.

Conceptual outcomes are combinations, not necessarily serialized enum values:

| Outcome | Membership | Allowed overrides | Meaning |
| --- | --- | --- | --- |
| Fully overridable | Optional | All supported presentation/source-option overrides | Administrator pushes a default that the user can disable, remove, and customize |
| Partially restricted | Optional or mandatory | Administrator-selected subset | Membership follows its own rule; some presentation or source options are protected |
| Mandatory/locked | Mandatory | No personal source overrides | User cannot disable, remove, rename, recolor, or change protected source options |

### Membership requirements

- An optional source **MUST** be removable and disableable by the user.
- A mandatory source **MUST NOT** be removed or disabled by the user.
- Mandatory membership does not by itself lock name, color, logo, or source-specific filters.
- A mandatory source **MAY** remain customizable when the applicable overrides are allowed.

### Override capabilities

The policy **MUST** be able to control these independently where the source supports them:

- name;
- color;
- source-logo visibility; and
- source-specific filters or options, including Planner assignment/completion filters.

Disable and remove behavior comes from membership and is not duplicated in the override capability set.

The exact persisted schema, field names, default capability set, and migration representation are unresolved. They **MUST NOT** be invented during an unrelated implementation.

### Policy changes and stale overrides

- A newly disallowed override **MUST** stop affecting effective settings immediately.
- An override for a removed, inaccessible, or no-longer-applicable administrator source is orphaned and **MUST** stop affecting runtime settings.
- Disallowed and orphaned overrides **MUST** be removed on the next successful personal-settings save.
- A policy change **MUST NOT** restore a previously disallowed stale value later.

### Current deviation

`IAdminAssignedSource` has no membership or allowed-override policy. Every applicable administrator source can currently be renamed, recolored, disabled, or removed. Current code ignores overrides whose administrator source is not resolved, but can retain them on disk until a later user save.

## Audiences

**Read when:** changing administrator targeting, group discovery, group membership evaluation, empty audiences, audience caching, or fail-closed behavior. Related records: DEC-012 and DEV-002.

### Targeting contract

- Audience targets **MUST** be groups, not individual users.
- Supported targets **MUST** include security groups, mail-enabled security groups, and Microsoft 365 groups.
- An assignment with no audience targets **MUST** apply to everyone.
- Multiple groups **MUST** use OR semantics: membership in any selected group grants the assignment.
- Group-targeted evaluation failure **MUST** be fail-closed.
- Audience targeting applies to administrator-assigned event sources and administrator ICS catalog entries.
- `AudienceService` **MUST** own discovery and membership evaluation but **MUST NOT** decide the effect of source policy.

### Current implementation

- `AudienceService.getSecurityGroups` returns only `mailEnabled eq false and securityEnabled eq true` groups, at most 50 results.
- The UI cannot create or retain an assignment with no groups; normalization drops such entries.
- `/me/checkMemberGroups` evaluates up to 20 IDs per batch and supports transitive membership.
- Positive and negative results use session storage for five minutes.
- A failed batch contributes no matches and writes no result for those IDs, so the next evaluation can retry.
- The cache key contains the group ID but not tenant or user identity.

The first two items conflict with the confirmed target model. The cache-key scope is a current implementation risk and is not a confirmed policy change.

## Storage and migration

**Read when:** changing administrator properties, current/backup recovery, OneDrive App Folder access, personal save/reset, schema normalization, or legacy migration. Related records: DEC-013, DEV-010, DEV-011, INT-001, OQ-003, OQ-004, and OQ-005.

### Ownership

- `MyCalendarsWebPart` **MUST** own accepted administrator and personal persistence callbacks and settings composition.
- `CalendarSettingsService` **MUST** own validation, normalization, migration, precedence, and persisted/effective conversion.
- `SettingsStorageService` **MUST** own personal-settings access to the OneDrive App Folder and **MUST NOT** resolve policy.
- Settings panels **MAY** edit isolated drafts but **MUST NOT** write persistence directly.

### Administrator loading and recovery

| Input | Condition | Result |
| --- | --- | --- |
| `adminSettings` | Parses and normalizes | Used |
| `adminSettingsBackup` | Current is invalid; backup parses and normalizes | Used with warning notice |
| Legacy `settings` | Current forms unavailable/invalid; legacy shape parses | Migrated subset with notice |
| Hardcoded defaults | No recoverable input | Used; notice shown after invalid current/backup data |

Administrator normalization drops malformed list entries. It rejects the complete payload when two assigned sources have the same source identity or two ICS catalog items have the same case-insensitive URL.

### Administrator save

1. `AdminSettingsPanel` edits a deep-cloned draft.
2. Saving calls `MyCalendarsWebPart.handleAdminSettingsSave`.
3. The web part serializes the same accepted value to `adminSettings` and `adminSettingsBackup`.
4. It reevaluates audiences, rebuilds effective settings, refreshes the property pane, and renders React.

**Fact:** backup is a same-save mirror, not a rotated previous revision. The save callback does not re-run normalization before assigning the in-memory draft.

### Personal save and reset

1. `SettingsPanel` edits a deep clone of effective settings.
2. `deriveUserCalendarSettings` produces minimal personal values relative to applicable administrator settings.
3. The web part resolves settings immediately and asks `SettingsStorageService` to save asynchronously.
4. The UI renders the in-memory result before persistence success is known.
5. Save failure is logged only; there is no user-visible error or rollback.

Reset deletes current and legacy personal files. In-memory reset occurs only after both deletions report success. A failure leaves current in-memory settings unchanged and is logged.

### Personal migration

Current personal settings use schema version 4. Normalization:

- drops malformed personal sources and overrides;
- migrates legacy `userStartHour` to minutes;
- derives visible hours from legacy start/end hours;
- retains only boolean Exchange states and supported optional scalar values.

When the current file is absent or unreadable, the web part reads legacy `calendar-settings.json`. Migration treats all legacy sources as personal, carries supported logo/automatic settings, and creates no administrator overrides. It attempts to save the migrated current file; the legacy file remains until reset.

### Persistence and cache inventory

| Mechanism/key | Data | Lifetime | Failure behavior |
| --- | --- | --- | --- |
| SPFx `adminSettings` | Current administrator JSON | Web-part persistence | Invalid value falls through to backup |
| SPFx `adminSettingsBackup` | Same-save mirror | Web-part persistence | Used when current is invalid |
| SPFx `settings` | Legacy combined settings | Until externally removed | Migration input only |
| OneDrive `Apps/SPFx-My-Calendar-Webpart/user-calendar-settings.json` | Personal settings | Until changed/reset | Read errors become unavailable; save returns false |
| OneDrive `Apps/SPFx-My-Calendar-Webpart/calendar-settings.json` | Legacy personal settings | Until reset/external cleanup | Migration fallback only |
| `sessionStorage.myCalendarsAudienceMembershipCache` | Group membership/expiry | Five minutes per entry | Errors logged; evaluation continues |
| `localStorage.currentUserEmailCache` | Email/timestamp | Twelve hours | Errors logged; Graph retried |
| `localStorage.currentUserMailboxSettingsCache` | Working hours/time zone | Twelve hours | Errors logged; Graph retried |
| React/coordinator fields | Events, ranges, discovery promises | Component lifetime/reset | Not persisted |

Personal settings are the only current product data written to OneDrive. Loaded events are not persisted. Mandatory-source error records are an intention with no confirmed schema, retention, or implementation.

## Settings interfaces

**Read when:** changing `SettingsPanel`, `AdminSettingsPanel`, property-pane integration, settings source-creation flows, exposed controls, or toolbar preference persistence. Related records: DEC-014, DEV-001, DEV-002, and DEV-003.

### Administrator interface

The current administrator panel exposes:

- default Day, Week, or Month;
- weekend default;
- slot duration;
- preferred start time;
- visible-hour count;
- assigned Exchange, SharePoint, Planner, Unified Group/Team, and Teams Shifts sources;
- source name, color, enabled state, source-specific creation options, and audiences; and
- audience-targeted ICS catalog entries.

Confirmed administrator settings missing from the UI are organization color, five source-type logo defaults, three automatic loading flags, and the Planner automatic assigned-to-me filter. Their current persisted/default values remain effective.

`PropertyPaneAdminCalendarManager` adapts the panel to the SPFx custom property-field lifecycle. The property field obtains its own Graph client for discovery, mounts/unmounts its React subtree, and forwards the accepted draft to the web part.

### Personal interface

The current personal panel exposes:

- personal preferred start, visible-hour, and weekend overrides;
- current-user Exchange calendar visibility;
- source-type logo preferences;
- automatic Planner, Unified Group/Team, and Teams Shifts loading;
- Planner automatic assigned-to-me filtering;
- creation and management of personal sources;
- current modification/removal of applicable administrator sources;
- ICS subscription links, optionally prefilled from the administrator catalog; and
- reset of current and legacy personal-settings files.

The panel deep-clones effective settings when opened. Unsaved changes are discarded on close. It emits `onSave` or `onReset`; it does not write storage.

Day/Week/Month selection in the toolbar is also a personal setting and is persisted immediately.

### Current interface gaps

- Optional/mandatory membership and allowed-override policy have no schema or controls.
- Every applicable administrator source can currently be removed, disabled, renamed, and recolored.
- Empty-audience assignments are impossible instead of applying to everyone.
- Audience discovery excludes mail-enabled security groups and Microsoft 365 groups.
- Confirmed administrator fields listed above are absent from the panel.
- Per-source logo behavior is implemented only for manual Exchange sources.
- Personal storage failure is not visible and does not roll back in-memory changes.
- No settings or policy behavior has automated test coverage.
