# Settings and Policy

Read only the section selected by [AGENTS.md](../AGENTS.md). Requirements are normative; statements about current interfaces, storage, and resolution are verified implementation facts. Related intent and known conflicts are recorded by ID in [Decisions, deviations, and open questions](decisions-deviations-and-open-questions.md).

## Settings layers

| Layer | Current representation | Persistence | Purpose |
| --- | --- | --- | --- |
| Hardcoded defaults | `defaultAdminWebPartSettings`, `defaultUserCalendarSettings`, `defaultCalendarSettings` | Source code | Complete fallback and normalization defaults |
| Administrator settings | `IAdminWebPartSettings` | SPFx property `adminSettings` | Web-part defaults, audience assignments, source catalog, ICS catalog |
| Legacy administrator input | `ILegacyCalendarSettings` in property `settings` | SPFx property | Migration input only |
| Personal settings | `IUserCalendarSettings` | OneDrive App Folder JSON | Personal sources and minimal overrides |
| Legacy personal input | `ILegacyCalendarSettings` | OneDrive App Folder legacy JSON | Migration input only |
| Audience result | `Set<string>` of matched group IDs | Runtime plus five-minute session cache | Select applicable administrator entries |
| Effective runtime settings | `ICalendarSettings` | Not persisted as one object | Input to `MyCalendars` and personal draft editing |
| Theme-derived value | Current `palette.themePrimary` | Runtime only | Highest-precedence organization/source fallback color |
| Event/source caches | Coordinator maps, promises, event state, optional browser entry | Runtime plus optional initial-range `localStorage` cache | Avoid duplicate discovery and month loads |

`CalendarSettingsService` owns validation, normalization, migration, precedence, and persisted/effective conversion. `SettingsStorageService` owns OneDrive access only. `AudienceService` owns group discovery and membership evaluation only.

## Precedence and fields

**Read when:** changing defaults, effective settings, `resolveCalendarSettings`, personal override derivation, automatic source modes, scalar settings, or source-entry fields. Related records: DEC-009, DEC-014, DEC-015, DEC-020, DEV-003, and DEV-004.

### Effective precedence

The current implementation resolves settings in this order:

1. Hardcoded defaults fill missing administrator and personal fields during normalization.
2. Administrator settings load from current JSON, legacy administrator JSON, or defaults.
3. Audience membership filters administrator-assigned sources and ICS catalog entries.
4. Personal sources are appended after applicable administrator sources.
5. Supported personal values replace administrator scalar defaults.
6. Applicable assignments are combined by stable identity and strongest visibility policy; permitted personal administrator-source overrides then customize or remove optional sources.
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
| `schemaVersion` | Current settings contracts; `7` | No | No | Versions 2–7 normalize to 7 | Settings service | Missing, invalid, older unknown, and future versions are rejected outside explicit legacy paths |
| `defaultView` | Admin `month` | Yes | Toolbar; full override | Personal → admin → default | Web part/coordinator | Scalar, not source-lockable |
| `showWeekends` | Admin `true` | Yes | Optional personal override | Personal → admin → default | Week/range logic | Override can be removed |
| `preferredStartMinutes` | Admin `480` | Yes | Optional personal override | Renderer: personal → admin | Day/Week | Clamped and snapped to admin slot duration |
| `visibleHourCount` | Admin `10` | Yes | Optional personal override | Renderer: personal → admin | Day/Week | Normalized to 1–24 |
| `slotDurationMinutes` | Admin `30`; 15/30/60 | Yes | No | Admin → default | Day/Week | Administrator-only |
| `enableCache` | Admin `true` | Yes | No | Admin → default | Coordinator | Administrator-only browser appointment cache |
| `cacheDurationMinutes` | Admin `10`; 1–60 | Yes | No | Admin → default | Coordinator | Whole-minute stale threshold |
| `organizationPrimaryColor` | `#0078d4` | **No, deviation** | No | Theme → admin → default | Automatic colors | Confirmed administrator setting |
| `exchangeCalendarStates` | Absent IDs enabled | No | Per discovered calendar | Personal map only | Exchange | User-owned |
| Source-type `showSourceLogo` fields | Admin `true` | **No, deviation** | Yes | Personal → admin → default | Event decoration | Five source types |
| `plannerShowAllCalendars` | Admin `true` | **No, deviation** | Yes | Personal → admin → default | Planner auto mode | Suppresses explicit Planner sources |
| `plannerShowAllAssignedToMeOnly` | Admin `false` | **No, deviation** | Yes in auto mode | Personal → admin → default | Planner auto filter | Does not affect explicit sources |
| `unifiedGroupShowAllCalendars` | Admin `true` | **No, deviation** | Yes | Personal → admin → default | Group auto mode | Suppresses explicit sources |
| `teamsShiftsShowAllCalendars` | Admin `true` | **No, deviation** | Yes | Personal → admin → default | Shifts auto mode | Suppresses explicit sources |
| `assignedSources` | Admin list | Yes | Currently modifiable/removable | Audience → personal override | Resolver/coordinator | Per-assignment visibility and shared per-field override permissions are enforced |
| `icsCatalog` | Admin list | Yes | Prefills subscription flow | Audience filter only | Settings panel | Not a runtime event source |
| `personalSources` | Personal list | No | Yes | Appended after admin sources | Resolver/coordinator | User-owned |
| `adminSourceOverridesById` | Personal map | No | Derived | Stable `adminSourceId` | Resolver/save derivation | Shared administrator source ID; explicit visibility intent and permitted field overrides |
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
| SharePoint site name | SharePoint | Captured during creation; lazy legacy enrichment | Separate from calendar name; optional for schema compatibility |
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

**Read when:** changing administrator-assigned event-source membership, visibility defaults, per-field overrides, or stale-override cleanup. Related records: DEC-005, DEC-013, DEC-015 and DEC-024.

### Membership and visibility

Assignments persist two booleans, not a policy enum:

| UI label | isMandatory | defaultEnabled |
| --- | --- | --- |
| Mandatory / Verplicht | true | true |
| Default / Standaard | false | true |
| Available / Beschikbaar | false | false |

- Mandatory sources **MUST** remain enabled and **MUST NOT** be removed by users. Normalization **MUST** force their defaultEnabled to true.
- When an applicable Default or Available source becomes Mandatory, it **MUST** reappear enabled even if the user's saved override contains `removed: true` or `isEnabled: false`. Restoration **MUST NOT** require resetting or saving personal settings first. Those now-disallowed membership overrides follow the observed-change cleanup rules below; permitted presentation overrides remain applicable.
- An open personal settings draft **MUST** also restore sources when it receives the effective Mandatory policy, retaining unrelated edits and permitted source presentation edits. A pending Save snapshot **MUST NOT** be changed; the panel reconciles after that operation finishes.
- Optional sources **MAY** be disabled or removed. Their effective visibility **MUST** use an explicit user override when present, otherwise defaultEnabled.
- An explicit user visibility choice **MUST** survive administrator-default changes, even when the saved choice equals the default. Untouched sources **MUST NOT** acquire visibility overrides. Following the administrator default removes the explicit choice.
- When multiple applicable assignments identify the same calendar, the resolver **MUST** create one effective source. Mandatory takes precedence over Default, which takes precedence over Available.
- Mailbox-rule exclusions apply only to that rule; another matching assignment can still provide the calendar.

### Shared presentation and override capabilities

Source definitions share name, color, source options and allowedOverrides across all audiences. Assignments retain independent visibility policies.

The allowedOverrides booleans independently control name, color, showSourceLogo, plannerAssignedToMeOnly and showCompletedTasks where supported. Defaults are true, including during migration. SharePoint field mapping remains administrator configuration. Mandatory membership **MUST NOT** automatically lock presentation. A disallowed logo override with no configured per-source value uses the administrator type-wide logo default, so a personal type-wide toggle cannot bypass the lock.

The settings service enforces permissions during effective resolution, draft updates and personal-override derivation. UI controls reflect those same permissions. Bulk visibility actions **MUST NOT** disable mandatory sources.

### Policy changes and cleanup

- Disallowed and orphaned overrides **MUST** stop affecting effective settings immediately upon an observed change.
- Composition **MUST** clear those values from its in-memory personal settings, and the next successful personal save **MUST** persist the cleanup, including toolbar saves.
- Discarded values **MUST NOT** revive when observed policies are relaxed later.
- Failed automatic-mailbox discovery **MUST NOT** be treated as successful absence; unresolved dynamic-source overrides are retained for retry.
- No policy-revision history is stored. A mandatory interval never observed by the user cannot invalidate a previously saved override. This is the accepted boundary of the observed-change cleanup guarantee.

## Audiences

**Read when:** changing audience discovery, membership, empty targeting, paging, caching or fail-closed behavior. Related records: DEC-012 and DEC-024.

- Targets **MUST** be Microsoft 365 groups, security groups or mail-enabled security groups. Distribution and dynamic distribution groups **MUST NOT** be selectable.
- The explicit Everyone audience is represented by an empty audience reference list. Missing or invalid audience references **MUST NOT** become Everyone.
- Multiple group references use OR membership. Targeted evaluation failure **MUST** remain fail-closed.
- Audience assignment **MUST NOT** grant or imply source permissions; delegated Exchange and SharePoint access remains authoritative.
- Discovery selects groupTypes, mailEnabled and securityEnabled, retains the classified type, follows every result page, deduplicates IDs and sorts locally. Advanced group filtering uses ConsistencyLevel: eventual and count; it does not use orderby.
- Search escapes OData literals. Loading, successful empty results and failures remain distinct. Obsolete requests are discarded; selected metadata survives search errors and empty results.
- Membership uses /me/checkMemberGroups in batches of twenty, with transitive behavior and a five-minute positive/negative session cache. Failed batches write no cache result and can retry.
- Existing untyped groups remain valid after migration; successful administrator discovery enriches known group types.

AudienceService owns discovery and membership; CalendarSettingsService owns their policy effects. ICS catalog entries use the same targeting contract and remain subscription deep links.

## Storage and migration

**Read when:** changing administrator properties, current/legacy/default loading, OneDrive App Folder access, personal save/reset, schema normalization, or legacy migration. Related records: DEC-013, DEC-021, DEC-025, DEC-026, DEV-010, INT-001, OQ-003 and OQ-004.

### Ownership

- `MyCalendarsWebPart` **MUST** own accepted administrator and personal persistence callbacks and settings composition.
- `CalendarSettingsService` **MUST** own validation, normalization, migration, precedence, and persisted/effective conversion.
- `SettingsStorageService` **MUST** own personal-settings access to the OneDrive App Folder and **MUST NOT** resolve policy.
- Settings panels **MAY** edit isolated drafts but **MUST NOT** write persistence directly.

### Administrator loading and recovery

| Input | Condition | Result |
| --- | --- | --- |
| `adminSettings` | Parses and normalizes | Used |
| Legacy `settings` | Current unavailable/invalid; legacy shape parses | Migrated subset with notice |
| Hardcoded defaults | No recoverable input | Used; notice shown after invalid current data |

Administrator normalization drops malformed entries. The persisted schema separates sourceCatalog and audienceGroups from assignment records referencing their IDs. Runtime/editor assignments hydrate those references for the existing panels. Duplicate source/audience combinations and case-insensitive duplicate ICS URLs reject the payload; the same source may be assigned to different audiences. Mailbox rules persist stable mailbox identity, policy, default override capabilities and calendar-ID exceptions.

Current administrator and personal locations accept only integer schema versions 2 through 7. Versions 2–6 are normalized to version 7; version 7 is accepted directly. A missing version, non-integer version, version below 2, or future version is rejected. Unversioned input is accepted only from the explicit legacy `settings` property or legacy OneDrive filename, after those locations have been selected by the recovery flow.

### Administrator save

1. `AdminDefaultsPanel` opens an isolated complete draft from the latest accepted administrator settings. Incoming props never replace an open draft. Reset replaces only the draft with the complete defaults; Cancel persists nothing.
2. The panel snapshots the draft and awaits its `Promise<void>` callback through `PropertyPaneAdminDefaultsManager` to `MyCalendarsWebPart`.
3. The web part validates and serializes once. Its required SPFx commit callback hands the JSON to `adminSettings`. The web part verifies that the property contains that JSON before accepting the configuration; a callback that silently ignores the transfer rejects Save. SPFx is the sole property writer; there is no additional local assignment or notification path.
4. The web part accepts the configuration, then attempts audience/mailbox resolution and runtime rebuilding. Post-transfer failures are logged and exposed separately; they do not reject successful property transfer.
5. Save closes the panel only after completion. The adapter reads current settings/notices through getters and updates its summary without `propertyPane.refresh()`.

**Requirements:** administrator settings **MUST** remain a complete `IAdminWebPartSettings`. Non-exposed fields, stable identities and policy **MUST** survive edits. UI-only wizard state **MUST NOT** be persisted. The panel and composition owner **MUST** prevent concurrent Save operations. A rejected transfer **MUST** retain the draft and restore controls; a partially changed property is restored and checked through the same callback, with any restoration failure logged separately.

**SPFx caveat:** the custom field's `onRender` callback supplies the supported framework change bridge. A successful callback is an SPFx hand-off, not proof of SharePoint page publication. Save/publication/reload remains host verification. `onDispose` unmounts the adapter; async completions never update disposed controls. No timers or internal SPFx APIs are used.

The former same-save administrator mirror has been removed. Loading uses current, explicit legacy input, then defaults. Existing mirror values are not promoted. Deserialization projects only the supported webpart properties so retired data is omitted from subsequent page serialization. See DEC-025 and OQ-005.

### Personal preview, save and reset

`UserSettingsPanel` edits a clone of effective settings. Every draft change previews its consequences in the calendar through `MyCalendars`, the webpart preview callback and the existing derive/resolve policy model. Preview does not write OneDrive or administrator properties. Cancel discards preview; reopening uses the accepted configuration.

**Requirement:** optimistic client updates are preferred when the desired outcome can be shown locally. The latest confirmed personal snapshot **MUST** stay separate from preview. A storage failure **MUST** restore that snapshot, resolved against current administrator policy, and show a localized error. Permitted overrides remain personal; administrator policy structures are never written to personal storage.

Save derives personal settings, snapshots them, previews them and awaits `SettingsStorageService.saveUserSettings`. A false result rejects the callback. Success promotes the snapshot and closes the editor. Failure restores calendar state but retains the editable panel draft; editing again or retrying Save reactivates preview. Reset previews personal defaults and awaits deletion before accepting them. Legacy deletion occurs first, preserving the current file if legacy cleanup fails; current deletion completes reset. Reset failure restores calendar state and retains the original draft.

Toolbar Day/Week/Month choices also update optimistically and await storage. Failure restores the confirmed view and displays a localized message. There is at most one personal write per webpart instance; Save/Reset, view selection and opening another draft are blocked during a write. Date and search navigation remain usable. Storage I/O remains exclusively in `SettingsStorageService`.

Panels **MUST** await lifecycle-dependent writes, freeze draft editing and dismissal while saving, and close only after success. Edit-session and mount checks **MUST** reject obsolete discovery results, enrichment and save completions. Background draft changes **MUST NOT** alter a running save snapshot.

Personal read failures continue to use the existing legacy/default fallback. Distinguishing absent files from unavailable storage remains open under OQ-004. Deleting two OneDrive files is not a server transaction; no retry/transaction framework is introduced.

### Administrator assignment migration

Version 2–6 assignments become optional with defaultEnabled taken from their existing source.isEnabled. Existing adminSourceId, source IDs, audience IDs, SharePoint mappings and permitted personal values remain stable. Migration never creates an automatic mailbox rule. Mailbox aliases are resolved to object IDs during composition/editor discovery without changing source IDs; unsuccessful identity resolution retains the original identifier.

### Personal migration

Current personal settings use schema version 7. Known versions 2–6 migrate through the same normalizer and are emitted as version 7. Normalization:

- drops malformed personal sources and overrides;
- migrates legacy `userStartHour` to minutes;
- derives visible hours from legacy start/end hours;
- retains only boolean Exchange states and supported optional scalar values.
- retains a trimmed SharePoint site name when present and accepts older sources without it.

When the current file is absent, unreadable, or rejected for an unsupported schema version, the web part reads legacy `calendar-settings.json`. The rejected current file is not deleted automatically. Migration treats all legacy sources as personal, carries supported logo/automatic settings, and creates no administrator overrides. It attempts to save the migrated current file; the legacy file remains until reset.

### Persistence and cache inventory

| Mechanism/key | Data | Lifetime | Failure behavior |
| --- | --- | --- | --- |
| SPFx `adminSettings` | Current administrator JSON | Web-part persistence | Invalid value falls through to explicit legacy input or defaults |
| SPFx `settings` | Legacy combined settings | Until externally removed | Migration input only |
| OneDrive `Apps/SPFx-My-Calendar-Webpart/user-calendar-settings.json` | Personal settings | Until changed/reset | Read errors become unavailable; save returns false |
| OneDrive `Apps/SPFx-My-Calendar-Webpart/calendar-settings.json` | Legacy personal settings | Until reset/external cleanup | Migration fallback only |
| `sessionStorage.myCalendarsAudienceMembershipCache` | Group membership/expiry | Five minutes per entry | Errors logged; evaluation continues |
| `localStorage.currentUserEmailCache` | Email/timestamp | Twelve hours | Errors logged; Graph retried |
| `localStorage.currentUserMailboxSettingsCache` | Working hours/time zone | Twelve hours | Errors logged; Graph retried |
| `localStorage.myCalendars:<tenant>:<user>:<instance>` | Structurally validated canonical event segments for the initial seven-month range; new writes are versionless | Until disabled, invalid, replaced, or browser-evicted | Any read or validation failure rejects the whole entry; best-effort removal and source retrieval rebuild it |
| React/coordinator fields | Events, ranges, discovery promises | Component lifetime/reset | Not persisted |

Personal settings are the only current product data written to OneDrive. When enabled by an administrator, initial-range event segments are additionally persisted in browser `localStorage`; no event data is written to OneDrive. Mandatory-source error records are an intention with no confirmed schema, retention, or implementation.

## Settings interfaces

**Read when:** changing `UserSettingsPanel`, `AdminDefaultsPanel`, property-pane integration, settings source-creation flows, exposed controls, or toolbar preference persistence. Related records: DEC-014, DEV-001, DEV-002, and DEV-003.

### Administrator interface

The current administrator panel exposes:

- default Day, Week, or Month;
- weekend default;
- slot duration;
- preferred start time;
- visible-hour count;
- appointment-cache enablement and duration from 1–60 minutes;
- assigned Exchange, SharePoint, Planner, Unified Group/Team, and Teams Shifts sources;
- group-first Exchange/SharePoint configuration, per-calendar Mandatory/Default/Available policy, shared source presentation and individual override permissions;
- Exchange all-calendar rules including future calendars and calendar-specific exceptions; and
- audience-targeted ICS catalog entries.

The administrator Exchange flow uses a local Fluent UI people-picker wrapper for mailbox suggestions. It searches display names, UPNs and primary SMTP addresses through the existing Exchange service, keeps manual UPN/object-ID/primary-SMTP input available, and still performs authoritative identity resolution and calendar discovery only after the administrator chooses **Load calendars**.

Confirmed administrator settings missing from the UI are organization color, five source-type logo defaults, three automatic loading flags, and the Planner automatic assigned-to-me filter. Their current persisted/default values remain effective.

`PropertyPaneAdminDefaultsManager` adapts the panel to the SPFx custom property-field lifecycle. The property field obtains its own Graph client for discovery, mounts/unmounts its React subtree, forwards the accepted draft to the web part, and provides the sole SPFx property writer to the webpart and reads current values through getters.

### Personal interface

The current personal panel exposes:

- personal preferred start, visible-hour, and weekend overrides;
- current-user Exchange calendar visibility;
- source-type logo preferences;
- automatic Planner, Unified Group/Team, and Teams Shifts loading;
- Planner automatic assigned-to-me filtering;
- creation and management of personal sources;
- policy-aware modification/removal of applicable administrator sources, grouped by audience without duplicate rows;
- ICS subscription links, optionally prefilled from the administrator catalog; and
- reset of current and legacy personal-settings files.

The Outlook and SharePoint section headers expose a tri-state bulk visibility control derived entirely from their individual states. Outlook combines `exchangeCalendarStates` with configured/effective Exchange `isEnabled`; SharePoint uses configured/effective `isEnabled`. Mixed or hidden groups become fully visible when activated, while fully visible groups hide only their optional calendars. Mandatory calendars remain enabled. The controls do not store group state and do not alter Planner, Unified Group, or Teams Shifts automatic-mode fields.

SharePoint rows display site name separately from calendar name. Missing legacy names are resolved best-effort when a panel opens and are persisted only when the owning personal or administrator draft is explicitly saved.

The panel clones effective settings when opened, emits live `onPreview` changes, and awaits `onSave` or `onReset`. Cancel discards preview. It does not write storage.

Day/Week/Month selection is an optimistic personal preference with awaited persistence and rollback on failure.

### Discovery feedback

Both active panels **MUST** end mailbox discovery loading after success or failure, clear previous mailbox results/selection on new input or lookup, and ignore obsolete responses. They **MUST** show localized errors separately from successful empty results. The personal panel also shows failures of automatic current-user discovery and exposes manual mailbox input before the first successful lookup. Audience discovery feedback follows DEC-022 and preserves existing selected groups.

### Current interface gaps

- Confirmed administrator fields listed above are absent from the panel.
- Personal write failures restore confirmed runtime settings with visible feedback. Personal read failures still use the existing fallback and remain an open UX question.
- Administrator settings normalization and composition callbacks have focused automated coverage; pure policy, personal derivation, migration, dynamic mailbox expansion and bulk visibility also have coverage; focused React lifecycle tests use the existing Heft/Jest runner; SPFx host integration remains manual.
