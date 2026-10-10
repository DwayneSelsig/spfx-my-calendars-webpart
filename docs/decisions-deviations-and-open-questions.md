# Decisions, Deviations, and Open Questions

Search this register by the IDs referenced in the applicable architecture, calendar, source, or settings section. Do not read every record for every task.

## Status semantics

| Label | Meaning |
| --- | --- |
| Decision | Confirmed product or architecture choice. |
| Intention | Confirmed desired behavior that can be incomplete. |
| Deviation | Known difference between current behavior and confirmed intent. |
| Technical debt | Current code or structure that is not desired architecture. |
| Assumption | Working interpretation that has not been confirmed. |
| Open question | Decision that has not been made. |

Only decisions are confirmed choices. An intention, deviation, technical-debt item, assumption, or open question **MUST NOT** be represented as confirmed architecture.

## Product and rendering

### DEC-001 — Read-only aggregation

- **Status:** Decision
- My Calendars reads and combines source data. It does not create, change, or delete source events or tasks.
- It can open a source event, calendar, task, or subscription flow through a deep link when supported.
- Event Details exposes Open only for an exact source-item link. It does not substitute a general source or application destination when an exact item link is unavailable.
- Write operations do not belong in source services or renderers.

### DEC-003 — Local calendar renderer

- **Status:** Decision
- A local read-only renderer provides Day, Week, and Month.
- Month preserves a seven-column, six-row calendar. Week uses rolling day cards and a complete 24-hour timeline. Day uses the Microsoft 365 Companion interaction and visual direction.
- The custom Search view remains active and keeps the hidden calendar view mounted.
- Schedule is not supported.

### DEC-007 — Dynamic retrieval range

- **Status:** Decision
- Initial loading covers the current month plus or minus three months.
- Navigation loads missing visible months per source.
- Only successful source/month responses enter the range cache.
- Manual refresh clears events and range state.

### DEC-009 — Display preferences and regional formatting

- **Status:** Decision
- An administrator sets the default weekend visibility.
- A user can store an explicit personal weekend preference or return to the administrator default.
- Date and time labels use the current SharePoint page culture; there is no separate 12/24-hour setting.

### DEC-011 — Supported hosts

- **Status:** Decision
- All declared or runtime-recognized hosts are supported: SharePoint Web Part, SharePoint full-page, Teams personal app, Teams tab, Office, and Outlook.
- A change to host-sensitive behavior must verify the relevant host rather than infer parity from the manifest.

### DEC-017 — Administrator-controlled appointment cache

- **Status:** Decision
- Browser appointment caching is an administrator-only setting, enabled by default with a ten-minute duration normalized to 1–60 whole minutes.
- New `localStorage` entries are versionless and isolated by tenant, user, web-part instance, and result-affecting configuration signature. A schema-version field in an existing entry is ignored. Entries contain canonical events per source/month for only the initial seven-month range.
- Compatible cached events render immediately and are reprocessed locally. Fresh segments suppress retrieval; stale segments remain visible while retrieval runs in the background.
- Cache expiry by itself does not schedule retrieval while the component remains mounted. A later load that finds stale segments uses the existing loading and toolbar-status workflow.
- Successful source/month retrieval atomically replaces that segment, including with an empty result. Failure preserves stale data and the source failure remains visible.
- Manual refresh preserves visible appointments, bypasses freshness, and forces the initial and visible ranges. Any read, parse, structural, identity, or signature failure rejects the entire entry; removal is best-effort, no partial data is hydrated, and normal source retrieval rebuilds the cache. Invalid, disabled, unavailable, read-only, or quota-constrained cache state falls back safely to normal retrieval.

### DEC-018 — Event detail semantics

- **Status:** Decision
- Organizer metadata and `isOrganizer` exist only for meetings with at least one normalized attendee; an attendee-less appointment has no organizer semantics.
- SharePoint all-day values use local calendar boundaries, converting the source's inclusive end date to an exclusive end boundary.
- Description formatting is explicit. An absent format means escaped plain text, while declared HTML is sanitized before rendering and retained links open in a protected new tab.

### DEC-019 — Source context and Graph event status

- **Status:** Decision
- SharePoint calendar name and site name are separate fields. Legacy site names are resolved lazily, and renderers consume local event metadata rather than calling Graph.
- `showAs` is normalized for Exchange and Microsoft 365 Group events and is presented uniformly as calendar availability, not necessarily the signed-in user's availability.
- `responseStatus` is personal UI only for events retrieved through `/me`; mailbox and group responses are ignored.
- Calendar color remains a solid identity edge. Availability uses an additional icon and secondary surface treatment; declined personal events remain visible, muted, and struck through.

### DEC-020 — Settings group visibility controls

- **Status:** Decision
- Outlook and SharePoint have tri-state bulk visibility controls derived from individual calendar states; no separate group visibility state is stored.
- A fully visible group becomes hidden. A hidden or mixed group becomes fully visible.
- Outlook bulk visibility updates current-user `exchangeCalendarStates` and configured/effective Exchange `isEnabled`; SharePoint updates configured/effective `isEnabled`.
- Planner, Microsoft 365 Group/Team, and Teams Shifts discovery-mode settings are not visibility controls and receive no group visibility control in this scope.

### DEC-023 — Responsive search and action bar

- **Status:** Decision
- A flex layout replaces the top CommandBar; SearchBox and Fluent UI CommandBarButton actions remain.
- Layout follows the web-part content width, observed with ResizeObserver: labels at 600 px and above, icons only from 360 to below 600 px, and right-aligned actions above full-width search below 360 px.
- Search stays directly visible and preserves input and focus during resizing. Result filtering retains its existing debounce and calendar-view restoration.
- Icon-only Settings and Refresh retain localized accessible names and tooltips. The existing loading-status action and callout remain available.

### DEBT-001 — Inactive Schedule view

- **Status:** Resolved
- The inactive `ScheduleView.tsx` and its commented command-bar branch were removed.
- Schedule remains unsupported under DEC-003.

## Sources and resilience

### DEC-002 — ICS subscription flow

- **Status:** Decision
- ICS opens the Exchange Online subscription wizard through a generated deep link.
- The web part does not fetch, parse, normalize, cache, or render ICS content.
- The administrator ICS catalog publishes approved subscription links to audiences; it is not an event-source catalog.

### DEC-004 — Independent source loading

- **Status:** Decision
- A source failure does not cancel successful independent source loads.
- Valid partial results remain visible.
- Source-service failures have visible per-service status.

### DEC-008 — Source errors and invalid dates

- **Status:** Decision
- Active source failures remain visible to the coordinator and do not receive a successful range-cache entry.
- A source adapter rejects an event with unusable required dates. It does not invent the current time.

### DEC-015 — Source logo precedence

- **Status:** Decision
- A source type provides the default source-logo visibility.
- An individual configured source can override that type default.
- Administrator source policy can independently permit or deny the per-source override.

### DEC-016 — Teams Shifts scope

- **Status:** Decision
- Teams Shifts includes all shifts returned for all Teams where the current user is a direct member.
- It is not limited to shifts assigned to the current user.

### DEC-022 — Mailbox identity and discovery feedback

- **Status:** Decision
- Configured Exchange mailboxes accept UPN, object ID and primary SMTP (`mail`). Secondary aliases remain outside scope; delegated `User.ReadBasic.All` supports the needed basic-field identity resolution.
- Identity resolution is separate from calendar access. Zero/multiple identity matches and request failures are distinct; discovery/event 404 does not prove a missing mailbox.
- Active settings discovery completes loading on success and failure, clears stale results, ignores obsolete requests and shows localized error/empty feedback. Audience selections survive discovery failures and changes to search results.
- Audience discovery follows all result pages and sorts locally without Graph `$orderby`. Membership semantics remain unchanged; DEC-024 expands discovery to the confirmed group types and Everyone.
- LimitedDetails support has no unverified fallback; existing event fields/endpoints remain in use and tenant validation is required.

### INT-001 — Mandatory-source failure records

- **Status:** Intention
- Store a dated record in the OneDrive App Folder when a Graph error prevents a mandatory source from loading.
- Retry the mandatory source after a delay.
- Do not store source event content unless a later decision explicitly permits it.
- Record schema, delay, retry limit, retention, cleanup, and UI behavior remain open.

### DEV-004 — Per-source logo precedence is inconsistent

- **Status:** Resolved
- Configured Exchange, SharePoint, Planner, Unified Group and Teams Shifts sources use their per-source logo preference before the source-type default. Automatic discovery keeps its existing type-wide default.

### DEV-005 — `MailboxSettings.Read` has no current consumer

- **Current state:** both permission manifests request `MailboxSettings.Read`.
- **Current state:** no `/me/mailboxSettings` endpoint is called; `UserHelper` uses `/me/calendar/getSchedule`, covered by `Calendars.Read`.
- **Required follow-up:** remove or justify the permission in a separately scoped runtime/configuration change.

### DEV-006 — Invalid Planner dates can fail a plan load

- **Current state:** Planner skips a task with no dates, but a malformed non-empty date can reach `toISOString()` and throw.
- **Desired state:** reject only the invalid item under DEC-008.

### DEV-007 — Exchange mailbox validation uses the generic HTTP client

- **Status:** Resolved
- PR #5 switched the original user lookup to authenticated Graph access.
- DEC-022 replaces boolean existence validation with identity resolution for UPN, object ID and primary SMTP, followed by authoritative calendar discovery. Identity lookup, discovery and event errors remain distinguishable.
- There is no generic HTTP fallback.

### DEV-008 — Some “all” source modes do not page all data

- **Current state:** Planner group discovery uses `$top=100` without following `@odata.nextLink`.
- **Current state:** Exchange and Group events use `$top=500` without paging.
- **Current state:** SharePoint runtime items and several discovery calls do not follow paging.
- **Consequence:** “all” is limited to the implemented returned pages, not a guaranteed complete tenant data set.

### DEV-009 — Teams Shifts failure isolation is coarser than Team scope

- **Current state:** joined Teams are queried sequentially; a non-404 error aborts the call and drops earlier collected Team results.
- **Required contract:** source-family isolation remains required; Team-level isolation inside Shifts is not confirmed.

### DEV-012 — Source registry and inactive UI contain obsolete ICS wording

- **Status:** Resolved
- The inactive dialog was removed and active ICS descriptions now describe an Outlook subscription.
- The active UI opens an Outlook subscription link and does not parse content.

### DEBT-002 — Inactive Add Calendar dialog

- **Status:** Resolved
- The inactive `AddCalendarDialog.tsx` was removed; the integrated settings-panel flow remains authoritative.

### DEBT-003 — Unused service dependencies

- **Status:** Resolved
- Unused `HttpClient` dependencies were removed from Planner, SharePoint, Teams Shifts, and Unified Group services.
- The unused Exchange `HttpClient` dependency and Graph base-URL constant were also removed; all Exchange calls use the authenticated Graph client.

### OQ-002 — Mandatory-source retry

- **Status:** Open question
- What delay, retry limit, and triggering event apply?
- How does the UI show a persistent mandatory-source failure?
- Does Team-level failure inside one Shifts source count as failure of the mandatory source?

### OQ-003 — OneDrive mandatory-source error records

- **Status:** Open question
- What file and record schema is used?
- Which diagnostic fields are permitted?
- What retention and cleanup rules apply?
- How is sensitive source/event content excluded?

## Settings and policy

### DEC-005 — Administrator source policy dimensions

- **Status:** Decision
- Policy applies to administrator-assigned event sources, not scalar defaults or ICS subscriptions.
- Persist isMandatory and defaultEnabled separately; mandatory normalizes to enabled. UI labels are Mandatory, Default and Available.
- Optional sources can be disabled or removed; mandatory sources cannot.
- Source definitions share presentation/options and individual allowedOverrides booleans across audiences. Membership does not imply presentation locking.
- Allowed keys are name, color, showSourceLogo, plannerAssignedToMeOnly and showCompletedTasks; applicable defaults are true.
- Schema 7 sourceCatalog/audienceGroups store shared metadata; persisted assignments reference their IDs. Hydrated editor/runtime records preserve existing component contracts.

### DEC-012 — Audience target model

- **Status:** Decision
- Administrator audience targets are groups, not individual users.
- Security groups, mail-enabled security groups, and Microsoft 365 groups are supported.
- An empty audience means everyone.
- Multiple groups use OR semantics.
- Group-targeted evaluation is fail-closed.

### DEC-013 — Stale personal overrides

- **Status:** Decision
- A disallowed or orphaned override stops affecting effective settings immediately.
- It is removed from personal storage on the next successful personal-settings save.
- A policy change does not restore a previously disallowed stale value later.
- An observed Default/Available to Mandatory transition restores a previously removed or disabled applicable source immediately, including when loading an older personal settings file. It requires no personal save/reset; the next successful personal save persists the removal/visibility cleanup while retaining permitted presentation overrides.

### DEC-014 — Administrator configuration fields

- **Status:** Decision
- Organization color, source-type logo defaults, automatic Planner/Group/Shifts loading defaults, and the Planner automatic assigned-to-me default are administrator-configurable product settings.
- Presence only in model/default values does not satisfy the required administrator UI.

### DEC-021 — Persisted settings version compatibility

- **Status:** Decision
- Integer administrator/personal schema versions 2–7 normalize to version 7. Missing, non-integer, older unknown and future versions are rejected in current locations.
- Versions 2–6 preserve administrator source IDs, user source IDs and permitted personal choices. Existing optional source visibility becomes defaultEnabled. Migration creates no automatic mailbox rules.
- Unversioned data is accepted only through explicit legacy locations. Existing recovery order and same-save backup semantics remain.
- Persistence-failure UX and historical backup rotation remain unresolved under OQ-004 and OQ-005.

### DEC-024 — Grouped assignments and automatic Exchange mailbox rules

- **Status:** Decision
- Administrator means the person authorized to edit web-part properties.
- Exchange/SharePoint configuration starts with a supported Entra group or explicit Everyone, then a mailbox/site, then selected calendars and policy.
- Exchange supports an opt-in all-calendar rule including future calendars, discovered on initialization, accepted administrator changes and manual refresh. No polling is introduced.
- Rule exceptions inherit policy, replace it or exclude a calendar only from that rule. SharePoint remains explicit selection with per-list field mapping.
- This opt-in rule replaces the earlier blanket requirement that new mailbox calendars never become assigned; explicit selection retains that original behavior.
- Stable mailbox/calendar and site/list identities deduplicate matching assignments; Mandatory > Default > Available. Each effective calendar loads once.
- Name, color, options and override capabilities are shared per calendar. Group policies are independent.
- Personal settings show one row; sources matching several audiences appear under Multiple groups with their provenance. Explicit visibility choices remain until reset or an observed policy disallows them.
- Cleanup follows observed changes in memory and the next successful personal save; no revision history covers missed policy intervals. Discovery failure is not confirmed absence.
- Audience selection never grants or implies Exchange/SharePoint authorization. Mandatory failures remain visible.

### DEC-025 — Administrator property change hand-off

- **Status:** Decision
- Normalize and serialize a draft, notify SPFx of current and backup property changes, then synchronize the local property bag and accept runtime settings.
- Do not prewrite both properties before their callbacks: a freshly reset SPFx host snapshot can treat the already-complete state as its baseline and fail to deliver a changed state to the page.
- Current and backup remain identical same-save mirrors. Schema, assignment/source IDs, recovery order and personal choices are unchanged.
- Obsolete save completions do not refresh the administrator property pane over a newer accepted save.

### DEV-014 — Administrator edits disappear after page reload

- **Status:** Repair implemented; tenant verification pending
- **Reported behavior:** an existing Exchange assignment changes from Default to Mandatory in the current session but reverts after republishing and reloading.
- **Local evidence:** the policy JSON round-trip succeeds. Prewriting both properties reproduces lost host delivery in a focused model of the installed SPFx snapshot-reset behavior; notification-first delivery passes that regression.
- **Removal scenario evidence:** `MandatoryCalendarRestoration.test.ts` exercises Default assignment, personal removal, saved Mandatory promotion, old personal-file reload and subsequent cleanup. The current local policy implementation passes for explicit Exchange/SharePoint assignments, mailbox rules and supported older personal schemas. This verifies settings resolution and JSON round-trips, not the deployed bundle or published tenant page.
- **Remaining verification:** confirm the active deployed bundle, compare accepted/callback/published assignment values and repeat the page publication/reload flow in SharePoint. The reported tenant failure's exact cause has not been observed directly.

### DEV-001 — Administrator source policy is not implemented

- **Status:** Resolved
- Schema 7 implements visibility policy, shared field-level override permissions, personal controls and observed-change stale-override cleanup under DEC-005, DEC-013 and DEC-024.

### DEV-002 — Audience support is narrower than the confirmed model

- **Status:** Resolved
- Discovery supports Microsoft 365, security and mail-enabled security groups. Explicit Everyone applies without membership.
- OR matching, paging, stale-request protection and fail-closed evaluation are retained. Invalid references do not become Everyone.

### DEV-003 — Administrator settings UI is incomplete

- **Current state:** the model contains DEC-014 fields, but the panel does not expose them.
- **Current effect:** hardcoded, migrated, or manually serialized values can apply and users can override supported fields.
- **Desired state:** provide controls without silently changing confirmed precedence.

### DEV-010 — Personal persistence failures are not visible

- **Current state:** in-memory personal changes apply before OneDrive save completes; failure is logged without rollback or user feedback.
- **Current state:** a read failure is treated like a missing file and can trigger legacy/default fallback.
- **Desired failure UX:** open question.

### DEV-011 — Administrator backup is not historical

- **Current state:** a save writes identical JSON to current and backup properties.
- **Consequence:** backup can recover isolated current-property corruption but is not the previous known-good revision.
- **Desired semantics:** open question.

### OQ-001 — Persisted administrator source-policy schema

- **Status:** Resolved by DEC-005, DEC-021 and DEC-024
- Schema 7 separates shared source/audience definitions and assignment policies. Existing assignments retain their visibility default and permit supported presentation overrides.
- Supported capability keys are explicit normalized booleans; future keys require a confirmed contract/schema change.

### OQ-004 — Personal persistence failure UX

- **Status:** Open question
- Should failed saves roll back, remain pending, or show retry state?
- How does the product distinguish a missing file from unavailable storage?

### OQ-005 — Administrator backup semantics

- **Status:** Open question
- Should backup remain a same-save mirror or rotate the previous validated revision?
- When is a recovered backup promoted to current?

## Project and delivery

### DEC-006 — Test strategy baseline

- **Status:** Decision
- The existing SPFx Heft/Jest runner is the repository test framework; no second framework is introduced.
- Regression coverage grows incrementally, starting with pure settings, property-persistence, and browser-cache behavior.
- SPFx host integration remains manually verified until a dedicated host-test harness is confirmed.

### DEC-010 — Documentation authority and progressive reading

- **Status:** Decision
- Confirmed requirements in the core architecture contract and the task-specific normative document section selected by `AGENTS.md` are authoritative.
- Confirmed decision records are authoritative for their recorded choice.
- Registered deviations and accepted exceptions take precedence over treating code as compliant.
- Code and tests are evidence of current implementation, not desired behavior by themselves.
- `AGENTS.md` routes contributors to the smallest sufficient document sections; unrelated sections are not mandatory reading.

### DEV-013 — Version values can differ before release

- **Status:** Deviation
- `package.json`, SharePoint package configuration, and README history can contain different development-time versions.
- The release workflow derives npm and SharePoint versions from the release tag.
- The release tag is treated as the version source only under ASM-001.

### DEBT-004 — Partial automated project coverage

- **Status:** Technical debt
- Focused pure tests cover administrator settings normalization, property-persistence hand-off, and core browser-cache behavior.
- React panel interaction, SPFx host integration, personal settings, migration breadth, policy, broader source mapping, full orchestration, and host failure contracts still lack complete automated regression protection. Pure Exchange orchestration, mailbox identity/discovery, audience discovery and obsolete-request handling have focused mock/helper tests.
- `npm run build` remains the production test, build, and package verification command.

### ASM-001 — Release version source

- **Status:** Assumption
- A semantic release tag is the release version source.
- The release workflow derives package versions from the tag.

## Change discipline

Future agents do not select a source-policy schema, migration default, retry policy, log schema, retention rule, persistence-failure UX, backup strategy, or test architecture without confirmation.

When a decision is confirmed:

1. add or update its record in this register;
2. update the applicable normative contract section;
3. update specialized documentation and matrices;
4. update README only when user-visible or operational behavior changes; and
5. keep a deviation until implementation actually meets the decision.
