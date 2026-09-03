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

### DEBT-001 — Inactive Schedule view

- **Status:** Technical debt
- `ScheduleView.tsx` exists but is not imported by the runtime.
- A commented command-bar branch refers to Schedule.
- Its presence is not support or intent.

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

### INT-001 — Mandatory-source failure records

- **Status:** Intention
- Store a dated record in the OneDrive App Folder when a Graph error prevents a mandatory source from loading.
- Retry the mandatory source after a delay.
- Do not store source event content unless a later decision explicitly permits it.
- Record schema, delay, retry limit, retention, cleanup, and UI behavior remain open.

### DEV-004 — Per-source logo precedence is inconsistent

- **Current state:** a configured manual Exchange source can override the source-type logo value.
- **Current state:** SharePoint, Planner, Unified Group, and Teams Shifts ignore entry-level `showSourceLogo` at runtime.
- **Desired state:** type default followed by allowed per-source override, subject to administrator policy.

### DEV-005 — `MailboxSettings.Read` has no current consumer

- **Current state:** both permission manifests request `MailboxSettings.Read`.
- **Current state:** no `/me/mailboxSettings` endpoint is called; `UserHelper` uses `/me/calendar/getSchedule`, covered by `Calendars.Read`.
- **Required follow-up:** remove or justify the permission in a separately scoped runtime/configuration change.

### DEV-006 — Invalid Planner dates can fail a plan load

- **Current state:** Planner skips a task with no dates, but a malformed non-empty date can reach `toISOString()` and throw.
- **Desired state:** reject only the invalid item under DEC-008.

### DEV-007 — Exchange mailbox validation uses the generic HTTP client

- **Current state:** mailbox validation calls Graph with SPFx `HttpClient`; calendar discovery uses the authenticated Graph client.
- **Risk:** validation can reject an accessible mailbox before the authoritative Graph-client call.
- **Desired state:** not confirmed; do not refactor without scope and verification.

### DEV-008 — Some “all” source modes do not page all data

- **Current state:** Planner group discovery uses `$top=100` without following `@odata.nextLink`.
- **Current state:** Exchange and Group events use `$top=500` without paging.
- **Current state:** SharePoint runtime items and several discovery calls do not follow paging.
- **Consequence:** “all” is limited to the implemented returned pages, not a guaranteed complete tenant data set.

### DEV-009 — Teams Shifts failure isolation is coarser than Team scope

- **Current state:** joined Teams are queried sequentially; a non-404 error aborts the call and drops earlier collected Team results.
- **Required contract:** source-family isolation remains required; Team-level isolation inside Shifts is not confirmed.

### DEV-012 — Source registry and inactive UI contain obsolete ICS wording

- **Current state:** `CalendarSourceRegistry` and unused `AddCalendarDialog` describe adding or pasting ICS content.
- **Confirmed behavior:** active UI opens an Outlook subscription link and does not parse content.

### DEBT-002 — Inactive Add Calendar dialog

- **Status:** Technical debt
- `AddCalendarDialog.tsx` is not imported and duplicates an older subset of the integrated settings-panel flow.
- It is not an architectural component of the active UI.

### DEBT-003 — Unused service dependencies

- **Status:** Technical debt
- Planner, SharePoint, Teams Shifts, and Unified Group services retain unused `HttpClient` fields.
- This does not establish a desired dual-client architecture.

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
- Policy applies to administrator-assigned event-source entries, not scalar defaults or ICS catalog entries.
- Membership has two states: `optional` and `mandatory`.
- Optional sources can be disabled or removed; mandatory sources cannot.
- Allowed presentation and source-option overrides are a separate field-level dimension.
- The combination must represent fully overridable, partially restricted, and mandatory/locked outcomes.
- The serialized schema is not defined.

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

### DEC-014 — Administrator configuration fields

- **Status:** Decision
- Organization color, source-type logo defaults, automatic Planner/Group/Shifts loading defaults, and the Planner automatic assigned-to-me default are administrator-configurable product settings.
- Presence only in model/default values does not satisfy the required administrator UI.

### DEV-001 — Administrator source policy is not implemented

- **Current state:** `IAdminAssignedSource` has no membership or allowed-override policy.
- **Current state:** users can rename, recolor, disable, or remove every applicable administrator source.
- **Desired state:** enforce DEC-005 and DEC-013.
- **Missing decision:** persisted schema and migration defaults.

### DEV-002 — Audience support is narrower than the confirmed model

- **Current state:** discovery returns only mail-disabled security groups.
- **Current state:** entries require at least one group; normalization drops entries with no groups.
- **Desired state:** support DEC-012 group types and empty-means-everyone.
- **Compatible behavior:** selected groups use OR semantics and failures are fail-closed.

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

- **Status:** Open question
- What names and migration represent membership and allowed overrides?
- What policy applies to existing administrator sources after migration?
- How are source-specific capabilities versioned as models evolve?

### OQ-004 — Personal persistence failure UX

- **Status:** Open question
- Should failed saves roll back, remain pending, or show retry state?
- How does the product distinguish a missing file from unavailable storage?

### OQ-005 — Administrator backup semantics

- **Status:** Open question
- Should backup remain a same-save mirror or rotate the previous validated revision?
- When is a recovered backup promoted to current?

## Project and delivery

### DEC-006 — Test strategy timing

- **Status:** Decision
- Automated tests will be introduced as one repository-wide initiative.
- Documentation work does not invent a partial future test architecture.

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

### DEBT-004 — No automated project tests

- **Status:** Technical debt
- The repository contains no project test files.
- `npm run build` is the current production verification.
- Settings, migration, policy, mapping, caching, and failure contracts have no regression protection.

### ASM-001 — Release version source

- **Status:** Assumption
- A semantic release tag is the release version source.
- The release workflow derives package versions from the tag.

### OQ-006 — Test architecture

- **Status:** Open question
- Which framework and test layers will the repository-wide initiative use?
- Which pure settings functions and source mappings form the first test set?

## Change discipline

Future agents do not select a source-policy schema, migration default, retry policy, log schema, retention rule, persistence-failure UX, backup strategy, or test architecture without confirmation.

When a decision is confirmed:

1. add or update its record in this register;
2. update the applicable normative contract section;
3. update specialized documentation and matrices;
4. update README only when user-visible or operational behavior changes; and
5. keep a deviation until implementation actually meets the decision.
