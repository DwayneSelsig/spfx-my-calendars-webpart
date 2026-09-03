# Architecture and Behavior Contract

**Read for every behavior or architecture change.** This is the small, cross-cutting core contract. Use [AGENTS.md](../AGENTS.md) to select only the task-specific documents and decision records that are also required.

## Document semantics

The key words **MUST**, **MUST NOT**, **SHOULD**, **SHOULD NOT**, and **MAY** have the meanings in [RFC 2119](https://www.rfc-editor.org/rfc/rfc2119) and [RFC 8174](https://www.rfc-editor.org/rfc/rfc8174) only when they appear in all capitals.

| Label | Meaning |
| --- | --- |
| Requirement | Confirmed normative product or architecture contract. |
| Fact | Behavior verified in current code or configuration. |
| Decision | Confirmed choice recorded in a decision domain file. |
| Intention | Confirmed desired behavior that is not necessarily implemented. |
| Deviation | Known difference between confirmed intent and current implementation. |
| Technical debt | Current code or structure that is not desired architecture. |
| Assumption | Working interpretation that has not been confirmed. |
| Open question | A decision that has not been made. |

## Source-of-truth order

Use this order when sources disagree:

1. Requirements in this core contract and the task-specific normative document section selected by `AGENTS.md`, plus confirmed decision records.
2. Registered deviations and accepted exceptions in the decisions register.
3. Code and tests as evidence of current implementation.

README is operational and user-facing. Graphify is a navigation aid. Neither overrides confirmed intent.

When code conflicts with confirmed intent and no deviation is registered, an agent **MUST** report the conflict. It **MUST NOT** silently rewrite intent, silently rewrite code, or expand the task.

An intention, deviation, technical-debt item, assumption, open question, or inactive component **MUST NOT** be described as desired or implemented architecture.

## Product boundary

- The web part **MUST** aggregate source events and tasks as read-only data.
- It **MUST NOT** create, change, or delete events or tasks in a source system.
- It **MAY** open a source calendar, event, task, or subscription flow through a deep link.
- It **MUST NOT** claim that an inactive view, inert model type, or unimplemented policy is supported behavior.
- SharePoint Web Part, SharePoint full-page, Teams personal app, Teams tab, Office, and Outlook **MUST** be treated as supported hosts. Host-sensitive changes still require verification in affected hosts.

**Fact:** an administrator in this product is a person who can edit web-part properties. This does not imply a tenant administrator role; tenant administration can still be required for Graph permission approval.

## Architectural boundaries

- `MyCalendarsWebPart` **MUST** own SPFx lifecycle, Graph-client composition, settings composition, audience evaluation, theme projection, and persistence callbacks.
- `CalendarSettingsService` **MUST** own settings validation, migration, normalization, precedence, and persisted/effective conversion.
- `SettingsStorageService` **MUST** own OneDrive App Folder access and **MUST NOT** own policy resolution.
- `AudienceService` **MUST** own audience discovery and membership evaluation and **MUST NOT** decide source-policy effects.
- A source service **MUST** own external access and conversion to `ICalendarEvent`. It **MUST NOT** decide policy, persistence, or rendering.
- `MyCalendars` **MUST** coordinate loading, range state, deduplication, service status, search state, and renderer selection.
- A renderer **MUST** display prepared local events and **MUST NOT** access storage or source APIs.
- Administrator and personal settings panels **MAY** edit isolated drafts. The web part **MUST** own accepted persistence.
- Source display metadata **SHOULD** come from the shared registry unless a source supplies an explicit runtime override.

See [Components and data flow](components-and-data-flow.md) for the component map and focused flow routes.

## Common event contract

`ICalendarEvent` is the canonical local, read-only event model.

- A normalized event **MUST** have stable event and source identifiers, a title, and valid ISO start/end values.
- An adapter **MUST** reject an item when required dates cannot be made valid and **MUST NOT** substitute the current time.
- A normalized event **MUST** identify its source type before rendering when a logo can be shown.
- A renderer **MUST NOT** expose an internal source-type key as the user-facing source name.
- An event **SHOULD** include a display color and **MAY** include all-day state, attendees, organizer, location, progress, draft state, meeting link, or source link.
- The event contract **MUST NOT** imply write access.

## Failure contract

- A source-family failure **MUST NOT** remove valid results from another independent family.
- The UI **MUST** identify each enabled source service as loading, ready, or failed.
- A failed source **MUST** have a visible source-level indication.
- Invalid event data **MUST NOT** create a fabricated event.
- A failed source/month combination **MUST** remain eligible for retry.
- A service **SHOULD** preserve diagnostic context without exposing sensitive event content.

## Verification contract

**Fact:** the repository contains no project test files. `npm run build` is the current production build and package verification command.

The future test framework and layers are unresolved. A change **MUST NOT** invent a partial test architecture unless the task includes a confirmed test decision.

## Focused documentation

- [Components and data flow](components-and-data-flow.md)
- [Calendar](calendar.md)
- [Sources and permissions](sources-and-permissions.md)
- [Settings and policy](settings-and-policy.md)
- [Decisions, deviations, and open questions](decisions-deviations-and-open-questions.md)
