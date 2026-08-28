# Architecture and Behavior Contract

## Document conventions

This document contains normative and descriptive sections.

The key words **MUST**, **MUST NOT**, **SHOULD**, **SHOULD NOT**, and **MAY** have the meanings in [RFC 2119](https://www.rfc-editor.org/rfc/rfc2119) and [RFC 8174](https://www.rfc-editor.org/rfc/rfc8174) only when they appear in all capitals. Lower-case forms have no normative meaning.

This document uses controlled technical prose. It follows ASD-STE100 Simplified Technical English where practical. It uses short sentences, active voice, and consistent terms.

Descriptive statements use these labels:

- **Fact:** behavior that the current code or configuration shows.
- **Intention:** confirmed desired behavior that is not necessarily implemented.
- **Technical debt:** current code that is not part of the desired architecture.
- **Assumption:** a working interpretation that is not confirmed.
- **Open question:** a decision that has not been made.

For incomplete behavior and unresolved decisions, see [Decisions, deviations, and open questions](decisions-deviations-and-open-questions.md).

## Product boundary

### Normative requirements

- The web part **MUST** treat source events and tasks as read-only data.
- The web part **MUST NOT** create, change, or delete events or tasks in a source system.
- The web part **MAY** open a source calendar, event, task, or subscription flow through a deep link.
- The web part **MUST NOT** claim support for a view or source flow that the active runtime cannot provide.

### Current system context

**Fact:** My Calendars is one SPFx web part. SharePoint is the primary product context. The manifest also declares Teams personal app, Teams tab, and SharePoint full-page hosts.

**Fact:** The runtime uses Microsoft Graph for Exchange, mailbox data, OneDrive storage, SharePoint lists, Planner, Groups and Teams, Entra group membership, and Teams Shifts.

**Fact:** An administrator is a user who can edit the web part properties. This term does not imply a tenant administrator role. A tenant administrator can still be required to approve Graph permissions.

## Runtime data flow

The application uses this flow:

1. `MyCalendarsWebPart` loads administrator settings from web part properties.
2. It recovers the backup or migrates legacy settings when this action is possible.
3. It loads personal settings from the OneDrive App Folder. It can migrate the legacy personal settings file.
4. `AudienceService` evaluates the current user's membership of configured Entra security groups.
5. `CalendarSettingsService` resolves administrator settings, audience results, and personal overrides into runtime settings.
6. `MyCalendars` starts enabled source loads in parallel.
7. Source services convert external data to the common PnP `IEvent` model.
8. PnP Calendar renders Day, Week, and Month. `SearchResultsView` renders search results.

**Fact:** Search hides the PnP view but keeps it mounted. This reduces the cost when the user clears a search.

**Fact:** The current retrieval window starts on the first day of the month three months before today. It ends on the last day of the month three months after today. Navigation does not change this window.

**Technical debt:** A user can navigate outside the loaded window. The visible period can then contain no events even when source events exist.

## Component boundaries

### Normative requirements

- The web part class **MUST** own SPFx lifecycle work, client creation, settings composition, audience evaluation, and settings persistence.
- The settings service **MUST** own settings validation, migration, normalization, precedence, and runtime projection.
- A source service **MUST** own external source access and conversion to the common event contract.
- A source service **MUST NOT** decide administrator or user policy.
- The main React component **MUST** coordinate source loading and source status.
- A renderer **MUST** display prepared event data.
- A renderer **MUST NOT** access persistent settings storage or an external calendar API.
- A settings panel **MAY** edit a draft settings object. The web part **MUST** persist the accepted result.

### Current responsibilities

**Fact:** `MyCalendarsWebPart` creates the Graph client, loads settings, resolves audiences, applies the theme, and passes effective settings to `MyCalendars`.

**Fact:** `CalendarSettingsService` validates JSON shapes. It rejects duplicate administrator source identities and duplicate ICS catalog URLs. It also derives the smallest personal override set from effective settings.

**Fact:** `MyCalendars` creates source service instances for each load. It starts independent asynchronous tasks and appends successful event groups to component state.

**Fact:** The settings panels contain source discovery and selection flows. They do not write events to those sources.

## Settings and source policy

### Storage and precedence

**Fact:** The current administrator settings and their backup are serialized web part properties. Legacy web part settings can supply migration input.

**Fact:** Personal settings are stored in `user-calendar-settings.json` under the `SPFx-My-Calendar-Webpart` folder in the OneDrive App Folder. The storage service can migrate `calendar-settings.json`.

**Fact:** Effective settings use this precedence:

1. Hardcoded defaults provide a complete base.
2. Valid administrator settings replace the base values.
3. Audience membership selects assigned administrator sources and ICS catalog entries.
4. Personal values replace supported administrator defaults.
5. The current theme primary color can replace the configured organization color at runtime.

**Fact:** The application does not persist loaded event data.

### Audience rules

- Only confirmed group membership **MUST** grant access to an audience-assigned source or ICS catalog entry.
- An audience evaluation failure **MUST NOT** grant access to an unconfirmed assignment.

**Fact:** Audience checks use `/me/checkMemberGroups` in batches. The browser session cache stores positive and negative results for five minutes.

### Administrator source policy

**Fact:** The current implementation treats every administrator source as an overridable default. A user can rename it, change its color, disable it, or remove it for that user.

**Intention:** Policy will vary by source. A future source can be fully overridable, mandatory, or partly restricted.

**Intention:** A Graph failure for a mandatory source will create a dated record in the OneDrive App Folder. The application will retry the source after a delay.

**Open question:** The source-policy schema, retry delay, record format, retention period, and permitted record content are not defined.

## Caches and persisted data

**Fact:** `UserHelper` stores the current user email and mailbox working-hours data in local storage. Each entry has a twelve-hour lifetime.

**Fact:** `AudienceService` stores membership results in session storage. Each entry has a five-minute lifetime.

**Fact:** These lifetimes are implementation values. They are not product invariants.

**Fact:** Personal preferences are the only product data that the current application stores in OneDrive. Future mandatory-source error records are confirmed intent but are not implemented.

## Common event contract

The repository augments the PnP `IEvent` type with source metadata.

### Normative requirements

- A normalized event **MUST** have a stable source event identifier.
- A normalized event **MUST** have a title, start value, end value, and source identifier.
- Start and end values **MUST** be valid ISO date-time strings.
- A source adapter **MUST** reject an event when it cannot produce valid date values.
- A source adapter **MUST NOT** replace an invalid source date with the current time.
- A normalized event **MUST** identify its source type before rendering when the UI can show a source logo.
- A normalized event **SHOULD** include a display color.
- A normalized event **MAY** include an all-day flag, attendees, a location, progress, draft status, or a source deep link.
- The common contract **MUST NOT** imply write access to the source.

### Current source mapping

**Fact:** Exchange and Group events can include attendee, organizer, online-meeting, and web-link data.

**Fact:** SharePoint list events use explicit or detected field mappings. Items without a title are skipped.

**Fact:** Planner tasks without a start or due date are skipped. A single available date becomes both start and end.

**Fact:** Teams Shifts can emit a draft or shared shift event. Draft shifts use italic display text.

**Fact:** ICS URLs are not converted to events. The settings flow creates an Exchange Online subscription deep link.

## Rendering

**Fact:** PnP Calendar is the active renderer for Day, Week, and Month.

**Fact:** Search uses `SearchResultsView`. It filters a prepared lower-case index of event titles and locations.

**Fact:** The repository contains inactive custom Day, Week, Month, and Schedule view code. It also contains an inactive toolbar.

**Intention:** The current hybrid behavior will remain until a replacement for PnP Calendar is selected.

**Open question:** The future renderer and supported view set are not selected.

## Error behavior

### Normative requirements

- A failure in one source **MUST NOT** remove valid results from another source.
- The UI **MUST** identify each enabled source service as loading, ready, or failed.
- A failed source **MUST** have a visible source-level error indication.
- A service **SHOULD** preserve enough error context for diagnosis without exposing sensitive event content.
- Invalid event data **MUST NOT** create a fabricated event.

### Current deviations

**Fact:** `MyCalendars` loads sources independently and preserves successful partial results.

**Deviation:** Several source services catch Graph errors and return an empty array. The caller can then report a ready state instead of a failed state.

**Technical debt:** Exchange and Group date conversion can replace an invalid date with the current time. SharePoint conversion can also use the current time when fields are absent.

**Intention:** Source-level errors will remain visible. Invalid events will be rejected instead of receiving invented dates.

## Installation, permissions, and build

See the [README](../README.md#microsoft-graph-permissions) for permissions, installation, and build instructions. This document does not duplicate those operational details.
