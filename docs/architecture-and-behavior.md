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
7. Source services convert external data to the local `ICalendarEvent` model.
8. Local renderers display Day, Week, and Month. `SearchResultsView` renders search results.

**Fact:** Search hides the active calendar view but keeps it mounted. This preserves its view state and scroll position when the user clears a search.

**Fact:** The initial retrieval window starts on the first day of the month three months before today. It ends at the start of the month four months after today.

**Fact:** Navigation requests missing months that intersect the visible Day, Week, or Month range. The runtime caches successful source and month combinations. Failed combinations remain eligible for retry.

**Fact:** Event results are merged by source identifier and event identifier. Results from an obsolete load generation are ignored.

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

**Fact:** `onThemeChanged` updates theme-dependent state and CSS only. SPFx owns the subsequent render. The initial React mount occurs after `onInit` completes, and `onDispose` unmounts the React tree.

**Fact:** `CalendarSettingsService` validates JSON shapes. It rejects duplicate administrator source identities and duplicate ICS catalog URLs. It also derives the smallest personal override set from effective settings.

**Fact:** `MyCalendars` creates source service instances for each load. It starts independent asynchronous tasks and appends successful event groups to component state.

**Fact:** The settings panels contain source discovery and selection flows. They do not write events to those sources.

**Fact:** The administrator manager registers its React lifecycle through the SPFx custom property-field contract. Opening the personal settings panel closes an open SharePoint property pane through the SPFx property-pane accessor.

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

The repository defines `ICalendarEvent` as its canonical local event model.

### Normative requirements

- A normalized event **MUST** have a stable source event identifier.
- A normalized event **MUST** have a title, start value, end value, and source identifier.
- Start and end values **MUST** be valid ISO date-time strings.
- A source adapter **MUST** reject an event when it cannot produce valid date values.
- A source adapter **MUST NOT** replace an invalid source date with the current time.
- A normalized event **MUST** identify its source type before rendering when the UI can show a source logo.
- A renderer **MUST NOT** expose an internal source type key as a user-facing source name.
- A normalized event **SHOULD** include a display color.
- A normalized event **MAY** include an all-day flag, attendees, a location, progress, draft status, or a source deep link.
- The common contract **MUST NOT** imply write access to the source.

### Current source mapping

**Fact:** Exchange and Group events can include attendee, organizer, online-meeting join-link, and web-link data.

**Fact:** SharePoint list events use explicit or detected field mappings. Items without a title are skipped.

**Fact:** Planner tasks without a start or due date are skipped. A single available date becomes both start and end.

**Fact:** Teams Shifts can emit a draft or shared shift event. Draft shifts use italic display text.

**Fact:** Events can carry a concrete source display name. The shared source registry supplies a friendly type name and default Fluent UI icon when an adapter does not supply an override.

**Fact:** ICS URLs are not converted to events. The settings flow creates an Exchange Online subscription deep link.

## Rendering

**Fact:** A local Fluent UI renderer is active for Day, Week, and Month.

**Fact:** The shared toolbar selects Day, Week, or Month through one compact dropdown. Each option has a view-specific icon. On very small screens, the closed selector hides its text and keeps the selected icon visible.

**Fact:** The administrator view is the initial default. A toolbar view selection is stored immediately as an explicit personal override, including when it equals the administrator default. Reset to defaults removes that override and restores the current administrator default.

**Fact:** Month uses a Sunday-first grid with seven columns and six rows. Each day cell scrolls its own event list.

**Fact:** Week starts at the selected date. It renders seven consecutive calendar days, or the next five weekdays when weekends are disabled. Day cards use a synchronized 24-hour timeline.

**Fact:** Day uses the same timeline calculations as Week. It shows all-day events, overlapping timed events, a current-time indicator, and an accessible event-details dialog.

**Fact:** In Day and Week, the preferred start time is the first timeline time below the sticky all-day section. The all-day section does not add an offset to the preferred start time.

**Fact:** Month, Week, Day, Search, and event details use the same source-icon fallback. An explicit user setting can hide the icon.

**Fact:** Timeline settings use schema version 3. Administrator settings supply a preferred start time, visible-hour count, slot duration, and weekend policy. Personal settings can override the preferred start time and visible-hour count.

**Fact:** Search uses `SearchResultsView`. It filters a prepared lower-case index of event titles and locations.

**Fact:** The repository contains inactive Schedule view code. Schedule is not a supported view.

## Error behavior

### Normative requirements

- A failure in one source **MUST NOT** remove valid results from another source.
- The UI **MUST** identify each enabled source service as loading, ready, or failed.
- A failed source **MUST** have a visible source-level error indication.
- A service **SHOULD** preserve enough error context for diagnosis without exposing sensitive event content.
- Invalid event data **MUST NOT** create a fabricated event.

### Current behavior

**Fact:** `MyCalendars` loads sources independently and preserves successful partial results.

**Fact:** Active source retrieval methods propagate failures to the coordinator. The coordinator keeps failed source/month combinations eligible for retry.

**Fact:** Exchange, Group, and SharePoint mappings reject missing or invalid required dates. They do not replace invalid dates with the current time.

## Installation, permissions, and build

See the [README](../README.md#microsoft-graph-permissions) for permissions, installation, and build instructions. This document does not duplicate those operational details.
