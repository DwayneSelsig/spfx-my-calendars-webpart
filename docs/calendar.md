# Calendar

Read only the section selected by [AGENTS.md](../AGENTS.md). The requirements in this document are normative. Statements labeled **Fact** describe verified current implementation. Related intent and known conflicts are recorded by ID in [Decisions, deviations, and open questions](decisions-deviations-and-open-questions.md).

## Loading, range, and cache

**Read when:** changing `MyCalendars`, source coordination, visible ranges, refresh, deduplication, loading status, partial failures, or runtime caches. Related records: DEC-004, DEC-007, and DEC-008.

Read the common and applicable source section in [Sources and permissions](sources-and-permissions.md) as well only when an adapter, endpoint, mapping, permission, or source-specific failure boundary changes.

### Coordination contract

- `MyCalendars` **MUST** coordinate source loading, range state, deduplication, source status, search state, and renderer selection.
- Source families **MUST** load independently so a failure in one family does not remove valid results from another.
- A failed source/month combination **MUST** remain eligible for retry.
- A manual refresh **MUST** clear event and range state.
- Loaded event, discovery, and range-cache data **MUST NOT** be persisted.
- Results from an obsolete load generation **MUST NOT** be merged into current state.

### Initialization and settings hand-off

1. `MyCalendarsWebPart.onInit` creates `SettingsStorageService` and attempts to obtain an `MSGraphClientV3`.
2. Administrator and personal settings are loaded, audiences are evaluated, and `resolveCalendarSettings` produces effective settings.
3. SPFx calls `render`; React mounts only after `onInit` completes.
4. `MyCalendars` creates source-service instances and starts enabled service-family loads independently.
5. Source services map external values to `ICalendarEvent`.
6. The coordinator adds source identity and presentation metadata, deduplicates events, and selects Day, Week, Month, or Search rendering.

The detailed persistence and resolution flow is in [Settings and policy](settings-and-policy.md).

### Reset load

For a reset load, the coordinator:

1. increments the load generation and load ID;
2. clears appointments, month/source caches, discovery promises, and the retained Teams Shifts service;
3. creates source-service instances and calculates the initial seven-month request;
4. starts one asynchronous task per enabled service family;
5. loads independent configured sources in parallel where implemented;
6. normalizes presentation metadata and merges results by `sourceId:eventId`;
7. marks a source/month set only after that request succeeds; and
8. ignores append or mark operations from an obsolete load ID or generation.

**Fact:** Exchange is always considered enabled because current-user calendars are automatic. Other service families are enabled by configured enabled sources or their source-type automatic-loading flag.

**Fact:** the initial range begins at the first day of the month three months before the current month and ends at the first day of the month four months after it.

### Visible-range loading

Navigation loads missing months that intersect the visible Day, Week, or Month range. Visible-range loads are serialized through `rangeLoadPromise`. A request includes only service families with at least one known source missing a visible month.

Only successful source/month combinations enter the runtime range cache. Failed combinations remain unmarked and can therefore be retried by later navigation or refresh.

### Runtime-only caches

| Cache/state | Owner | Lifetime/invalidation | Failure behavior |
| --- | --- | --- | --- |
| Appointments | `MyCalendars.state` | Component lifetime; cleared on reset/refresh | Successful partial results remain |
| Loaded months by source | `MyCalendars` | Component cache generation; cleared on reset/settings reload | Only successes are marked |
| Known source IDs by service | `MyCalendars` | Component cache generation | Includes service sentinels for empty families |
| Exchange calendar discovery promise | `MyCalendars` | Until reset; cleared after discovery failure | Later request can retry |
| Planner plan discovery promise | `MyCalendars` | Until reset; cleared after discovery failure | Later request can retry |
| Unified-group and joined-Team discovery promises | `MyCalendars` | Until reset; cleared after relevant failure | Later request can retry |
| Joined Teams in `TeamsShiftsService` | Retained service instance | Until reset; rejected promise is cleared | Later request can retry |
| Planner current user ID | One `PlannerTaskService` instance | One load invocation because the service is recreated | Failure returns `null`; assigned-only filtering is then not applied |
| Search index text | Each in-memory event | Appointment lifetime | Recomputed on demand if absent |

### Identity and partial failures

**Fact:** results are deduplicated by `sourceId:event.id`.

The coordinator exposes `loading`, `ready`, or `error` per service family, not per configured source. The displayed message is a generic family-level summary; detailed errors go to the browser console.

Independent successes are retained, but isolation inside a family varies:

- Exchange isolates individual automatic calendars and manual sources.
- SharePoint isolates configured list sources.
- Planner isolates plan task loads; automatic plan-discovery failure fails that service family.
- Unified Groups isolates group event loads; group or Team discovery failure affects automatic mode.
- Teams Shifts can lose earlier Team results when a later non-404 Team request fails.
- ICS performs no runtime retrieval and is marked ready.

### Loading verification focus

There are no automated tests for successful-month caching, obsolete-load rejection, deduplication, partial failures, automatic versus explicit source selection, or retry eligibility. Until a test architecture is confirmed, `npm run build` is the production verification command.

## Rendering and interaction

**Read when:** changing Day, Week, Month, Search, toolbar navigation, date/time formatting, event layout, event details, or renderer styling. Related records: DEC-003, DEC-009, DEC-011, and DEBT-001.

Also read [Loading, range, and cache](#loading-range-and-cache) only when a view change alters visible-range loading, refresh, status, or cached state.

### Renderer contract

- The active local renderer **MUST** provide Day, Week, and Month.
- A renderer **MUST** display prepared local event data and **MUST NOT** access persistent settings storage or a source API.
- Search **MUST** remain a separate result view while preserving the mounted calendar view.
- Schedule **MUST NOT** be presented as supported while `ScheduleView` remains inactive.
- Visible date and time text **MUST** use the current SharePoint page culture unless a later confirmed localization decision replaces it.
- A renderer **MUST NOT** expose an internal source-type key as a user-facing source name.
- Rendering remains read-only; links may open source or meeting destinations, but renderers do not mutate source data.

### View behavior

**Fact:** `currentView` can be `search`; `previousView` remains Day, Week, or Month.

**Fact:** Month uses a Sunday-first, seven-column, six-row grid. Each cell owns its event-list overflow.

**Fact:** Week starts at the selected date and shows seven days, or five successive weekdays when weekends are hidden. Day and Week use the shared `TimelineDay` calculations and a complete 24-hour timeline.

**Fact:** Day and Week calculate the scroll target from the preferred start time. The sticky all-day section does not add to that time value.

### Search

- Search matches lower-case event title and location only.
- Input of one to three trimmed characters is debounced by 100 ms. Longer input and clearing apply immediately.
- The active calendar stays mounted inside a hidden container during search.
- Clearing search restores the previous calendar view and its scroll state.
- `SearchResultsView` groups and renders already-filtered results; it performs no source loading.

### Navigation and preferences

`CalendarToolbar` owns date navigation, today, date selection, and Day/Week/Month selection. Selecting a view immediately asks the web part to persist that value as an explicit personal `defaultView`, even when it equals the administrator default. Reset removes the explicit personal value.

Weekend visibility can be an administrator default or explicit personal override. Slot duration is administrator-only. The renderer consumes optional personal start-time and visible-hour overrides ahead of administrator base values.

### Locale and formatting

`calendarFormatting` uses `Intl.DateTimeFormat` with the SharePoint culture passed by the web part. It does not force 12- or 24-hour output. A separate 12/24-hour product setting does not exist.

### Event details and links

`EventDetailsDialog` displays one normalized event. It can open Graph-supplied `joinUrl` or `webLink` through `safeOpen` in a new tab.

Exchange and Microsoft 365 Group mappings can currently supply those links. Planner, SharePoint, and Teams Shifts mappings do not currently produce event deep links.

### Inactive Schedule view

`ScheduleView.tsx` exists but is not imported by the coordinator. A commented command-bar branch refers to Schedule. Its presence is technical debt, not supported behavior or product intent.

### Rendering verification focus

There are no automated tests for date-range calculations, overlap layout, locale formatting, search behavior, view-state preservation, or renderer link behavior. Verify affected interactions manually and run `npm run build`.
