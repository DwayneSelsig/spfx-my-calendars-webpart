# Calendar

Read only the section selected by [AGENTS.md](../AGENTS.md). The requirements in this document are normative. Statements labeled **Fact** describe verified current implementation. Related intent and known conflicts are recorded by ID in [Decisions, deviations, and open questions](decisions-deviations-and-open-questions.md).

## Loading, range, and cache

**Read when:** changing `MyCalendars`, source coordination, visible ranges, refresh, deduplication, loading status, partial failures, or runtime caches. Related records: DEC-004, DEC-007, DEC-008, and DEC-017.

Read the common and applicable source section in [Sources and permissions](sources-and-permissions.md) as well only when an adapter, endpoint, mapping, permission, or source-specific failure boundary changes.

### Coordination contract

- `MyCalendars` **MUST** coordinate source loading, range state, deduplication, source status, search state, and renderer selection.
- Source families **MUST** load independently so a failure in one family does not remove valid results from another.
- A failed source/month combination **MUST** remain eligible for retry.
- A manual refresh **MUST** invalidate range state and force source retrieval. With administrator caching enabled, existing appointments **MUST** remain visible until successful replacements are available; without it, reset behavior clears them.
- Loaded discovery and navigation-range data **MUST NOT** be persisted. When the administrator enables appointment caching, normalized events and successful source/month state for the initial seven-month range **MAY** be persisted in browser `localStorage` under DEC-017.
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
2. clears runtime month/source caches, discovery promises, and the retained Teams Shifts service;
3. hydrates compatible cached appointments before retrieval when administrator caching is enabled;
4. creates source-service instances and calculates the initial seven-month request;
5. starts one asynchronous task per enabled service family needing retrieval;
6. loads independent configured sources in parallel where implemented;
7. normalizes presentation metadata and merges results by `sourceId:eventId`;
8. marks and optionally persists a source/month set only after that request succeeds; and
9. ignores append or mark operations from an obsolete load ID or generation.

**Fact:** Exchange is always considered enabled because current-user calendars are automatic. Other service families are enabled by configured enabled sources or their source-type automatic-loading flag.

**Fact:** the initial range begins at the first day of the month three months before the current month and ends at the first day of the month four months after it.

### Visible-range loading

Navigation loads missing months that intersect the visible Day, Week, or Month range. Visible-range loads are serialized through `rangeLoadPromise`. A request includes only service families with at least one known source missing a visible month.

Only successful source/month combinations enter the runtime range cache. Failed combinations remain unmarked and can therefore be retried by later navigation or refresh.

### Runtime-only caches

| Cache/state | Owner | Lifetime/invalidation | Failure behavior |
| --- | --- | --- | --- |
| Appointments | `MyCalendars.state` | Component lifetime; replaced on settings reset and retained during manual refresh | Successful partial results remain |
| Loaded months by source | `MyCalendars` | Component cache generation; cleared on reset/settings reload | Only successes are marked |
| Known source IDs by service | `MyCalendars` | Component cache generation | Includes service sentinels for empty families |
| Exchange calendar discovery promise | `MyCalendars` | Until reset; cleared after discovery failure | Later request can retry |
| Planner plan discovery promise | `MyCalendars` | Until reset; cleared after discovery failure | Later request can retry |
| Unified-group and joined-Team discovery promises | `MyCalendars` | Until reset; cleared after relevant failure | Later request can retry |
| Joined Teams in `TeamsShiftsService` | Retained service instance | Until reset; rejected promise is cleared | Later request can retry |
| Planner current user ID | One `PlannerTaskService` instance | One load invocation because the service is recreated | Failure returns `null`; assigned-only filtering is then not applied |
| Search index text | Each in-memory event | Appointment lifetime | Recomputed on demand if absent |

### Optional persistent appointment cache

- Administrator settings `enableCache` and `cacheDurationMinutes` exclusively control the feature. Defaults are enabled and ten minutes; duration is normalized to 1–60 whole minutes.
- The cache key includes normalized tenant, user, and web-part instance identity. New entries have no schema version; a schema-version field in an existing entry is ignored. The entry carries a stable signature of result-affecting effective settings and is validated structurally as one unit.
- Cached canonical events are always reprocessed for presentation, search indexing, and deduplication. Fresh source/month segments suppress GET requests; stale segments remain visible while they refresh in the background.
- Cache expiry by itself **MUST NOT** schedule retrieval while the component remains mounted. When a later load finds stale segments, that retrieval **MUST** use the existing loading and toolbar-status workflow.
- A successful source GET replaces that source/month segment, including with an empty result. A failed GET retains the prior segment and remains visibly failed.
- Persistent segments are restricted to the moving initial seven-month range. Visible months outside that range remain runtime-only.
- Manual refresh keeps appointments visible, bypasses freshness, and refreshes the initial and currently visible ranges. Disabling cache removes the current identity's entry. Any read, parse, identity, signature, segment, or event validation failure rejects the entire entry and attempts to remove it; no partial data is hydrated. Normal source loading rebuilds the cache. Unavailable, read-only, or full storage degrades to uncached loading, and physical removal is best-effort.

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

Focused unit tests cover cache-duration normalization, fresh/stale boundaries, versionless persistence, successful empty segment replacement, structural rejection, compatibility with an existing version field, best-effort removal, and recovery after browser-storage failures. Source orchestration and host storage behavior still require manual verification.

There are no automated tests for successful-month caching, obsolete-load rejection, deduplication, partial failures, automatic versus explicit source selection, or retry eligibility. Until a test architecture is confirmed, `npm run build` is the production verification command.

## Rendering and interaction

**Read when:** changing Day, Week, Month, Search, toolbar navigation, date/time formatting, event layout, event details, or renderer styling. Related records: DEC-003, DEC-009, DEC-011, DEC-018, and DEC-019.

Also read [Loading, range, and cache](#loading-range-and-cache) only when a view change alters visible-range loading, refresh, status, or cached state.

### Renderer contract

- The active local renderer **MUST** provide Day, Week, and Month.
- A renderer **MUST** display prepared local event data and **MUST NOT** access persistent settings storage or a source API.
- Search **MUST** remain a separate result view while preserving the mounted calendar view.
- Schedule **MUST NOT** be presented as supported.
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

`EventDetailsDialog` displays one normalized event. It can open a `joinUrl` or an exact source-item `webLink` through `safeOpen` in a new tab. The Open action is omitted when an adapter cannot supply an exact item destination; it does not fall back to a general source or application page.

Organizer details are displayed only for meetings with at least one normalized attendee. Descriptions without a format are rendered as escaped plain text. Descriptions explicitly marked as HTML are sanitized before rendering; active content, inline handlers, inline styles, embeds, and unsafe URLs are removed, and retained links open in a protected new tab.

Exchange and Microsoft 365 Group mappings use Graph-provided event links. SharePoint derives the list path from the Graph list item's `webUrl` and constructs its display-form URL from that path and the item ID; it does not open the internal `FileRef` target because classic calendar items can otherwise download an empty `_.000` file. Planner constructs the current Planner task URL from the plan, task, and tenant identifiers. Teams Shifts does not produce an event deep link because no reliable exact shift destination is available.

SharePoint events display their persisted or lazily resolved site name separately from the calendar name. Missing legacy metadata uses a neutral localized fallback; Event Details does not call Graph.

Exchange and Microsoft 365 Group events can display `showAs` as calendar availability. The calendar color remains a solid identity edge while icons and secondary fill, outline, or pattern styling communicate availability without color alone. `free` uses the untinted underlying theme surface, `tentative` uses diagonal bands, and `oof` uses cross-hatching made from two opposing diagonal patterns. A meeting response is displayed only for events retrieved through `/me`; values from another mailbox or a group calendar are not treated as the signed-in user's response. Declined `/me` events remain visible with a muted surface, a response glyph, and a struck-through title. Day, Week, Month, Search, and Event Details use the same status labels and presentation rules.

### Rendering verification focus

Focused pure tests cover event-status presentation. There are no automated tests for date-range calculations, overlap layout, locale formatting, search behavior, view-state preservation, renderer links, or complete React interaction. Verify affected interactions manually and run `npm run build`.
