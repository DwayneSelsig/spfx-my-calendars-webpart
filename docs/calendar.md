# Calendar

Read only the section selected by [AGENTS.md](../AGENTS.md). The requirements in this document are normative. Statements labeled **Fact** describe verified current implementation. Related intent and known conflicts are recorded by ID in [Decisions, deviations, and open questions](decisions-deviations-and-open-questions.md).

## Loading, range, and cache

**Read when:** changing `MyCalendars`, source coordination, visible ranges, refresh, deduplication, loading status, partial failures, or runtime caches. Related records: DEC-004, DEC-007, DEC-008, and DEC-017.

Read the common and applicable source section in [Sources and permissions](sources-and-permissions.md) as well only when an adapter, endpoint, mapping, permission, or source-specific failure boundary changes.

### Coordination contract

- The web part **MUST** share one scheduler for Graph retrieval, discovery and settings requests. Each HTTP operation takes one slot, including a continuation page or a batch envelope. The composition root applies the accepted administrator limit.
- `CalendarRangeCoordinator` **MUST** schedule independent source/month work, reuse pending work, and let a failed source/month retry. Navigation changes priority without invalidating the source generation.
- `MyCalendars` **MUST** keep the current date, view and search state when pages arrive. An obsolete configuration generation **MUST NOT** publish results or write to the current cache.

### Initialization and navigation

1. The web part loads administrator and personal settings, evaluates audiences, discovers applicable administrator Exchange mailbox rules, and resolves effective settings before React mounts.
2. After mounting, the coordinator starts two normal-priority background jobs. Forward loading starts with the current month and includes 13 full months ahead. Backward loading starts with the previous month and includes 13 full months behind. The ranges share a boundary without overlapping. Dates use local months with an inclusive start and exclusive end.
3. Exchange and Microsoft 365 Group retrieval sends one date window per month. SharePoint, Planner and Teams Shifts each retrieve a collection per source and reuse it across month consumers. Discovery and sources can progress independently. Results from each successful page become available progressively.
4. Navigation prioritizes the selected month and adjacent months actually used by the renderer. The coordinator promotes queued and future pages of existing work, demotes obsolete visible work, and fills skipped months at normal priority. Empty successful results count as complete. Failed ranges remain eligible for retry.

### Browser appointment cache

- `enableCache` and `cacheDurationMinutes` are administrator controls. The cache defaults to enabled with a ten-minute freshness threshold. Duration is normalized to 1–60 minutes. Expiry by itself schedules no network activity.
- IndexedDB stores all successfully obtained normalized appointments, including navigation ranges and results outside loaded months from a Graph collection. Each entry is isolated by tenant, user, web-part instance and a signature of result-affecting settings. Discovery data and personal settings have separate lifetimes and storage.
- Fresh, complete results suppress retrieval. Stale or partial results remain available while their queries are fetched again. Each successful page can be stored independently. A request becomes complete only after the final page succeeds. A successful replacement can be empty; failure preserves usable previous and partial data.
- Compatible legacy `localStorage` rows migrate on successful IndexedDB commit. Legacy rows from previously unpaged endpoints are used as stale display data. With storage disabled, the instance clears its browser appointment cache. Unavailable or full storage falls back to runtime appointments.
- Manual refresh rediscovers applicable administrator Exchange mailbox rules first, then forces the initial and visible ranges. With caching enabled, results remain visible until successfully replaced. Without caching, refresh clears results first.

### Identity and failures

The coordinator deduplicates appointments by source and event ID. Automatic discovery and configured sources load independently. Stable source identity, policy precedence and administrator assignment provenance are resolved by CalendarSettingsService. Source-family status is derived from pending and failed work. Relevant visible failures and discovery errors remain visible after the toolbar returns to Refresh.

### Loading verification focus

Heft/Jest covers scheduler queues, live limits, page boundaries, priority promotion, independent month jobs, stale or partial cache data, failed and empty months, and obsolete generations. Browser storage and live SPFx host behavior require browser and host verification.

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

The top action bar **MUST** keep search directly visible outside an overflow menu. Its layout **MUST** follow the web-part content width: at 600 px and above, Settings and Refresh show icons and text; from 360 to below 600 px, they show icons only; below 360 px, the actions appear right-aligned above a full-width search field. Icon-only actions **MUST** retain localized accessible names and tooltips. Resizing **MUST NOT** discard search input or focus. The existing loading-status action replaces Refresh as before. See DEC-023.

`CalendarToolbar` owns date navigation, today, date selection, and Day/Week/Month selection. Selecting a view shows its result immediately and awaits storage of an explicit personal `defaultView`, even when it equals the administrator default. Failed storage restores the confirmed view with a localized error. Concurrent preference writes are blocked. The personal settings panel previews each draft edit without storage; Cancel or a failed Save restores confirmed effective settings. Reset removes the explicit personal value after successful deletion.

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
