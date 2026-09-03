# Components and Data Flow

**Read when:** changing component ownership, SPFx lifecycle, composition boundaries, or hand-offs between the web part, settings, sources, coordinator, and renderers.

This document describes verified current responsibilities. Normative cross-cutting boundaries are in [Architecture and behavior](architecture-and-behavior.md).

## Component map

| Component | Responsibility | State or side effect |
| --- | --- | --- |
| `MyCalendarsWebPart` | SPFx composition root; settings/audience/theme composition and React lifecycle | Web-part properties, OneDrive callbacks, effective settings |
| `ICalendarSettings.ts` | Persisted, effective, source, and default data definitions | Static types/defaults; no precedence logic |
| `CalendarSettingsService` | Validation, normalization, migration, resolution, minimal personal overrides | Pure except identifier generation |
| `SettingsStorageService` | Current/legacy personal JSON in OneDrive App Folder | Graph I/O; no policy resolution |
| `AudienceService` | Group discovery and current-user membership | Five-minute session cache; no policy effects |
| `MyCalendars` | Source coordination, visible range, event merge, status, search, renderer selection | Memory-only events and caches |
| `SettingsPanel` | Personal draft and source discovery | `onSave`, `onReset`, deep links; no persistence writes |
| `AdminSettingsPanel` | Administrator draft, discovery, audiences, ICS catalog | `onSave`; no property writes |
| `PropertyPaneAdminCalendarManager` | Property-pane adapter for administrator panel | Separate Graph client and React mount lifecycle |
| Source services | External access and `ICalendarEvent` mapping | Graph/HTTP requests; no policy or rendering |
| `CalendarSourceRegistry` | Source-type display metadata | Static definitions |
| `sourceIconHelper` | Explicit/registry icon and name fallbacks | Pure derived metadata |
| `CalendarToolbar` | Date, today, and view navigation | Date/view callbacks |
| Day/Week/Month | Prepared-event rendering and local details state | No Graph or persistence |
| `SearchResultsView` | Already-filtered result rendering | No source loading |
| `EventDetailsDialog` | Event details and safe external links | Opens a new browser tab |
| Calendar helpers | Dates, layout, colors, locale, labels, safe links | Pure except `safeOpen` |
| `UserHelper` | Current identity and mailbox working-hours/time-zone data | Graph calls and twelve-hour local cache |
| Localization | English/Dutch resources and declarations | Runtime strings |
| Build configuration | SPFx bundle, hosts, permissions, packaging | Build-time output |

## Flow routing

| Change | Continue with |
| --- | --- |
| Source loads, visible ranges, refresh, deduplication, status, or runtime cache | [Calendar: loading, range, and cache](calendar.md#loading-range-and-cache) |
| Day, Week, Month, Search, toolbar, locale, layout, or event details | [Calendar: rendering and interaction](calendar.md#rendering-and-interaction) |
| Adapter, endpoint, mapping, pagination, or permission | [Sources and permissions](sources-and-permissions.md) |
| Resolution, policy, audience, persistence, migration, or settings UI | [Settings and policy](settings-and-policy.md) |

## Lifecycle facts

- `onThemeChanged` can update CSS variables and effective theme-derived settings without mounting React.
- `onInit` initializes storage and Graph access before the framework `render` mounts React.
- `onDispose` unmounts the React subtree.
- Accepted administrator and personal changes return to the web part before persistence.

## Inactive and legacy components

| Component/data | Classification | Handling |
| --- | --- | --- |
| `ScheduleView.tsx` | Inactive technical debt | Not imported; not supported |
| `AddCalendarDialog.tsx` | Inactive legacy UI | Not imported; obsolete subset of integrated settings flow |
| `ics` source shape | Legacy/inert runtime type | No adapter; coordinator marks ready |
| Unused source-service `HttpClient` fields | Technical debt | Do not infer a desired dual-client architecture |
| PnP Calendar history/comments | Historical | No active renderer dependency |
