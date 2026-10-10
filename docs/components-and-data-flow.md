# Components and Data Flow

**Read when:** changing component ownership, SPFx lifecycle, composition boundaries, or hand-offs between the web part, settings, sources, coordinator, and renderers.

This document describes verified current responsibilities. Normative cross-cutting boundaries are in [Architecture and behavior](architecture-and-behavior.md).

## Component map

| Component | Responsibility | State or side effect |
| --- | --- | --- |
| `MyCalendarsWebPart` | SPFx composition root; settings/audience/theme composition and React lifecycle | Web-part properties, OneDrive callbacks, effective settings |
| `ICalendarSettings.ts` | Persisted, effective, source, and default data definitions | Static types/defaults; no precedence logic |
| `CalendarSettingsService` | Validation, schema 7 migration/reference hydration, policy resolution, explicit visibility and permitted personal overrides | Pure except identifier generation |
| `SettingsStorageService` | Current/legacy personal JSON in OneDrive App Folder | Graph I/O; no policy resolution |
| `AudienceService` | Typed group discovery and current-user membership | Five-minute session cache; no policy effects |
| `MyCalendars` | Source coordination, visible range, event merge, status, search, renderer selection | Runtime state plus optional administrator-controlled browser cache |
| `SettingsPanel` | Personal draft and source discovery | `onSave`, `onReset`, deep links; no persistence writes |
| `AdminSettingsPanel` | Administrator draft, group-first calendar discovery, mailbox rules, policy and shared presentation, ICS catalog | `onSave`; no property writes |
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
- Accepted administrator and personal changes return to the web part before persistence. Administrator JSON is handed to SPFx callbacks before local property synchronization; runtime acceptance follows that hand-off. Older administrator save completions cannot refresh the property pane over a newer save.
- The web part discovers applicable automatic Exchange mailbox rules using the Exchange service before pure settings resolution, on initialization, administrator changes and manual refresh. Generation checks discard obsolete results.
- AssignmentControls supplies reusable policy, capability and grouping presentation; CalendarSettingsService retains all policy decisions.

## Inactive and legacy components

| Component/data | Classification | Handling |
| --- | --- | --- |
| `ics` source shape | Legacy/inert runtime type | No adapter; coordinator marks ready |
| PnP Calendar history/comments | Historical | No active renderer dependency |
