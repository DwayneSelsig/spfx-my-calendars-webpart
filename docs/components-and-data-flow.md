# Components and Data Flow

**Read when:** changing component ownership, SPFx lifecycle, composition boundaries, or hand-offs between the web part, settings, sources, coordinator, and renderers.

This document describes verified current responsibilities. Normative cross-cutting boundaries are in [Architecture and behavior](architecture-and-behavior.md).

## Component map

| Component | Responsibility | State or side effect |
| --- | --- | --- |
| `MyCalendarsWebPart` | SPFx composition root; Graph scheduling, settings, audiences, theme and React lifecycle | Web-part properties, shared Graph executor, persistence callbacks and effective settings |
| `ICalendarSettings.ts` | Persisted, effective, source, and default data definitions | Static types/defaults; no precedence logic |
| `CalendarSettingsService` | Validation, schema 7 migration/reference hydration, policy resolution, explicit visibility and permitted personal overrides | Pure except identifier generation |
| `SettingsStorageService` | Current/legacy personal JSON in OneDrive App Folder | Graph I/O; no policy resolution |
| `AudienceService` | Typed group discovery and current-user membership | Five-minute session cache; no policy effects |
| `GraphRequestScheduler` and Graph executor | Order Graph HTTP operations by priority and capacity. Execute continuation pages through the same scheduler | One instance per web part |
| `CalendarRangeCoordinator` | Discovery, source/month work, priority, freshness and deduplication | Runtime state and optional IndexedDB appointment cache |
| `MyCalendars` | Feed the visible range to the coordinator, present source status, and render search and calendar views | React state and renderer selection |
| `UserSettingsPanel` | Personal draft, live preview and source discovery | Awaited `onSave`/`onReset`, preview/cancel, deep links; no persistence writes |
| `AdminDefaultsPanel` | Administrator draft, group-first calendar discovery, mailbox rules, policy and shared presentation, ICS catalog | `onSave`; no property writes |
| `PropertyPaneAdminDefaultsManager` | Property-pane adapter for administrator panel | Required SPFx change bridge, current-value getters, shared Graph executor and React mount lifecycle |
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
- `onInit` applies the persisted Graph concurrency limit before Graph settings I/O. React mounts after settings resolution and starts calendar jobs without awaiting them.
- The web part creates one React 18 root on its first render and reuses it for subsequent renders. `onDispose` unmounts and clears that root.
- The administrator property-pane adapter owns one React 18 root per control container, reuses it across renders and SPFx field-descriptor replacements, and unmounts and removes it on disposal (DEC-031).
- The webpart owns accepted settings callbacks. Administrator saves use one SPFx callback writer and one property; personal writes use OneDrive. Panels await completion before closing; the adapter does not refresh itself during a save.
- The web part discovers applicable automatic Exchange mailbox rules using the Exchange service before pure settings resolution, on initialization, administrator changes and manual refresh. Generation checks discard obsolete results.
- AssignmentControls supplies reusable policy, capability and grouping presentation; CalendarSettingsService retains all policy decisions.

## Inactive and legacy components

| Component/data | Classification | Handling |
| --- | --- | --- |
| `ics` source shape | Legacy/inert runtime type | No adapter; coordinator marks ready |
| PnP Calendar history/comments | Historical | No active renderer dependency |

## Settings flows

```text
UserSettingsPanel -> MyCalendars -> MyCalendarsWebPart
  preview -> derive/resolve -> runtime only
  cancel  -> discard preview -> confirmed runtime
  save/reset -> await SettingsStorageService -> accept or rollback

SPFx Property Pane -> PropertyPaneAdminDefaultsManager -> AdminDefaultsPanel
  complete draft -> MyCalendarsWebPart -> SPFx change bridge -> adminSettings
  accepted configuration -> runtime rebuild -> close -> local summary update

Administrator defaults + confirmed personal settings (or preview) + matched audiences
  -> resolveCalendarSettings() -> MyCalendars
```
