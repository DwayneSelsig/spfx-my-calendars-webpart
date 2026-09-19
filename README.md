# <picture><img src="docs/images/icon.svg" alt="My Calendars icon" width="32" height="32" /></picture> My Calendars

My Calendars is a read-only SharePoint Framework web part. It combines calendar data from multiple Microsoft 365 and SharePoint sources. It presents the data in one calendar interface.

The web part does not create, change, or delete events or tasks in a source system. When a source provides a suitable URL, the web part can link the user to the source calendar or event.

![My Calendars web part in SharePoint](docs/images/my-calendars-screenshot.png "My Calendars web part in SharePoint")

## Capability status

| Capability | Status | Notes |
| --- | --- | --- |
| Exchange calendars | Available | Includes the current user's calendars and configured shared mailboxes. |
| SharePoint list calendars | Available | Supports explicit field mapping for compatible lists. |
| Planner tasks | Available | Supports assignment and completion filters. |
| Microsoft 365 Group and Teams calendars | Available | Uses one calendar source per selected group or Team. |
| Teams Shifts | Available | Includes shared and draft shifts. |
| ICS subscription | Available | Opens the Exchange Online subscription wizard with a generated deep link. The web part does not parse ICS data. |
| Day, Week, and Month views | Available | Uses the local read-only calendar renderer. |
| Search view | Available | Uses a custom search result view. The active calendar stays mounted while search is active. |
| Schedule view | Not supported | Day, Week, and Month are the supported calendar views. |
| Partial source results | Available | A source failure does not remove events that other sources loaded successfully. |

## Main features

- Combine Exchange, SharePoint list, Planner, Microsoft 365 Group and Teams, and Teams Shifts data through a common event contract in one view.
- Use the responsive local calendar renderer in Day, Week, and Month views, with calendar navigation and styling designed for the web part rather than an external calendar component.
- Use optimized search and settings flows. Search covers event titles and locations while the active calendar remains mounted.
- Navigate beyond the initial seven-month window with source- and month-aware loading, and optionally cache the initial range in the browser for an administrator-defined lifetime. An unreadable or incompatible appointment-cache entry is discarded as a whole and rebuilt from source data.
- Handle source-specific date and time behavior defensively, including mailbox time zones, Graph all-day values, and SharePoint's inclusive all-day end dates.
- Show sanitized HTML descriptions and meeting organizer information without presenting an attendee-less appointment as a meeting.
- Enable or disable individual sources. Outlook and SharePoint sections also provide a tri-state bulk visibility control that changes those same individual visibility settings.
- Keep automatic Planner, Group and Teams, and Teams Shifts discovery separate from visibility controls. Automatically discovered sources are enabled by default; Planner includes all accessible plans, not only tasks assigned to the current user.
- Show the SharePoint site name separately from the editable calendar name in personal settings, administrator settings, and event details. Older saved sources are enriched when their site can be resolved.
- Display Graph calendar availability in Day, Week, Month, Search, and Event Details while preserving the calendar color as source identity. Availability uses icons, outlines, and theme-aware patterns; the signed-in user's response is shown only for events loaded from their own mailbox.
- Keep declined personal events visible with a response icon, muted presentation, and struck-through title.
- Set a color for each configured source and optionally show source logos by service type.
- Set a preferred timeline start, visible hours, weekends, and a 15, 30, or 60 minute grid.
- Follow the SharePoint page culture for locale-aware date and time formatting, including regional 12- or 24-hour clocks.
- Use organization theme colors and support light and dark themes.
- Assign administrator sources and ICS catalog entries to Entra groups. The current implementation only discovers non-mail-enabled security groups; the confirmed target model is broader.
- Store personal settings and per-source visibility choices in the OneDrive App Folder.
- Preserve partial results when one source fails and use defensive Graph mapping for incomplete or future values.
- Open exact Outlook and Microsoft 365 Group calendar events, SharePoint list events, and Planner tasks from Event Details. Sources without a reliable exact item link, including Teams Shifts, do not receive a general application fallback.

## Scope and non-goals

My Calendars aggregates read data. Event and task management stays in the source system.

The web part does not:

- create, update, or delete source events;
- create, update, or delete Planner tasks;
- parse or display ICS feeds directly;
- provide a Schedule view.

## Installation
Go to the [SharePoint admin center → **More features**](https://go.microsoft.com/fwlink/?linkid=2185077) → **Apps** → **Open** → **Upload** the `.sppkg` file. Approve Microsoft Graph permissions when prompted.

### Upgrades
Upload the new `.sppkg` file and overwrite the existing one when prompted.

> **Note:** SharePoint add-ins are being retired, but SharePoint Framework (SPFx) solutions like this one are not affected and remain fully supported.

For more information, see the SharePoint App Catalog documentation:
https://learn.microsoft.com/sharepoint/use-app-catalog

## Configuration

Web part editors can set administrator defaults in the property pane. Users can set personal preferences in the settings panel.

The administrator settings model includes:

- default calendar view and display values;
- browser-cache enablement and a cache duration from 1 to 60 minutes;
- organization color and source-logo defaults;
- audience-assigned calendar sources;
- audience-assigned ICS catalog entries;
- automatic loading defaults for Planner, Groups and Teams, and Teams Shifts;

The current administrator panel does not expose every modeled field. Organization color, service-level source-logo defaults, automatic source-loading defaults, and the Planner assigned-to-me default are registered implementation gaps.

Personal settings include:

- personal calendar sources;
- source names, colors, and enabled states;
- personal weekend, start-time, and visible-hour overrides and a default view;
- Exchange calendar visibility;
- automatic loading and source-logo overrides.

Date and time labels follow the current SharePoint page culture. This includes the regional 12- or 24-hour time format.

The confirmed source policy separates membership (`optional` or `mandatory`) from the set of user overrides an administrator allows. Mandatory sources cannot be disabled or removed. The current implementation has no persisted policy schema and treats every administrator source as user-overridable. See [Settings and policy](docs/settings-and-policy.md#administrator-source-policy) and [DEC-005](docs/decisions-deviations-and-open-questions.md#dec-005--administrator-source-policy-dimensions).

## Documentation

The repository uses progressive reading so a change only loads relevant instructions:

1. [Architecture and behavior](docs/architecture-and-behavior.md) is the small cross-cutting contract.
2. [AGENTS.md](AGENTS.md) routes a task directly to the relevant section in the calendar, source, or settings document.
3. [Decisions, deviations, and open questions](docs/decisions-deviations-and-open-questions.md) is searched by the related record IDs; it is not mandatory cover-to-cover reading.

Use [Calendar](docs/calendar.md) for loading and rendering behavior, [Sources and permissions](docs/sources-and-permissions.md) for adapters, APIs, mappings, pagination, and permissions, and [Settings and policy](docs/settings-and-policy.md) for precedence, policy, audiences, storage, migration, and settings interfaces.

## Requirements

- A SharePoint or Microsoft 365 tenant that supports the configured host.
- The required Microsoft Graph permissions.
- Access to each configured calendar source.
- Node.js `>=22.14.0 <23.0.0` for local development.

The supported host contract covers SharePoint web-part pages, SharePoint full-page apps, Microsoft Teams personal apps and tabs, Microsoft 365, and Outlook. Some host identities are recognized through the Teams SDK rather than listed as distinct manifest values. Host-sensitive behavior still requires validation in each host.

## Microsoft Graph permissions

The solution requests these delegated Microsoft Graph permissions:

| Permission | Purpose |
| --- | --- |
| `Calendars.Read` | Read the current user's calendars and events. |
| `Calendars.Read.Shared` | Read calendars that other users shared with the current user. |
| `MailboxSettings.Read` | Requested in both manifests, but no current runtime path reads mailbox settings; this scope is registered as redundant pending removal or a confirmed consumer. |
| `Files.ReadWrite.AppFolder` | Store and load personal settings in the OneDrive App Folder. |
| `Sites.Read.All` | Discover SharePoint sites and read compatible lists. |
| `Tasks.Read` | Read Planner plans and tasks. |
| `Group.Read.All` | Discover groups and read Group or Team calendar events. |
| `Team.ReadBasic.All` | Discover joined Teams and select the correct source icon. |
| `Schedule.Read.All` | Read Teams Shifts data. |

A tenant administrator must approve permissions that require administrator consent in the SharePoint API access page.

See [Sources and permissions](docs/sources-and-permissions.md#common-contract-permissions-and-resilience) for endpoint use, least-privilege notes, pagination limits, and authoritative Microsoft documentation links.

## Build and local development

Install dependencies and start the local workbench:

```text
npm install
npm start
```

Create a production package:

```text
npm run build
```

The package is written to `sharepoint/solution/`.

`npm run build` runs the focused Jest regression tests, including browser-cache and administrator-settings persistence coverage, before creating the production package.

## Installation and upgrade

1. Download a release package or create one with `npm run build`.
2. Open the SharePoint App Catalog.
3. Upload the `.sppkg` file.
4. Approve the requested Microsoft Graph permissions.
5. Add the web part to a supported page or host.

For an upgrade, upload the new package and replace the existing package.

## Contributing

- Report a defect with reproduction steps and environment details.
- Propose a feature with its user goal and source-system constraints.
- Keep user-visible text and documentation in English unless a localization file supplies the translation.
- Do not treat inactive code as a confirmed product requirement.

## References

- [SharePoint Framework setup](https://learn.microsoft.com/sharepoint/dev/spfx/set-up-your-development-environment)
- [SharePoint App Catalog](https://learn.microsoft.com/sharepoint/use-app-catalog)
- [Microsoft Graph Calendar API](https://learn.microsoft.com/graph/api/resources/calendar)
- [Microsoft Graph Planner API](https://learn.microsoft.com/graph/api/resources/planner-overview)
- [Heft documentation](https://heft.rushstack.io/)

The icon is from Microsoft Fluent UI System Icons and uses the MIT License.

## Disclaimer

THIS CODE IS PROVIDED AS IS, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED.
