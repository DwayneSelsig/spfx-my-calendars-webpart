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
| Schedule view | Not supported | The repository contains inactive schedule code. This code is not a supported capability. |
| Partial source results | Available | A source failure does not remove events that other sources loaded successfully. |

## Main features

- Combine supported calendar sources in one view.
- Enable or disable individual sources.
- Set a color for each configured source.
- Set a preferred timeline start, visible hours, weekends, and a 15, 30, or 60 minute grid.
- Navigate beyond the initial seven-month window with source- and month-aware loading.
- Search event titles and locations.
- Show source logos by service type.
- Use organization theme colors and light or dark themes.
- Assign administrator sources and ICS catalog entries to Entra groups. The current implementation only discovers non-mail-enabled security groups; the confirmed target model is broader.
- Store personal settings in the OneDrive App Folder.
- Load all accessible Planner, Group and Teams, or Teams Shifts sources automatically. These automatic sources are enabled by default; Planner includes all accessible plans, not only tasks assigned to the current user.
- Open source systems through deep links when the source supports this function.

## Scope and non-goals

My Calendars aggregates read data. Event and task management stays in the source system.

The web part does not:

- create, update, or delete source events;
- create, update, or delete Planner tasks;
- parse or display ICS feeds directly;
- provide a Schedule view.

## Configuration

Web part editors can set administrator defaults in the property pane. Users can set personal preferences in the settings panel.

The administrator settings model includes:

- default calendar view and display values;
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

The repository does not contain automated project tests. `npm run build` is the current production verification command. A repository-wide test strategy is planned for a later change.

## Installation and upgrade

1. Download a release package or create one with `npm run build`.
2. Open the SharePoint App Catalog.
3. Upload the `.sppkg` file.
4. Approve the requested Microsoft Graph permissions.
5. Add the web part to a supported page or host.

For an upgrade, upload the new package and replace the existing package.

## Release versions

Release tags use semantic versioning, for example `v1.2.3`. The release workflow derives the npm version and the four-part SharePoint package version from the tag. Repository files can contain different development-time version values before the workflow runs.

The version history below is descriptive. It is not an architecture or behavior contract.

| Version | Date | Summary |
| --- | --- | --- |
| 0.0.1 | 2026-02-21 | Initial release. |
| 0.0.2 | 2026-02-21 | Added Planner calendars. |
| 0.0.3 | 2026-02-25 | Added Teams Shifts calendars. |
| 0.0.4 | 2026-03-08 | Added Teams and Microsoft 365 Group calendars. |
| 0.0.5 | 2026-03-15 | Added PnP calendar rendering, the common event contract, mailbox time-zone handling, and user caches. |
| 0.0.6 | 2026-03-24 | Improved search and settings performance. |
| 0.0.7 | 2026-03-25 | Added defensive handling for Graph all-day dates. |
| 0.0.8 | 2026-03-28 | Adjusted PnP Calendar menu styles. |
| 0.0.9 | 2026-03-29 | Added automatic source loading options. |
| 0.0.10 | 2026-04-06 | Improved responsive layout. |
| 0.1.0 | 2026-08-23 | Replaced PnP Calendar with own calendar |
| 0.2.0 | 2026-08-29 | Improved locale-aware calendar formatting, updated development dependencies, and clarified rendering documentation. |

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
