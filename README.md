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
| Day, Week, and Month views | Temporary implementation | Uses the PnP Calendar control. A replacement is planned but is not selected. |
| Search view | Available | Uses a custom search result view. The PnP calendar stays mounted while search is active. |
| Schedule and inactive custom views | Planned but not decided | The repository contains inactive view code. This code is not a supported capability. |
| Partial source results | Available | A source failure does not remove events that other sources loaded successfully. |

## Main features

- Combine supported calendar sources in one view.
- Enable or disable individual sources.
- Set a color for each configured source.
- Set working hours, weekends, the first day of the week, and slot duration.
- Search event titles and locations.
- Show source logos by service type.
- Use organization theme colors and light or dark themes.
- Assign administrator sources and ICS catalog entries to Entra security groups.
- Store personal settings in the OneDrive App Folder.
- Load all accessible Planner, Group and Teams, or Teams Shifts sources automatically.
- Open source systems through deep links when the source supports this function.

## Scope and non-goals

My Calendars aggregates read data. Event and task management stays in the source system.

The web part does not:

- create, update, or delete source events;
- create, update, or delete Planner tasks;
- parse or display ICS feeds directly;
- provide a supported Schedule view;
- define the future replacement for PnP Calendar.

## Configuration

Web part editors can set administrator defaults in the property pane. Users can set personal preferences in the settings panel.

Administrator settings include:

- default calendar view and display values;
- organization color and source-logo defaults;
- audience-assigned calendar sources;
- audience-assigned ICS catalog entries;
- automatic loading defaults for Planner, Groups and Teams, and Teams Shifts;
- optional PnP telemetry preference.

Personal settings include:

- personal calendar sources;
- source names, colors, and enabled states;
- personal working hours and default view;
- Exchange calendar visibility;
- automatic loading and source-logo overrides.

The current implementation treats all administrator sources as user-overridable defaults. The intended source policy also supports mandatory and partly restricted sources. That policy is not implemented. See [Decisions, deviations, and open questions](docs/decisions-deviations-and-open-questions.md).

## Documentation

Read these documents before you change behavior:

1. [Architecture and behavior contract](docs/architecture-and-behavior.md)
2. [Decisions, deviations, and open questions](docs/decisions-deviations-and-open-questions.md)

The architecture contract contains the normative product boundaries and invariants. The register separates confirmed decisions from intentions, technical debt, and unresolved questions.

## Requirements

- A SharePoint or Microsoft 365 tenant that supports the configured host.
- The required Microsoft Graph permissions.
- Access to each configured calendar source.
- Node.js `>=18.17.1 <19.0.0` or `>=22.14.0 <23.0.0` for local development.

SharePoint is the primary product context. The manifest also lists Teams personal app, Teams tab, and SharePoint full-page hosts.

## Microsoft Graph permissions

The solution requests these delegated Microsoft Graph permissions:

| Permission | Purpose |
| --- | --- |
| `Calendars.Read` | Read the current user's calendars and events. |
| `Calendars.Read.Shared` | Read calendars that other users shared with the current user. |
| `MailboxSettings.Read` | Resolve mailbox time-zone and working-hours data. |
| `Files.ReadWrite.AppFolder` | Store and load personal settings in the OneDrive App Folder. |
| `Sites.Read.All` | Discover SharePoint sites and read compatible lists. |
| `Tasks.Read` | Read Planner plans and tasks. |
| `Group.Read.All` | Discover groups and read Group or Team calendar events. |
| `Team.ReadBasic.All` | Discover joined Teams and select the correct source icon. |
| `Schedule.Read.All` | Read Teams Shifts data. |

A tenant administrator must approve permissions that require administrator consent in the SharePoint API access page.

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
- [PnP SPFx React Controls Calendar](https://pnp.github.io/sp-dev-fx-controls-react/controls/Calendar/)
- [Heft documentation](https://heft.rushstack.io/)

The icon is from Microsoft Fluent UI System Icons and uses the MIT License.

## Disclaimer

THIS CODE IS PROVIDED AS IS, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED.
