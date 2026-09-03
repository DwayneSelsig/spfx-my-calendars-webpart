# Agent instructions

The key words **MUST**, **MUST NOT**, **SHOULD**, **SHOULD NOT**, and **MAY** have the meanings in [RFC 2119](https://www.rfc-editor.org/rfc/rfc2119) and [RFC 8174](https://www.rfc-editor.org/rfc/rfc8174) only when they appear in all capitals.

## Authority and conflict handling

Use this evidence order:

1. Confirmed requirements in [Architecture and behavior](docs/architecture-and-behavior.md), the task-specific normative document section selected below, and confirmed records in [Decisions, deviations, and open questions](docs/decisions-deviations-and-open-questions.md).
2. Registered deviations and accepted exceptions in that decisions register.
3. Code and tests as evidence of current implementation.

README and Graphify output are navigation or operational evidence. They are not product intent by themselves.

When code conflicts with a confirmed requirement and the conflict is not registered, an agent **MUST** report the conflict. It **MUST NOT** silently change the requirement, treat implementation as desired architecture, or expand the task into a behavior change.

An intention, deviation, technical-debt item, assumption, open question, or inactive component **MUST NOT** be described as desired or implemented architecture.

## Progressive reading

Read the smallest sufficient document set:

1. For a behavior or architecture change, read [Architecture and behavior](docs/architecture-and-behavior.md).
2. Select every matching task row below; follow its anchor directly to the relevant section and do not read unrelated sections.
3. Read only the related IDs in [Decisions, deviations, and open questions](docs/decisions-deviations-and-open-questions.md) needed to establish intent or known conflicts.
4. Read the relevant source files and tests.
5. Read README only for product overview, installation, build, permissions overview, or release work.

| Task | Required task-specific section |
| --- | --- |
| Component ownership, SPFx lifecycle, or composition hand-offs | [Components and data flow](docs/components-and-data-flow.md) |
| Source coordination, loading, visible range, refresh, deduplication, status, partial failure, or runtime cache | [Calendar: loading, range, and cache](docs/calendar.md#loading-range-and-cache) |
| Day, Week, Month, Search, toolbar, layout, event details, or locale formatting | [Calendar: rendering and interaction](docs/calendar.md#rendering-and-interaction) |
| Shared event normalization, permission manifests, pagination, retry, throttling, or cross-source contract | [Sources: common contract, permissions, and resilience](docs/sources-and-permissions.md#common-contract-permissions-and-resilience) |
| Exchange calendars or mailboxes | [Sources: common contract](docs/sources-and-permissions.md#common-contract-permissions-and-resilience) and [Exchange](docs/sources-and-permissions.md#exchange-calendars) |
| SharePoint list calendars | [Sources: common contract](docs/sources-and-permissions.md#common-contract-permissions-and-resilience) and [SharePoint](docs/sources-and-permissions.md#sharepoint-list-calendars) |
| Planner plans or tasks | [Sources: common contract](docs/sources-and-permissions.md#common-contract-permissions-and-resilience) and [Planner](docs/sources-and-permissions.md#planner) |
| Microsoft 365 Group or Team calendars | [Sources: common contract](docs/sources-and-permissions.md#common-contract-permissions-and-resilience) and [Groups and Teams](docs/sources-and-permissions.md#microsoft-365-group-and-teams-calendars) |
| Teams Shifts | [Sources: common contract](docs/sources-and-permissions.md#common-contract-permissions-and-resilience) and [Teams Shifts](docs/sources-and-permissions.md#teams-shifts) |
| ICS subscription or legacy ICS records | [Sources: ICS subscription](docs/sources-and-permissions.md#ics-subscription) |
| Settings defaults, effective precedence, automatic modes, scalar fields, or source-entry fields | [Settings: precedence and fields](docs/settings-and-policy.md#precedence-and-fields) |
| Optional/mandatory membership, locking, allowed overrides, or stale overrides | [Settings: administrator source policy](docs/settings-and-policy.md#administrator-source-policy) |
| Audience targeting, group discovery, membership evaluation, or audience cache | [Settings: audiences](docs/settings-and-policy.md#audiences) |
| Administrator properties, OneDrive storage, save/reset, recovery, or migration | [Settings: storage and migration](docs/settings-and-policy.md#storage-and-migration) |
| Administrator/personal settings panels, property-pane integration, or exposed controls | [Settings: interfaces](docs/settings-and-policy.md#settings-interfaces) |
| Product intent, deviations, unresolved choices, or documentation authority | Read only the relevant IDs in [Decisions, deviations, and open questions](docs/decisions-deviations-and-open-questions.md) |

When a task spans rows, read all matching sections.

Example: a pure calendar-rendering feature reads the core contract, the rendering section in `calendar.md`, the applicable DEC-003/DEC-009 entries, and relevant components/tests. It does not require source or settings sections unless data access or settings behavior also changes.

Documentation-only changes still require a targeted check of relevant source. Code behavior alone does not establish intent.

## Change rules

- Preserve the read-only product boundary. Source adapters and renderers do not create, update, or delete source data.
- Do not implement an unresolved policy schema, retry policy, error-record schema, retention rule, persistence-failure UX, backup strategy, or test architecture without a confirmed decision.
- Keep source-specific API access and normalization in source services.
- Keep settings policy and precedence in the settings service/web-part composition boundary.
- Treat `ScheduleView` and `AddCalendarDialog` as inactive.
- Treat ICS as a subscription deep-link feature, not an event adapter.
- Update the applicable normative document and decision/deviation records when behavior or intent changes.
- Do not remove legacy or inactive code unless the task explicitly requires it.

## Graphify

This project has a knowledge graph at `graphify-out/`. It is a navigation and code-analysis aid, not product authority.

- For a codebase question, first run `graphify query "<question>"` when `graphify-out/graph.json` exists.
- Use `graphify path "<A>" "<B>"` for relationships and `graphify explain "<concept>"` for focused concepts.
- Use `graphify-out/wiki/index.md` for broad navigation when it exists.
- Read `graphify-out/GRAPH_REPORT.md` only for a broad architecture review or when scoped graph commands are insufficient.
- Dirty generated graph files are not a reason to skip Graphify.
- After changing code or repository-visible documentation, run `graphify update .`. Do not edit generated Graphify output manually.
