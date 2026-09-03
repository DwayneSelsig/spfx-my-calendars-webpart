# Sources and Permissions

Read the [common contract](#common-contract-permissions-and-resilience) and only the source section selected by [AGENTS.md](../AGENTS.md). Requirements are normative; current endpoint, mapping, cache, and failure descriptions are verified implementation facts. Related intent and known conflicts are recorded by ID in [Decisions, deviations, and open questions](decisions-deviations-and-open-questions.md).

## Capability status

| Source/capability | Status | Runtime events | Configuration |
| --- | --- | --- | --- |
| Current-user Exchange calendars | Active | Yes | Automatic; individual calendars can be hidden personally |
| Configured Exchange/shared mailbox calendar | Active | Yes | Administrator or personal source |
| SharePoint list calendar | Active | Yes | Administrator or personal source with optional field mapping |
| Planner plan | Active | Yes | Automatic accessible plans or configured plan |
| Microsoft 365 Group/Team calendar | Active | Yes | Automatic joined groups or configured group |
| Teams Shifts | Active | Yes | All shifts from all directly joined Teams |
| ICS subscription | Active action, not an event source | No | Personal deep-link flow or administrator audience catalog |
| Legacy `ics` source record | Inert/legacy | No | Can survive normalization/migration; marked ready by coordinator |
| Schedule view | Inactive | N/A | Source-independent inactive renderer |

All active adapters are read-only. They return normalized local events or discovery values and never create, update, or delete external source data.

## Common contract, permissions, and resilience

**Read when:** changing the shared event model, source coordination boundaries, Graph permission manifests, authentication context, pagination, retry, throttling, or cross-source verification. Related records: DEC-001, DEC-004, DEC-008, DEV-005, DEV-008, DEBT-003, and OQ-006.

### Shared adapter contract

- Every active adapter **MUST** produce `ICalendarEvent` values that satisfy the [common event contract](architecture-and-behavior.md#common-event-contract).
- Source-specific API access and normalization **MUST** remain in the source service.
- A source service **MUST NOT** decide administrator or personal policy, persist settings, or render events.
- An adapter **MUST** reject an item when it cannot produce valid required dates and **MUST NOT** replace an invalid source date with the current time.
- Source failures **MUST** remain visible to the coordinator and **MUST NOT** cancel independent source families.

The coordinator adds or normalizes source identity, display name, color, type, icon, and search text before rendering. Events are merged by `sourceId:event.id`.

### Delegated access

All current runtime calls use delegated user context because the web part obtains `MSGraphClientV3` from SPFx and uses `/me` endpoints. No application-only code path exists.

### Microsoft Graph permissions

| Requested permission | Current code paths | Required by current path | Notes |
| --- | --- | --- | --- |
| `Calendars.Read` | User/shared Exchange discovery and `calendarView`; `/me/calendar/getSchedule` | Yes | Covers event details and exceeds the documented `Calendars.ReadBasic` minimum for `getSchedule` |
| `Calendars.Read.Shared` | `/users/{mailbox}/calendars` and shared/delegated mailbox calendar views | Yes for configured shared/delegated calendars | Does not grant access beyond the signed-in user's source permissions |
| `MailboxSettings.Read` | No `/me/mailboxSettings` path | No verified consumer | Requested in both manifests; `UserHelper` gets working hours/time zone through `getSchedule` |
| `Files.ReadWrite.AppFolder` | Create/read/write/delete under `/me/drive/special/approot` | Yes | Stores personal settings only; Microsoft labels delegated AppFolder permission as preview |
| `Sites.Read.All` | Site/list/column/item discovery and SharePoint event retrieval | Yes | Broad delegated read access to site collections available to the user |
| `Tasks.Read` | Planner plan/task access | Yes | Current code does not write tasks |
| `Group.Read.All` | Unified-group calendar/discovery, Planner group discovery, administrator audience discovery | Yes for the current aggregate set | `/me/checkMemberGroups` can use lower permissions, but `/groups` paths require group read access |
| `Team.ReadBasic.All` | `/me/joinedTeams` for Shifts discovery and Team-versus-Group icons | Yes | Returns directly joined Teams; shared-channel host Teams have documented limitations |
| `Schedule.Read.All` | `/teams/{teamId}/schedule/shifts` | Yes | Current behavior reads all returned shifts for joined Teams |

The requests are duplicated in `config/package-solution.json` and `MyCalendarsWebPart.manifest.json`. Both lists **MUST** remain aligned when permissions change.

### Pagination, retry, and throttling

The code follows `@odata.nextLink` for Unified-group discovery, joined-Team discovery, and Shifts retrieval. It does not follow paging for Exchange calendars/events, SharePoint items, Planner group/plan/task discovery, or administrator audience search.

There is no explicit exponential backoff, `Retry-After` handling, or retry limit in a source service. Retry occurs indirectly when a failed month remains uncached and later navigation or manual refresh starts another load. Selected failed discovery promises are cleared so a later load can rediscover.

A retry policy, mandatory-source retry policy, and error-record schema remain unresolved and **MUST NOT** be invented as part of an unrelated source change.

### Verification references

- [List calendars](https://learn.microsoft.com/graph/api/user-list-calendars)
- [Get schedule](https://learn.microsoft.com/graph/api/calendar-getschedule)
- [OneDrive App Folder](https://learn.microsoft.com/graph/onedrive-sharepoint-appfolder)
- [List SharePoint list items](https://learn.microsoft.com/graph/api/listitem-list)
- [List Planner tasks](https://learn.microsoft.com/graph/api/plannerplan-list-tasks)
- [Check signed-in-user group membership](https://learn.microsoft.com/graph/api/directoryobject-checkmembergroups)
- [List joined Teams](https://learn.microsoft.com/graph/api/user-list-joinedteams)
- [List Shifts](https://learn.microsoft.com/graph/api/schedule-list-shifts)

### Missing automated coverage

There are no automated tests for endpoint construction, pagination, source mappings, invalid-date handling, permission alignment, ICS link generation, source identity, deduplication, status, or retry eligibility. `npm run build` remains the production verification command.

## Exchange calendars

**Read when:** changing current-user calendars, shared/configured mailboxes, Outlook calendar discovery, `calendarView`, Exchange mapping, or Exchange links. Related records: DEC-004, DEC-008, DEV-007, and DEV-008.

### Discovery and retrieval

- Current-user calendars are always discovered from `GET /me/calendars`.
- A configured mailbox uses `GET /users/{mailbox}/calendars`.
- Personal UI loads current-user calendars automatically and can discover a manually entered mailbox.
- Administrator and personal source flows can select one calendar from a mailbox.

Events use `GET /me/calendars/{calendarId}/calendarView` or `GET /users/{mailbox}/calendars/{calendarId}/calendarView` with an inclusive/exclusive date window, `Prefer: outlook.timezone="UTC"`, selected fields, and `$top=500`.

### Mapping contract

The adapter maps ID, subject, preview, start/end, all-day state, location, organizer, attendees, online meeting, join URL, and web link. The coordinator supplies source display name, source ID, Outlook color, source type, and logo setting.

Missing or invalid required timed dates fail that calendar request. Date-only/all-day values are converted through local midnight before ISO serialization. The current user email determines `isOrganizer`; failure to obtain it produces `false` without failing retrieval.

### Settings, links, cache, and failures

- `exchangeCalendarStates` hides automatic current-user calendars; absent IDs are enabled.
- Configured sources use mailbox, calendar ID, name, color, enabled state, and optional per-source logo.
- Event details can open Graph-provided `joinUrl` and `webLink`.
- Calendar discovery is cached as a coordinator promise until reset; events use the successful source/month cache.
- Individual automatic calendars and configured Exchange sources are isolated from each other.
- Calendar and event responses do not follow `@odata.nextLink`; discovery and retrieval can be incomplete beyond the returned page or 500 items.

**Registered deviation:** mailbox existence is checked with SPFx `HttpClient` against `https://graph.microsoft.com/v1.0/users/...`, not the authenticated Graph client. A false response blocks mailbox selection even though the later Graph calendar call is authoritative. Do not silently treat this validation path as desired architecture.

## SharePoint list calendars

**Read when:** changing SharePoint site/list discovery, column inspection, field mapping, list-item retrieval, or SharePoint event conversion. Related records: DEC-004, DEC-008, DEV-004, and DEV-008.

### Discovery

- Site discovery uses `GET /sites?search=*` or a search term, with `$top=999`.
- List discovery uses `GET /sites/{siteId}/lists` with `$top=200` and accepts visible `events` and `genericList` templates.
- Settings panels read list columns and can inspect the first item to propose a field mapping.

### Retrieval and mapping

Runtime retrieval calls `GET /sites/{siteId}/lists/{listId}/items?expand=fields`, then filters the requested range on the client.

Configured or detected columns map title, start, end, description, location, and all-day state. Default names are `Title`, `EventDate`, `EndDate`, `Description`, `Location`, and `fAllDayEvent`.

Items without a title or start date, items with invalid dates, and items outside the range are skipped. A missing end uses the start value.

### Settings, cache, and failures

- Source identity is site ID plus list ID.
- Name, color, enabled state, IDs, and field mapping are stored in the source entry.
- Runtime currently uses the source-type-wide logo flag, not per-source `showSourceLogo`.
- Configured sources load independently and successful months are cached in memory.
- Site search falls back to wildcard accessible-site discovery after an error; several discovery methods return an empty array on failure.
- Runtime list-item retrieval propagates errors to the coordinator.
- Runtime retrieval does not follow paging and does not send a server-side date filter. Large lists can be incomplete or expensive.

## Planner

**Read when:** changing Planner group/plan discovery, task retrieval, automatic plan mode, assignment/completion filters, or Planner date mapping. Related records: DEC-004, DEC-008, DEV-006, and DEV-008.

Automatic and configured Planner sources expose task dates as all-day calendar events.

### Discovery and retrieval

Plan discovery:

1. reads up to 100 groups from `GET /me/memberOf/microsoft.graph.group`;
2. calls `GET /groups/{groupId}/planner/plans` sequentially for each returned group; and
3. suppresses errors for an individual group, treating it as having no Planner.

Configured plan access can be checked with `GET /planner/plans/{planId}`. Runtime tasks use `GET /planner/plans/{planId}/tasks`.

### Mapping and filters

- `plannerAssignedToMeOnly` filters the task `assignments` map against the current user ID.
- `showCompletedTasks=false` excludes tasks with `percentComplete >= 100`.
- A task with neither start nor due date is skipped.
- A single date becomes both start and end; reversed start/end dates are swapped.
- Tasks are all-day events and can include progress/checklist text and `percentComplete`.
- No Planner task deep link is mapped.

**Deviation:** invalid non-empty Planner date strings are not explicitly rejected before `toISOString`; they can throw and fail that plan load rather than skip only the invalid task.

### Automatic mode, cache, and failures

- `plannerShowAllCalendars=true` creates one runtime source for every discovered plan.
- `plannerShowAllAssignedToMeOnly` applies to generated sources.
- Automatic mode always includes completed tasks.
- Configured Planner sources are not loaded while automatic mode is enabled.
- Plan discovery is cached until reset. Per-plan task loads are parallel and isolated; successful months are cached.
- Group discovery has `$top=100` without paging. Plan/task calls do not follow paging.
- If current-user ID lookup fails while assigned-only mode is requested, filtering is not applied and all returned tasks remain visible.

## Microsoft 365 Group and Teams calendars

**Read when:** changing Unified-group discovery, joined-Team icon discovery, group `calendarView`, automatic group mode, or Group/Team event mapping. Related records: DEC-004, DEC-008, DEV-004, and DEV-008.

### Discovery

- `GET /me/memberOf/microsoft.graph.group` is paged and filtered client-side to groups whose `groupTypes` contains `Unified`.
- `GET /me/joinedTeams?$select=id` is paged separately to choose a Teams or Group icon.
- A Team calendar and its backing Microsoft 365 Group calendar use the same group calendar endpoint.

### Retrieval and mapping

`GET /groups/{groupId}/calendarView` uses the requested range, UTC preference, selected event fields, and `$top=500`.

Mapping matches Exchange: title, preview, dates, all-day state, location, organizer, attendees, online meeting, join URL, and web link. Missing or invalid required dates fail that group request. The coordinator supplies group/source identity, display name, color, logo, and a Team-versus-Group icon.

### Automatic mode, cache, and failures

- `unifiedGroupShowAllCalendars=true` loads all discovered Unified groups.
- Explicit group sources are not loaded while automatic mode is enabled.
- Discovery promises and successful source/month results are cached until reset.
- Event loads are isolated per group; discovery failure affects the complete service family.
- Event retrieval does not follow paging beyond `$top=500`.
- Runtime uses the source-type logo setting and ignores per-source `showSourceLogo`.

## Teams Shifts

**Read when:** changing joined-Team discovery, schedule/shifts retrieval, draft/shared shift mapping, Shifts range filtering, or Team-level failure behavior. Related records: DEC-004, DEC-008, DEC-016, DEV-004, DEV-008, and DEV-009.

### Product scope and discovery

Teams Shifts **MUST** include all shifts returned for all Teams where the current user is a direct member. It is not limited to shifts assigned to the current user.

`GET /me/joinedTeams` is paged. Each Team is queried sequentially through `GET /teams/{teamId}/schedule/shifts`; no server-side date filter is sent. Returned shifts are filtered for overlap on the client.

### Mapping

- A shift with a draft representation emits the draft only; otherwise it emits the shared representation.
- Missing start or end values are skipped.
- Start/end are swapped when reversed.
- Duration of at least 12 hours is classified as all-day.
- Team name becomes title and location.
- Notes and activity names form the description.
- Shift theme can override the configured source color.
- Draft events set `isDraft`; renderers display draft text in italics.
- No Shifts deep link is mapped.

### Settings, cache, and failures

- `teamsShiftsShowAllCalendars=true` creates one automatic logical source covering all joined Teams.
- A configured Shifts source also covers all joined Teams; it does not identify one Team.
- Automatic mode suppresses configured Shifts sources.
- Joined-Team discovery is cached in the retained service instance until reset. Successful months use the coordinator cache.
- Team-not-found/404 is treated as an empty Team. Another Team failure aborts the full source call, so results collected earlier in that call are lost.
- The API supports server-side date filtering, but current code downloads every returned page for each Team before client filtering.
- Runtime uses the source-type logo flag and ignores per-source `showSourceLogo`.

## ICS subscription

**Read when:** changing ICS URL handling, Outlook subscription links, administrator ICS catalog entries, or legacy ICS source records. Related records: DEC-002, DEV-012, and DEBT-002. Read [Settings: audiences](settings-and-policy.md#audiences) when catalog targeting changes.

ICS is an active subscription action, not an event source.

- A user enters or selects a display name and ICS URL.
- The UI opens `https://outlook.office.com/calendar/addcalendar?url={encodedUrl}&name={encodedName}` in a new tab.
- An administrator can publish named ICS catalog entries to audiences; selecting one pre-fills the same personal flow.
- The web part **MUST NOT** fetch, parse, normalize, cache, or render ICS content.
- The subscription action is not persisted as a personal web-part source.
- An administrator catalog item is persisted in web-part properties but does not prove that a user completed the Outlook subscription.

### Legacy records

Legacy `ics` source records can be normalized or migrated. The coordinator registers them, performs no retrieval, and marks the ICS service ready. They are inert and **MUST NOT** be documented as rendered calendars.

`CalendarSourceRegistry` and inactive `AddCalendarDialog` contain obsolete wording about adding or pasting ICS content. That wording is a registered deviation and does not override this contract.
