# Shared mailbox and audience validation

## Evidence boundaries

Static inspection confirmed incomplete panel error handling, boolean identity validation, stale discovery state, swallowed audience API failures and a 50-result audience limit. Current-user calendar discovery also enclosed configured source loading, so its failure skipped those sources.

Heft/Jest regression tests use Graph mocks and pure helpers. They do not validate Microsoft 365 permissions, Exchange endpoint behavior, rendered panel interactions or SPFx hosts. The matrix below is **not executed in a Microsoft 365 tenant**. Record actual HTTP status, Graph error code, endpoint and granted permissions when running it; do not include sensitive event content.

## Prerequisites

Use an existing authorized test installation. Approve `User.ReadBasic.All`, `Calendars.Read` and `Calendars.Read.Shared` as declared by the solution. Verify the effective user rights in Exchange separately from Graph consent. For Reviewer/LimitedDetails tests use accounts without Full Access, broader folder rights, inherited access or other rights that would invalidate the comparison. Remove/re-add test grants or use separate accounts as needed, and account for permission propagation.

This change does not deploy, publish or merge an installation.

## Manual matrix

| Scenario | Actions | Expected result / evidence to record |
| --- | --- | --- |
| Shared mailbox with Full Access | In admin and personal panels discover by UPN, object ID and a different primary SMTP; choose/save a calendar and reload | Same owner identity/calendar; events load; saved SMTP resolves at runtime; no stuck spinner |
| Reviewer only | Test identity lookup, `/users/{id}/calendars` and the selected calendar's `calendarView` separately | Record which endpoints succeed; readable granted details display if retrieval succeeds; failure is visible and does not claim the mailbox is absent |
| LimitedDetails only | Repeat each endpoint; compare the existing event select with `id,subject,start,end,location,showAs` diagnostically | Record status/code per endpoint and actual returned fields. No assumption that narrow select fixes 404; no full-details promise or production fallback |
| Unknown/invalid input | Try whitespace, missing UPN/object ID/primary SMTP, an alias not equal to UPN/mail, and a valid input surrounded by spaces | Empty/missing input is distinguished from technical failure; valid surrounding spaces are trimmed; old calendars/selection disappear; loading ends |
| Identity success, discovery failure | Use a real resolvable mailbox without adequate calendar rights; test 403/404 where reproducible | Discovery error identifies possible access/identifier issues without calling the mailbox nonexistent; retry is possible |
| Mixed runtime sources | Own calendar plus a failing shared source and another working source; also test failed own discovery alongside saved sources | Successful events remain; Exchange has failed family status; failed source/month remains eligible for retry |
| Audience list/search | More than 50 matching non-mail-enabled security groups; prefix and apostrophe searches; no matches; denied query | All returned pages are available and sorted locally; empty results differ from errors; selected IDs/names survive subsequent searches/errors |
| Obsolete requests | Start slow lookup/search, change input/start a newer request, navigate back, close/reopen or unmount | Old responses do not overwrite current results, selection or loading state |
| Targeting | Publish a test assignment/ICS entry to supported groups in the test environment; view as member and non-member; simulate membership failure | OR semantics apply; non-matches/failures remain fail-closed; targeting grants no source-calendar access |

## Local verification

- `npm run build` passes: 21 Heft/Jest suites, 98 tests, production compilation/lint and solution packaging; zero failed tests and no build warnings.
- Regression coverage adds mailbox identity/discovery/error-stage tests, audience pagination/local sorting/error/membership tests, request-generation tests, selection-preserving state transitions and independent Exchange loading. Existing localization consistency tests also pass.
- `git diff --check` passes. The two permission declarations are aligned at ten scopes, including only the newly added `User.ReadBasic.All`.
- Changed areas are Exchange/Audience services, the two active settings panels, Exchange orchestration in `MyCalendars`, small pure helpers/tests, English/Dutch resources, permission manifests, README and the related normative/decision documents.
- `graphify update .` refreshes the code graph. Its CLI does not refresh semantic documentation extraction; generated graph documentation is not evidence of updated product intent. It reports five configuration JSON files with zero extracted nodes, and updates community names using hubs rather than an LLM.
- No Microsoft 365 tenant/host integration validation, deployment, publication or merge was performed. The existing full event field selection and the LimitedDetails limitation remain.

## References and limitations

- [List users](https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0): delegated basic-profile access; guests cannot use this collection API. Guest mailbox-address resolution is therefore not promised.
- [Graph permissions](https://learn.microsoft.com/en-us/graph/permissions-reference): `User.ReadBasic.All` includes `id`, `userPrincipalName` and `mail`, not `proxyAddresses` alias data.
- [Shared calendar access](https://learn.microsoft.com/en-us/graph/outlook-get-shared-events-calendars): delegated calendar access and owner/recipient calendar-ID contexts.
- [Calendar permission roles](https://learn.microsoft.com/en-us/graph/api/resources/calendarpermission?view=graph-rest-1.0): limitedRead permits availability, titles and locations.
- [Advanced directory queries](https://learn.microsoft.com/en-us/graph/aad-advanced-queries): filter plus orderby requires advanced-query conditions. This implementation leaves orderby absent and sorts after pagination; the original query's runtime failure is not retested here.

Exchange calendar/event pagination, wider audience group types, empty-means-everyone behavior, retry policies and tenant/user audience-cache scoping remain outside this change. Existing deviations remain registered.
