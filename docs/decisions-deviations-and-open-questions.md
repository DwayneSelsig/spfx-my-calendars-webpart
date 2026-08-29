# Decisions, Deviations, and Open Questions

This register is descriptive. It is not a normative specification. The [architecture and behavior contract](architecture-and-behavior.md) contains normative requirements.

## Status labels

| Label | Meaning |
| --- | --- |
| Decision | Confirmed product or architecture choice. |
| Intention | Confirmed desired behavior that can be incomplete. |
| Technical debt | Current code that is not part of the desired architecture. |
| Deviation | A known difference between current behavior and a confirmed requirement or description. |
| Assumption | A working interpretation that is not confirmed. |
| Open question | A decision that has not been made. |

Future contributors preserve these labels when they update this register. They do not convert an intention, assumption, or open question into a decision without confirmation.

## Product decisions

### DEC-001 — Read-only aggregation

- **Status:** Decision
- **Decision:** My Calendars reads and combines source data. It does not create, change, or delete source events or tasks.
- **Decision:** The application can open a source event, calendar, task, or management flow through a deep link when the source supports this action.
- **Consequence:** Write operations do not belong in source services or renderers.

### DEC-002 — ICS subscription flow

- **Status:** Decision
- **Decision:** ICS support means that the application creates a deep link to the Exchange Online calendar subscription wizard.
- **Decision:** The web part does not parse or render ICS content.
- **Consequence:** The administrator ICS catalog offers approved subscription links to selected audiences. It does not supply runtime events.

### DEC-003 — Local calendar renderer

- **Status:** Decision
- **Decision:** A local, read-only renderer provides Day, Week, and Month.
- **Decision:** Month preserves the established seven-column, six-row calendar behavior. Week uses rolling day cards and a complete 24-hour timeline. Day uses the Microsoft 365 Companion interaction and visual direction.
- **Decision:** The custom Search view remains active.
- **Consequence:** The runtime does not depend on PnP Calendar. Search keeps the hidden active calendar mounted.

### DEC-004 — Independent source loading

- **Status:** Decision
- **Decision:** A source failure does not cancel successful source loads.
- **Decision:** Valid partial results remain visible.
- **Decision:** Source failures have a visible per-source status.

### DEC-005 — Variable administrator source policy

- **Status:** Decision
- **Decision:** Source policy can vary per administrator source.
- **Decision:** A source can be an overridable default, mandatory, or partly restricted.
- **Consequence:** The future settings model needs an explicit source policy. Its shape is not yet defined.

### DEC-006 — Test strategy timing

- **Status:** Decision
- **Decision:** The project will introduce automated tests as one repository-wide initiative.
- **Consequence:** Individual documentation changes do not define a partial future test architecture.

### DEC-007 — Dynamic retrieval range

- **Status:** Decision
- **Decision:** The initial load covers the current month plus or minus three months. Navigation loads missing visible months per source.
- **Decision:** Only successful source/month responses enter the range cache. Manual refresh clears events and range state.
- **Consequence:** Navigation is not limited to the initial retrieval window. Failed sources remain eligible for retry.

### DEC-008 — Source errors and invalid dates

- **Status:** Decision
- **Decision:** Active source retrieval failures remain visible to the coordinator and do not receive a successful range-cache entry.
- **Decision:** A source adapter rejects an event with unusable required dates. It does not invent the current time.
- **Consequence:** A later navigation or manual refresh can retry failed source/month combinations.

### DEC-009 — Display preferences and regional formatting

- **Status:** Decision
- **Decision:** An administrator sets the default weekend visibility. A user can store an explicit personal weekend preference or return to the administrator default.
- **Decision:** Date and time labels use the current SharePoint page culture. The application does not expose a separate 12- or 24-hour clock setting.
- **Consequence:** Personal settings store a weekend value only when the user has selected an override.

## Intentions

### INT-002 — Mandatory-source failure records

- **Status:** Intention
- **Intention:** Store a dated record in the OneDrive App Folder when a Graph error prevents a mandatory source from loading.
- **Intention:** Retry the mandatory source after a delay.
- **Constraint:** Do not store source event content in an error record unless a later decision explicitly permits it.

## Technical debt and deviations

### DEBT-002 — Inactive Schedule view

- **Status:** Technical debt
- **Current state:** The repository contains an inactive Schedule view.
- **Warning:** Its presence does not make Schedule a supported capability.

### DEV-001 — Administrator sources are always overridable

- **Status:** Deviation
- **Current state:** A user can rename, recolor, disable, or remove every audience-assigned administrator source for that user.
- **Desired state:** Enforcement varies by source policy.
- **Missing decision:** The repository has no source-policy schema or migration rule.

### DEBT-004 — No automated project tests

- **Status:** Technical debt
- **Current state:** The repository contains no project test files.
- **Current verification:** `npm run build` performs the production build and package flow.
- **Desired state:** A later repository-wide initiative will define and add tests.

### DEV-004 — Version values differ before release

- **Status:** Deviation
- **Current state:** `package.json`, SharePoint package configuration, and README history can contain different version values.
- **Current release behavior:** The release workflow derives npm and SharePoint versions from the release tag.
- **Assumption:** The release tag is the release version source.

### DEV-005 — Documentation previously overstated capabilities

- **Status:** Deviation
- **Previous state:** The README described direct ICS aggregation and Schedule view support.
- **Current fact:** ICS opens an Exchange Online subscription flow. Schedule is not an active supported view.
- **Resolution:** The README and architecture contract now describe the active behavior.

## Assumptions

### ASM-001 — Primary host

- **Status:** Assumption
- **Assumption:** SharePoint is the primary product context.
- **Evidence:** The project is an SPFx web part and the installation flow uses the SharePoint App Catalog.
- **Limit:** Other hosts are described only when the manifest or runtime code confirms them.

### ASM-002 — Release version source

- **Status:** Assumption
- **Assumption:** A semantic release tag is the release version source.
- **Evidence:** The release workflow derives both package versions from that tag.

## Open questions

### OQ-002 — Administrator source policy schema

- **Status:** Open question
- How will a source declare an overridable, mandatory, or partly restricted policy?
- Which fields can a partly restricted source override?
- How will existing settings migrate when the policy is added?

### OQ-003 — Mandatory-source retry

- **Status:** Open question
- What delay and retry limit will apply?
- What event will start a new retry?
- How will the UI show a persistent mandatory-source failure?

### OQ-004 — OneDrive error records

- **Status:** Open question
- What file and record schema will be used?
- Which error fields are permitted?
- What retention and cleanup rules will apply?
- How will the design prevent sensitive event content from entering the log?

### OQ-005 — Automated tests

- **Status:** Open question
- Which framework and test layers will the repository-wide initiative use?
- Which current pure functions and source mappings will form the first test set?

## Rules for future decisions

Future agents and contributors do not select a source-policy schema, retry policy, log schema, retention rule, or test architecture without a confirmed decision.

When a decision is confirmed:

1. Add or update a decision entry.
2. Update the normative architecture contract when the decision changes a requirement or invariant.
3. Update the README only when the user-visible capability changes.
4. Keep the old deviation until the implementation meets the decision.
