# Dependency compatibility

This is an operational dependency review, not a product behavior contract.

## SPFx constraints

The solution targets **SPFx 1.24.1**, with **React and React DOM exactly 18.3.1** and **TypeScript 5.8.3**. Supported development runtimes are Node.js 22 (22.14.0 or later within 22.x), 24, and 26. Node.js 26 LTS is recommended for new environments. The release workflow uses Node 26.

See Microsoft's [1.24.1 release notes](https://learn.microsoft.com/sharepoint/dev/spfx/release-1.24.1), [compatibility reference](https://learn.microsoft.com/sharepoint/dev/spfx/compatibility), and [React 18 migration guide](https://learn.microsoft.com/sharepoint/dev/spfx/react-18-upgrade).

## Selected dependencies and configuration

All eleven direct SPFx runtime, lint, module-interface, Heft-plugin and build-rig packages use exactly `1.24.1`. React and React DOM use `18.3.1`; their type definitions use compatible `^18.3.0` ranges, resolved in the lockfile to `18.3.31` and `18.3.7`. The Yarn React type resolution follows the same range; npm uses the lockfile.

The published `@microsoft/generator-sharepoint@1.24.1` templates and dependency definitions were inspected separately from this solution. The project adopts the template's Node engine range, exact Heft `1.3.2`, TypeScript watch exclusions for generated Teams/Copilot outputs, and npm overrides:

| Package override | Template range |
| --- | --- |
| `qs` | `>=6.15.2` |
| `serialize-javascript` | `^7.1.3` |
| `uuid` | `>=11.1.1` |

The existing Heft scripts, rig-based TypeScript configuration, flat ESLint profile with project rules, source permissions, host declarations, solution ID and component ID are preserved. Existing compatible third-party dependencies remain in place, including Fluent UI v8; React 18 does not require Fluent UI v9. TypeScript remains within the supported 5.8 line.

The generator metadata now records `1.24.1`. The deployable solution version is incremented to `1.0.1.0` so the App Catalog can recognize the upgrade.

## React lifecycle

The web part retains one `createRoot` root for its DOM element and unmounts it in `onDispose`. The property-pane adapter retains a root per control container and clears it on disposal. Rerenders preserve the mounted tree and its settings draft. See [DEC-031](decisions-deviations-and-open-questions.md#dec-031--spfx-1241-and-react-18-lifecycle).

DOM regression tests use React 18 roots and `act` from React, with the test-only `IS_REACT_ACT_ENVIRONMENT` flag in the Jest configuration. Existing panel and adapter Save/Reset, failure, rollback and pending-operation regressions exercise the upgraded rendering behavior.

## Release workflow

The release workflow `.github/workflows/main.yml` uses `ubuntu-latest`, `actions/checkout@v7`, `actions/setup-node@v7`, and `softprops/action-gh-release@v3`. Node setup selects the latest available Node 26 patch with `check-latest: true`. The actions' internal runtime is independent of the Node runtime selected for the SPFx build.

## Validation and remaining checks

Run `npm ci` and `npm run build` to verify lockfile reproducibility, compilation, lint, Jest regressions and production packaging. Inspect the generated manifest to confirm SPFx-provided React resources. The package is written to `sharepoint/solution/spfx-my-calendars-webpart.sppkg`.

Verified locally with Node.js 26.13.0 and npm 11.20.0: `npm ci` succeeds; the production build compiles and packages successfully; all 29 Jest suites pass (243 tests, zero failures). The generated manifest resolves React, React DOM and `react-dom/client` to SPFx-provided 18.3.1 component resources. All eleven direct SPFx packages resolve to 1.24.1. `graphify update .` refreshes the code graph; it reports five existing configuration files with no extracted nodes. Documentation semantic extraction is a separate Graphify operation.

The dependency audit after installation reports 29 vulnerable dependency records: 2 low, 24 moderate, 2 high and 1 critical. These include a critical advisory against the existing direct `dompurify` version and high findings against `body-parser` and `zod`. The SPFx upgrade does not eliminate all audit findings; review these separately rather than using `npm audit fix --force` to change framework-owned dependencies indiscriminately.

Local checks do not execute a GitHub release or establish runtime integration. Validate rendering, theme changes, multiple web parts, property-pane reopening, awaited settings saves, source loading, dialogs and cleanup in SharePoint, Teams, Microsoft 365 and Outlook before deployment. Tenant publication/reload evidence for administrator settings remains tracked separately in DEV-014.
