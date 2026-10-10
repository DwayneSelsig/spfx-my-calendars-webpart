# Dependency compatibility

This is an operational dependency review, not a product behavior contract. Versions were checked against the npm registry and upstream release documentation on October 11, 2026.

## SPFx constraints

The project uses the latest stable SPFx release, **1.23.2**. Microsoft's [compatibility table](https://learn.microsoft.com/sharepoint/dev/spfx/compatibility) requires **Node.js 22**, **React and React DOM exactly 17.0.1**, and **TypeScript no newer than 5.8**. React patch versions must not be inferred from semver compatibility. The [SPFx release notes](https://learn.microsoft.com/sharepoint/dev/spfx/release-1.23.2) also require coordinated SPFx package updates.

The latest supported TypeScript patch is 5.8.3. React typings can use the latest 17.x patches: the published SPFx core package accepts React typings below 18. ESLint 9 is accepted by the SPFx ESLint config and plugin; ESLint 10 is excluded by their peer dependency ranges.

Microsoft's compatibility table does not certify every third-party package version. For those packages, the evidence is their published dependency/peer ranges and this project's build and regression tests. This distinction also applies to newer root build tools when the SPFx rig still owns older, exact transitive versions.

## Direct dependencies

“Before” is the version resolved by the previous lockfile. “Selected” is the version resolved by the updated lockfile. “Latest” refers to the stable npm `latest` tag, not previews.

| Package | Before | Selected | Latest | Reason |
| --- | --- | --- | --- | --- |
| `@fluentui/react` | 8.125.7 | 8.125.8 | 8.125.8 | Latest v8; published peers include React 17. |
| `dompurify` | 3.4.15 | 3.4.16 | 3.4.16 | Latest stable patch. |
| `react`, `react-dom` | 17.0.1 | 17.0.1 | 19.3.0 | Exact runtime version required by SPFx. |
| `tslib` | 2.3.1 | 2.8.1 | 2.8.1 | Latest stable TypeScript helper library. |
| `@rushstack/eslint-config` | 4.6.4 | 4.8.0 | 4.8.0 | Latest stable; peers accept ESLint 9 and TypeScript 5.8. |
| `@rushstack/heft` | 1.2.25 | 1.3.3 | 1.3.3 | Latest stable; SPFx plugin accepts `^1.1.5`. See upstream peer inconsistency below. |
| `@types/heft-jest` | 1.0.6 | 1.0.6 | 1.0.6 | Already latest. |
| `@types/react` | 17.0.45 | 17.0.93 | 19.3.0 | Latest React 17 typings; React 19 typings violate SPFx peers. |
| `@types/react-dom` | 17.0.17 | 17.0.26 | 19.3.0 | Latest React 17 typings; React 19 typings violate SPFx peers. |
| `@types/webpack-env` | 1.15.3 | 1.18.8 | 1.18.8 | Latest stable ambient Webpack types. |
| `@typescript-eslint/parser` | 8.48.1 | 8.71.1 | 8.71.1 | Peers accept ESLint 9 and TypeScript 5.8. |
| `css-loader` | 7.1.5 | 7.1.5 | 7.1.5 | Already resolved to latest; raised the manifest minimum. |
| `eslint` | 8.57.1 | 9.39.5 | 10.12.0 | Latest 9.x accepted by SPFx; migrated to the SPFx flat React profile. |
| `eslint-plugin-react` | 7.37.5 | 7.37.5 | 7.37.5 | Already resolved to latest; raised the manifest minimum. |
| `eslint-plugin-react-hooks` | 4.6.2 | 7.1.1 | 7.1.1 | Latest stable; peers accept ESLint 9. Flat config loads this root plugin. |
| `sass-embedded` | 1.85.1 | 1.105.1 | 1.105.1 | Latest root compiler; the rig's Sass plugin retains its own `~1.85.1` dependency. |
| `typescript` | 5.8.3 | 5.8.3 | 7.0.2 | Latest patch within Microsoft's supported 5.8 ceiling. |

All eleven direct SPFx packages remain at 1.23.2, which is their latest stable release:

- Runtime: `@microsoft/sp-component-base`, `@microsoft/sp-core-library`, `@microsoft/sp-lodash-subset`, `@microsoft/sp-office-ui-fabric-core`, `@microsoft/sp-property-pane`, and `@microsoft/sp-webpart-base`.
- Build: `@microsoft/eslint-config-spfx`, `@microsoft/eslint-plugin-spfx`, `@microsoft/sp-module-interfaces`, `@microsoft/spfx-heft-plugins`, and `@microsoft/spfx-web-build-rig`.

The lockfile also refreshes transitive dependencies within their owners' declared ranges. Exact SPFx-owned transitive versions remain intact. Installing a newer direct package does not replace an exact nested dependency. No `--force`, `--legacy-peer-deps`, or dependency overrides were used.

### Published package evidence

The manifests returned by `npm view <package>@<version> dependencies peerDependencies engines --json` establish the package constraints. Relevant upstream manifests are available from the registry:

- [SPFx core 1.23.2](https://registry.npmjs.org/@microsoft/sp-core-library/1.23.2): React and typings peers.
- [SPFx ESLint config 1.23.2](https://registry.npmjs.org/@microsoft/eslint-config-spfx/1.23.2): ESLint peer range and exact internal dependencies.
- [SPFx Heft plugins 1.23.2](https://registry.npmjs.org/@microsoft/spfx-heft-plugins/1.23.2): Heft peer range.
- [SPFx build rig 1.23.2](https://registry.npmjs.org/@microsoft/spfx-web-build-rig/1.23.2): rig-owned compiler, plugins, and tool versions.
- [Rushstack ESLint config 4.8.0](https://registry.npmjs.org/@rushstack/eslint-config/4.8.0), [TypeScript ESLint parser 8.71.1](https://registry.npmjs.org/@typescript-eslint/parser/8.71.1), and [Fluent UI 8.125.8](https://registry.npmjs.org/@fluentui/react/8.125.8): third-party compatibility ranges.

### Upstream Heft peer inconsistency

The SPFx rig pins Heft 1.2.17. Most of its plugins accept newer Heft 1.x versions, but its lint plugin 1.2.7 and TypeScript plugin 1.3.2 declare **exactly Heft 1.2.7** as a peer. Those two declarations conflict with the rig's own Heft version before any project update. npm therefore emits peer warnings even when the root version is aligned with the rig.

The root CLI is updated to Heft 1.3.3 as requested, while the rig's packages remain unchanged. A successful project build is evidence for this tested combination, not a claim that the upstream peer inconsistency is resolved or that all peers are satisfied.

## Release workflow

The existing file is `.github/workflows/main.yml`.

| Step / environment | Selected reference | Stable release checked |
| --- | --- | --- |
| Checkout | `actions/checkout@v7` | [7.0.1](https://github.com/actions/checkout/releases/tag/v7.0.1) |
| Node setup | `actions/setup-node@v7` | [7.1.0](https://github.com/actions/setup-node/releases/tag/v7.1.0) |
| Release upload | `softprops/action-gh-release@v3` | [3.0.3](https://github.com/softprops/action-gh-release/releases/tag/v3.0.3) |
| Ubuntu runner | `ubuntu-latest` | Follows GitHub's latest Ubuntu image automatically. |
| Project Node.js | `22`, `check-latest: true` | Latest available patch in the supported Node 22 line. |

The floating major action tags receive minor and patch releases automatically; a new action major still requires an explicit update. The actions' internal Node 24 runtime is independent of the Node 22 environment installed for the SPFx build. The release action's [v3 runtime change](https://github.com/softprops/action-gh-release/releases/tag/v3.0.0) addresses the Node 20 deprecation warning.

As requested, Ubuntu remains `ubuntu-latest`. GitHub's [migration announcement](https://github.com/actions/runner-images/issues/14748) schedules its move to Ubuntu 26.04 between October 19 and November 19, 2026. The announcement warning can remain until that migration completes; it is not an outdated action runtime warning.

## Rechecking updates

Run `npm outdated` to find newer direct dependencies, and verify the compatibility table and package manifests before changing constrained versions. At this review, the remaining newer direct `latest` tags are React/React DOM, their typings, TypeScript, and ESLint; all are intentionally bounded by SPFx.

Use `npm ci` and `npm run build` to check lockfile reproducibility, compilation, linting, existing Jest regressions, and production packaging. GitHub runner and upload behavior require a subsequent release workflow run.

## Audit results

The same-day `npm audit` comparison reports 67 vulnerable dependency records before the update (1 low, 34 moderate, 29 high, 3 critical) and 54 afterwards (33 moderate, 21 high, 0 critical). These totals include propagated findings on parent packages, not 54 distinct advisories. The remaining high findings are in the build tool dependency chains, including the SPFx rig, glob matching, development server, and certificate tools. The critical `fast-xml-parser`, `proxy-addr`, and `shell-quote` findings no longer appear after the compatible transitive refresh.

Updating direct packages does not eliminate findings behind fixed upstream dependencies. Resolving every remaining finding requires upstream package changes or a separate review of dependency overrides; this update does not force versions outside the published constraints.

## Validation

- A clean `npm ci` succeeds from the updated lockfile; upstream peer and deprecation warnings remain as described above.
- Production compilation and linting succeed with TypeScript 5.8.3 and ESLint 9.39.5. The flat config preserves the existing project rule overrides and disable-directive behavior.
- All 28 existing Jest suites pass: 234 successful tests, zero failures.
- Production packaging creates `sharepoint/solution/spfx-my-calendars-webpart.sppkg`.
- Workflow YAML parses successfully and selects `ubuntu-latest`, the three requested action major tags, and Node 22 with `check-latest: true`.
- All direct lockfile versions satisfy the manifest; the eleven SPFx packages remain at 1.23.2.
- `graphify update .` refreshes the graph. It warns that five configuration JSON files (`hooks.json`, `launch.json`, `settings.json`, `tasks.json`, and `.yo-rc.json`) produce no nodes and are omitted. It also reports changed community labels; these generated navigation labels are not product authority.

These checks do not execute a GitHub release or validate the web part in a live SharePoint tenant.
