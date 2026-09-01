# SPFx Toolchain, Packaging, Deployment, Debugging & CI/CD

Source: [Set up your dev environment](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment), [Heft toolchain overview](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/sharepoint-framework-toolchain-rushstack-heft), [Understanding the Heft toolchain](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/customize-heft-toolchain-overview), [Deploy web part to a page](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/serve-your-web-part-in-a-sharepoint-page), [Tenant-scoped deployment](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/tenant-scoped-deployment).

## Which toolchain? Heft (v1.22+) vs Gulp (legacy v1.0–v1.21.1)

**Always confirm the project's SPFx version first** (`@microsoft/sp-core-library` in `package.json`, or the presence of `heft.json` = Heft; `gulpfile.js` = legacy Gulp). Commands are not interchangeable.

| Task | Heft (v1.22+) | Gulp (legacy, v1.0–v1.21.1) |
|---|---|---|
| Trust dev cert | `heft trust-dev-cert` | `gulp trust-dev-cert` |
| Local dev server | `heft start` | `gulp serve` |
| Build | `heft build` | `gulp build` |
| Production bundle | `heft build --production` (bundling folded into build) | `gulp bundle --ship` |
| Package solution | `heft package-solution --production` (check project's npm scripts) | `gulp package-solution --ship` |

Heft (from [RushStack](https://heft.rushstack.io)) is a config-driven task orchestrator that invokes TypeScript, ESLint, Jest, Webpack, and API Extractor — conceptually similar to Vite/esbuild/Gulp/Grunt but standardized for large-scale, consistent builds. Webpack is still the underlying bundler in both toolchains; only the orchestration layer changed. `gulpfile.js` customizations are **not** compatible with Heft — see [Migrate from Gulp to Heft toolchain](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/migrate-gulptoolchain-hefttoolchain.md).

### Customizing the Heft build
- Concepts: actions, phases, tasks, task configurations, plugins, rigs (`references` a shared base config like `@microsoft/spfx-web-build-rig`).
- Use **Heft plugins** for file copy operations, running scripts, etc. — example (copy a license file after packaging) via the `copy-files-plugin` in `config/heft.json`, with `taskDependencies` to sequence after `package-solution`.
- Customize webpack via the **Heft webpack-patch plugin**, or fully **eject the webpack config** if you need complete control (last resort — loses automatic upgrades).
- `config.json` (external library/global config in the old Gulp model) is gone in Heft; equivalent config is now expressed as native webpack config.

### Customizing the legacy Gulp build
- Custom gulp tasks integrated into the build pipeline: [Integrate gulp tasks](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/integrate-gulp-tasks-in-build-pipeline.md).
- Extending webpack directly: [Extending Webpack in the Gulp-based toolchain](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/extending-webpack-in-build-pipeline.md).

## Prerequisites & environment setup

```console
npm install @rushstack/heft yo @microsoft/generator-sharepoint --global
```
- Match Node.js to the SPFx version — see `platforms-and-versions.md` for the compatibility table; SPFx is **only supported on Node LTS releases**, never Current.
- Editor: any, but docs/examples assume VS Code.
- **Self-signed dev cert**: SPFx's local server uses HTTPS; trust it once per machine with `heft trust-dev-cert` (or `gulp trust-dev-cert`). Troubleshooting: delete `{home}/.rushstack` and retry, or manually import `rushstack-serve.pem` into the OS's Trusted Root store if policy blocks automatic trust.
- `SPFX_SERVE_TENANT_DOMAIN` env var — templatizes `{tenantDomain}` in `serve.json` across machines/projects instead of hardcoding a tenant.
- Corporate proxy npm installs: configure npm's `proxy`/`https-proxy` settings.
- Don't run `npm audit fix` reflexively in an SPFx project — it can bump nested deps beyond what the framework has been tested against and break the build; only do targeted upgrades. See [Understanding npm audit vulnerabilities in SPFx projects](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/npm-vulnerabilities.md).

## Yeoman generator

[Yeoman generator for the SharePoint Framework](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/yeoman-generator-for-spfx-intro.md) (`@microsoft/generator-sharepoint`) scaffolds solution type, toolchain files, boilerplate, and (for web parts) the local workbench test harness. A **[SharePoint Framework CLI (spfx-cli)](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/sharepoint-framework-cli.md)** is also available for scripted/non-interactive scaffolding.

## Packaging (`package-solution.json`)

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json",
  "solution": {
    "name": "hello-world-client-side-solution",
    "id": "<GUID>",
    "version": "1.0.0.0",
    "includeClientSideAssets": true,
    "skipFeatureDeployment": true,
    "isDomainIsolated": false,
    "developer": { "name": "", "websiteUrl": "", "privacyUrl": "", "termsOfUseUrl": "", "mpnId": "" },
    "metadata": {
      "shortDescription": { "default": "..." },
      "longDescription": { "default": "..." },
      "screenshotPaths": [], "videoUrl": "", "categories": []
    },
    "webApiPermissionRequests": [ ]
  },
  "paths": { "zippedPackage": "solution/hello-world.sppkg" }
}
```
- **`includeClientSideAssets: true`** → SPFx auto-uploads bundled JS/CSS into the App Catalog's hidden `ClientSideAssets` library (self-contained, no external CDN needed). Set `false` and configure a `cdnBasePath` to host assets on your own CDN instead (e.g., Azure CDN — see [Deploy web parts to Azure CDN](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/deploy-web-part-to-cdn.md), [Enable the Microsoft 365 CDN](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/enable-Microsoft-365-content-delivery-network.md)).
- **`skipFeatureDeployment: true`** → lets the admin choose "make available tenant-wide" at App Catalog upload time, without site-by-site feature activation. See [Tenant-scoped solution deployment](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/tenant-scoped-deployment.md).
- The resulting `.sppkg` is a ZIP containing manifests, `ClientSideAssets.xml`, `feature.xml`, `package.json` metadata — inspectable by renaming to `.zip`.

## Deployment

1. Get/create an **App Catalog** (tenant-wide, one per tenant; or site-collection-scoped for narrower distribution — [site collection app catalog](https://learn.microsoft.com/en-us/sharepoint/dev/general-development/site-collection-app-catalog.md)).
2. Upload the `.sppkg`:
   - UI: drag/drop into the App Catalog's Apps for SharePoint library.
   - PnP PowerShell: `Add-PnPApp -Path .\solution.sppkg -Publish [-SkipFeatureDeployment] [-Overwrite]`.
   - CLI for Microsoft 365: `spo app add --filePath ./solution.sppkg --appCatalogUrl <url>` then `spo app deploy --id <id> [--skipFeatureDeployment]`.
3. If not tenant-wide, install the app into each target site ("Add an app" from Site Contents), or activate the tenant-wide extension listing for extensions with `skipFeatureDeployment`.
4. Deployment for **Microsoft Teams**-surfaced SPFx apps has its own considerations (Teams app package + Teams admin center or org store) — see [Deployment options for SPFx solutions & Microsoft Teams](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/deployment-spfx-teams-solutions).

## Debugging

- [Debug in Visual Studio Code](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/debug-in-vscode.md) — launch configs, breakpoints in TS source via source maps.
- [Debug on modern pages](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/debug-modern-pages.md) — load local dev bundles into a real SharePoint page using the `?debug=true&noredir=true&debugManifestsFile=...` query string (SPFx v1.21+ default: `https://localhost:4321/temp/build/manifests.js`).
- **[Debug Toolbar](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/debug-toolbar.md)** — the modern, recommended way to load/toggle local components on any SharePoint page (the classic hosted workbench is retiring Dec 1, 2026).
- **[Developer Dashboard](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-developer-dashboard.md)** — `Ctrl+F12` opens an in-page performance/logging panel; SPFx components can write to it via `Log.info/warn/error` from `@microsoft/sp-core-library`.
- Extensions specifically require `config/serve.json` pointed at a live page (see `extensions.md`) — they can't use the Workbench.

## CI/CD

- [Azure DevOps build & release stages](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/implement-ci-cd-with-azure-devops.md) — classic pipeline pattern (build stage → package → release stage → deploy to app catalog).
- [Azure DevOps multi-stage pipelines](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/implement-ci-cd-with-azure-pipelines.md) — modern YAML pipeline pattern.
- Typical stages: `npm ci` → lint/test → `heft build --production` (or `gulp bundle --ship`) → `package-solution --production`/`--ship` → publish `.sppkg` artifact → deploy step (PnP PowerShell or CLI for M365) against target app catalog, often gated per environment (dev/test/prod tenant).

## Content Security Policy (CSP)

[SharePoint Online support for CSP / trusted script sources](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/content-securty-policy-trusted-script-sources.md) — if the tenant enforces CSP, external script/style/API origins used by a component may need to be allow-listed; check this if a component works locally but fails silently (blocked network calls) once deployed.

## Optimize builds

[Optimize builds for production](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/optimize-builds-for-production.md) — bundle-size and load-time guidance (tree-shaking, splitting, avoiding heavy dependencies) — bring this up if the user reports slow-loading web parts.

## Dynamic loading & Dynamic Data

- [Dynamic loading of packages](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/dynamic-loading.md) — lazy-load heavy dependencies only when needed.
- [Connecting components with Dynamic Data](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/dynamic-data.md) — SPFx's built-in pub/sub mechanism for wiring one web part's output to another's input on the same page (source/consumer property panes).
