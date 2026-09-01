---
name: spfx-developer
description: Expert guidance for developing, debugging, packaging, and deploying SharePoint Framework (SPFx) solutions — web parts, extensions (Application Customizer, Field Customizer, Command Set, Form Customizer, Search Query Modifier), Adaptive Card Extensions for Viva Connections, Copilot Apps, and Teams/Outlook/Office apps built on SPFx, using PnPjs (`@pnp/sp`/`@pnp/graph`) for data access, `@pnp/spfx-controls-react` for canvas UI, and `@pnp/spfx-property-controls` for property pane fields. Use this skill any time the user mentions SPFx, SharePoint Framework, client-side web parts, SharePoint extensions, the Yeoman generator (`yo @microsoft/sharepoint`), `heft`/`gulp` SPFx commands, `package-solution.json`, `.sppkg` files, the SharePoint Workbench, PnPjs, `spfi`/`graphfi`, `@pnp/spfx-controls-react`, `@pnp/spfx-property-controls`, SPHttpClient/AadHttpClient/MSGraphClient, or asks to scaffold, build, style, debug, package, or deploy a SharePoint web part or extension — even if they just say "web part" or "SharePoint customization" without saying "SPFx" explicitly.
---

# SharePoint Framework (SPFx) Developer

Reference-grade guidance for building SPFx solutions, sourced from the official Microsoft Learn SharePoint Framework documentation (`learn.microsoft.com/sharepoint/dev/spfx`) and the `SharePoint/sp-dev-docs` GitHub source, current as of SPFx v1.24 (August 2026).

## What SPFx is (orient yourself first)

SPFx is Microsoft's client-side page-and-web-part extensibility model for SharePoint Online, SharePoint Server (2016+), Microsoft Teams, Outlook/Office, and Microsoft Viva Connections. Key properties to keep in mind while helping the user:

- **Framework-agnostic**: React, Angular, Vue, Handlebars, Knockout, or plain TS/JS all work.
- **Runs in the page context**: no iframes; controls render directly in the page DOM; automatic SSO; automatic hosting in SharePoint (no separate web server needed for production).
- **Toolchain**: npm, TypeScript, Yeoman, webpack. As of **SPFx v1.22+**, the build orchestrator is **Heft** (`heft` commands), replacing the legacy **Gulp** toolchain used in v1.0–v1.21.1. Always ask/check which SPFx version the user's project uses before giving build commands — the commands differ (see `references/toolchain-and-deployment.md`).
- **Component types**: Web parts, five Extension types, Adaptive Card Extensions (Viva Connections), Copilot Apps, and Library Components.
- **The page DOM is not an API** — never suggest depending on SharePoint's HTML/CSS structure directly; always use SPFx APIs.
- The classic **SharePoint Workbench** (hosted online workbench) is deprecated (retiring Dec 1, 2026); the recommended local/live testing surface going forward is the **SharePoint Framework Debug Toolbar**.

## Decision tree — which component type does the user need?

1. **Adds a self-contained, configurable block of content/UI that a page author drags onto a page (or Teams tab / Viva dashboard)?** → **Web part**. See `references/web-parts.md`.
2. **Needs to run automatically on every page load / inject UI into standard chrome (headers, notifications) without being manually added?** → **Application Customizer** extension.
3. **Needs to change how a specific list/library *column's values* render?** → **Field Customizer** extension.
4. **Needs to add buttons/menu items to the command bar (ribbon) for list items or the list itself?** → **Command Set** extension.
5. **Needs to replace the New/Edit/Display form for a list or content type?** → **Form Customizer** extension.
6. **Needs to intercept/modify SharePoint search queries before execution?** → **Search Query Modifier** extension.
7. **Building a card-based experience for the Viva Connections dashboard?** → **Adaptive Card Extension (ACE)**. See `references/teams-viva-copilot.md`.
8. **Building an agent/Copilot-surfaced app?** → **Copilot App**. See `references/teams-viva-copilot.md`.
9. **Sharing reusable non-UI logic across components?** → **Library component**.

All types share the same solution structure, manifest pattern, property model, and packaging pipeline — see `references/web-parts.md` and `references/extensions.md` for the specifics of each.

## Quick-start workflow (the golden path)

1. **Prereqs** (see `references/toolchain-and-deployment.md` for full detail and troubleshooting):
   ```console
   npm install @rushstack/heft yo @microsoft/generator-sharepoint --global
   ```
   Match Node.js LTS version to the SPFx version (SPFx v1.22.x → Node 22 LTS; check `references/platforms-and-versions.md` before assuming).

2. **Scaffold a project**:
   ```console
   md my-solution && cd my-solution
   yo @microsoft/sharepoint
   ```
   Prompts to expect: solution name → component type (WebPart/Extension/Library/Copilot App) → sub-type → component name → framework (No framework / React / etc.).

3. **Trust the dev cert once per machine** (Heft toolchain): `heft trust-dev-cert` (legacy Gulp: `gulp trust-dev-cert`).

4. **Run locally**: Heft toolchain → `heft start`. Legacy Gulp → `gulp serve`. This starts `localhost:4321`/`5432` and (for web parts) opens the hosted workbench or, for extensions, needs `./config/serve.json` pointed at a real modern page (extensions can't use the Workbench).

5. **Iterate on code** — key files for any component:
   - `src/**/<Component>.manifest.json` — id (GUID), type, version, `requiresCustomScript`.
   - `src/**/<Component>WebPart.ts` / `<Component>ApplicationCustomizer.ts` etc. — the component class extending the appropriate SPFx base class.
   - `config/package-solution.json` — solution-level packaging metadata, `webApiPermissionRequests`, `skipFeatureDeployment`.
   - `config/serve.json` — local debug configuration (required for testing extensions against a real page).

6. **Package for deployment**:
   - Heft: `heft build --production` then `heft package-solution --production` (check the project's actual npm scripts — many still expose `npm run package-solution`).
   - Legacy Gulp: `gulp bundle --ship` then `gulp package-solution --ship`.
   - Output: a `.sppkg` file under `sharepoint/solution/`.

7. **Deploy**: upload the `.sppkg` to a SharePoint **App Catalog** (tenant-wide or site-collection-scoped). If `skipFeatureDeployment: true` in `package-solution.json`, the admin can make it available tenant-wide without per-site activation. See `references/toolchain-and-deployment.md` for CLI/PowerShell deployment commands and CI/CD pipelines.

## Cheat sheet: connecting to data

**Default to PnPjs (`@pnp/sp` / `@pnp/graph`) for all SharePoint and Graph data access** — it's the preferred library for this skill. Reach for the raw native clients only for a single trivial one-off call.

| Need | Use | Reference |
|---|---|---|
| SharePoint REST (lists, items, files, folders, search, etc.) | **PnPjs** `spfi().using(SPFx(this.context))`, fluent chain e.g. `sp.web.lists.getByTitle(...).items()` | `references/pnpjs.md` |
| Microsoft Graph | **PnPjs** `graphfi().using(SPFx(this.context))`, or `this.context.msGraphClientFactory.getClient('3')` for one-offs | `references/pnpjs.md`, `references/apis-and-data.md` |
| One-off simple SharePoint REST GET, no new dependency wanted | `this.context.spHttpClient` (`SPHttpClient`) — built in | `references/apis-and-data.md` |
| Any other Entra ID (Azure AD)-secured API | `this.context.aadHttpClientFactory.getClient(resourceUri)` (`AadHttpClient`) — requires `webApiPermissionRequests` in `package-solution.json`, admin approval | `references/apis-and-data.md` |
| Anonymous/public API | `this.context.httpClient` (`HttpClient`) | `references/apis-and-data.md` |

## Cheat sheet: UI building blocks

**Before hand-building a picker, form, or common piece of chrome, check the controls libraries first:**
- **Property pane fields** (configuring the web part/extension itself) → `references/property-controls.md` (`@pnp/spfx-property-controls`) — list/site/term/color/date/people pickers, a collection-data editor, code/Monaco editors, callout-annotated fields, all usable directly inside `getPropertyPaneConfiguration()`.
- **Canvas/body UI** (what the component renders for end users) → `references/react-controls.md` (`@pnp/spfx-controls-react`) — pickers, an auto-generated list-item form, a filterable/groupable list view, permission-trimmed sections, the standard "unconfigured web part" placeholder, an edit-mode-aware web part title, charts, and more.

Use a matching control by default; only write custom UI when nothing there fits or the design genuinely needs something different.

## Common pitfalls to flag proactively

- **Toolchain mismatch**: giving `gulp` commands for a v1.22+ (Heft) project or vice versa. Always check `package.json`/`@microsoft/sp-core-library` version first.
- **Testing extensions in the Workbench** — doesn't work; extensions must be tested against a real modern page via `config/serve.json`.
- **Forgetting `webApiPermissionRequests`** when calling Graph/AAD-secured APIs — the call will fail until an admin approves the permission in the SharePoint admin center (or via PnP PowerShell/CLI for M365).
- **Assuming `skipFeatureDeployment` auto-grants API permissions** — it doesn't; API permission approval is a separate admin step and deployment succeeds either way.
- **Hardcoding colors instead of theme tokens** — breaks when site theme/section background changes; use SASS theme slots or `ThemeProvider`. See `references/design-and-theming.md`.
- **Taking a dependency on the SharePoint page DOM/CSS structure** — unsupported and breaks on any page-chrome update.
- **Running `npm audit fix`** in an SPFx project — can upgrade nested dependencies beyond what SPFx has been tested against; avoid unless the user explicitly wants to troubleshoot vulnerabilities (see `references/toolchain-and-deployment.md`).
- **On-prem SharePoint (2016/2019/SE)**: much lower SPFx version ceiling (2016→v1.1, 2019/SE→v1.4.1) and no AadHttpClient/MSGraphClient beyond that ceiling. Also caps the PnPjs version usable (PnPjs v4 needs SPFx ≥1.18 — on-prem projects must use PnPjs v2). Always ask about the target platform before recommending APIs. See `references/platforms-and-versions.md`, `references/pnpjs.md`.
- **PnPjs selective-import gotcha**: importing only the type (`import { IList } from "@pnp/sp/lists"`) without the matching side-effect import (`import "@pnp/sp/lists"`) compiles fine but throws at runtime. Both are needed. See `references/pnpjs.md`.
- **Building a custom picker/form/list-view from scratch** when `@pnp/spfx-controls-react` already has one — check `references/react-controls.md` first.
- **Using native `PropertyPaneTextField`/`PropertyPaneDropdown` for a list/site/term/people/color/date pane field** when `@pnp/spfx-property-controls` already has a purpose-built one — check `references/property-controls.md` first.

## Reference files (read the relevant one before giving detailed/code-level guidance)

- `references/web-parts.md` — web part anatomy, property pane, theming hooks, full-width column, Teams tabs, migration from Script Editor/JSLink/jQuery.
- `references/extensions.md` — all 5 extension types, manifests, placeholders, tenant-wide deployment, custom dialogs.
- `references/pnpjs.md` — **preferred data-access library**: `@pnp/sp`/`@pnp/graph` v4 setup, selective imports, CRUD, batching, error handling, search, v3→v4 migration.
- `references/react-controls.md` — **preferred canvas/body UI library**: `@pnp/spfx-controls-react` decision guide and usage patterns for pickers, forms, list views, placeholders, and more.
- `references/property-controls.md` — **preferred property-pane UI library**: `@pnp/spfx-property-controls` decision guide and usage patterns for list/site/term/people/color/date pickers, collection-data editor, code/Monaco editors, and callout-annotated fields.
- `references/apis-and-data.md` — native SPHttpClient/AadHttpClient/MSGraphClient clients (what PnPjs wraps), webhooks, anonymous APIs, permission-request/approval flow.
- `references/toolchain-and-deployment.md` — Heft vs Gulp toolchains, CLI commands, packaging, app catalog deployment, CI/CD (Azure DevOps), debugging (VS Code, Debug Toolbar, Dev Dashboard), CSP.
- `references/design-and-theming.md` — Fluent UI, theme tokens/SASS slots, ThemeProvider, responsive/section-aware design, accessibility, empty states.
- `references/teams-viva-copilot.md` — Teams tabs/personal apps, Viva Connections Adaptive Card Extensions, Copilot Apps, Outlook/Office extensibility.
- `references/platforms-and-versions.md` — SPFx version history/roadmap, Node.js compatibility matrix, supported platforms (SPO, SPS 2016/2019/SE, Teams, Viva), migration paths from Add-ins/JSLink/UserCustomActions.
- `references/governance-and-guidance.md` — enterprise/team-based development guidance, external libraries, CSP, dynamic loading, Dynamic Data, localization, publishing to the marketplace.

Always prefer these reference files and the cited documentation (Microsoft Learn, PnPjs, sp-dev-fx-controls-react, sp-dev-fx-property-controls) over general training knowledge — SPFx's toolchain (Gulp→Heft), PnPjs (v3→v4), and the controls libraries have all changed materially over time, and the user's project version matters a lot.
