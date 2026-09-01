# SPFx Web Parts

Source: [Overview of SharePoint client-side web parts](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/overview-client-side-web-parts), [Build your first web part](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part), [Connect to SharePoint (Hello World part 2)](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/connect-to-sharepoint), [Property pane overview](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/integrate-web-part-properties-with-sharepoint).

## What they are

Client-side web parts are controls that appear inside a SharePoint page and execute entirely in the browser. They're the building block for:
- Content on modern & classic SharePoint pages.
- **Single Part App Pages (SPAs)** — a web part as the entire page experience.
- **Microsoft Teams tabs** (personal apps and channel/group tabs) — the *same* web part code, no changes required.
- Can be built with plain JS/TS or any framework: React, Angular, Vue, Handlebars, Knockout.

## Scaffolding

```console
yo @microsoft/sharepoint
```
Prompts: solution name → **Which type of client-side component to create?** → `WebPart` → web part name → template (`No framework`, `React`, etc.).

## Anatomy

- **`<Name>WebPart.ts`** — the component class, extends `BaseClientSideWebPart<IPropsInterface>`. Implements:
  - `render()` — sets `this.domElement.innerHTML` (or mounts a React tree) using `this.properties.*`, `this.context.pageContext`, etc.
  - `getPropertyPaneConfiguration()` — defines the property pane UI.
  - Lifecycle hooks beyond render/load: `onInit()`, `onDispose()`, `onPropertyPaneFieldChanged()`, serialize/deserialize, configuration-change handling.
- **`<Name>WebPart.manifest.json`** — required for every web part:
  ```json
  {
    "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json",
    "id": "<GUID — unique per component>",
    "alias": "HelloWorldWebPart",
    "componentType": "WebPart",
    "version": "*",
    "manifestVersion": 2,
    "requiresCustomScript": false,
    "supportedHosts": ["SharePointWebPart", "TeamsPersonalApp", "TeamsTab", "SharePointFullPage"],
    "supportsThemeVariants": true,
    "preconfiguredEntries": [{
      "groupId": "5c03119e-3074-46fd-976b-c60198311f70",
      "group": { "default": "Advanced" },
      "title": { "default": "HelloWorld" },
      "description": { "default": "HelloWorld description" },
      "officeFabricIconFontName": "Page",
      "properties": { "description": "HelloWorld" }
    }]
  }
  ```
  - `requiresCustomScript: true` is needed only if the component lets end users embed arbitrary script (rare — avoid unless truly needed, since it restricts where the component can be installed).
  - `supportedHosts` controls whether it can run as a Teams personal app/tab in addition to SharePoint.

## Property pane

- Import field types from `@microsoft/sp-property-pane`: `PropertyPaneTextField`, `PropertyPaneCheckbox`, `PropertyPaneDropdown`, `PropertyPaneToggle`, `PropertyPaneSlider`, `PropertyPaneChoiceGroup`, etc.
- **For anything richer** — a list/site/term/color/date/people picker, a repeating collection-data editor, a code editor, or a callout-annotated field — use `@pnp/spfx-property-controls` instead of building it from native fields. See `references/property-controls.md` for the full decision guide and usage patterns; it uses the same `groupFields`-array placement as the native fields shown here.
- `getPropertyPaneConfiguration()` returns `pages[].groups[].groupFields[]`.
- Access values in code via `this.properties.<name>`; **always HTML-escape string properties** before interpolating into `innerHTML` (`escape(this.properties.description)`).
- Default values for properties go in the manifest's `preconfiguredEntries[].properties`.
- **Update behavior**: reactive (default, live preview as you type) vs non-reactive (`onPropertyPaneConfigurationStart`/explicit apply) — see [Design: reactive and nonreactive web parts](https://learn.microsoft.com/en-us/sharepoint/dev/design/reactive-and-nonreactive-web-parts).
- Validate property values: [Validate web part property values](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/validate-web-part-property-values).
- Extend the pane with **custom property pane controls** (e.g., a people picker, list picker) and **cascading dropdowns** — see `guidance/build-custom-property-pane-controls.md` and `guidance/use-cascading-dropdowns-in-web-part-properties.md` in the docs.

## Connecting to SharePoint data (Hello World part 2 pattern)

```typescript
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

private _getListData(): Promise<ISPLists> {
  return this.context.spHttpClient
    .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`,
         SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => response.json());
}
```
For anything beyond simple GETs (batching, complex filtering, type safety), prefer **PnPjs** — see `apis-and-data.md`.

## Local preview

- Heft toolchain: `heft start` (opens hosted workbench at the URL in `config/serve.json`'s `initialPage`, or use the newer **Debug Toolbar** since the online hosted workbench is deprecated Dec 1, 2026).
- Legacy Gulp: `gulp serve`.
- VS Code: `Ctrl+Shift+B` / `Cmd+Shift+B` runs the default build task.

## Full-width / responsive / section-aware design

- [Use full-width column](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/basics/use-web-parts-full-width-column.md) and [determine rendered web part width](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/basics/determine-web-part-width.md) — web parts should be responsive to the column width they're placed in, not assume a fixed size.
- [Support section backgrounds](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/supporting-section-backgrounds.md) — adapt to the 4 theme variants (main/neutral/soft/strong) a page section can use; see `design-and-theming.md`.
- Web part top actions (contextual command bar actions) — `guidance/getting-started-with-top-actions.md`.

## Toolbox / preconfigured entries

- [Preconfigure web parts](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/simplify-adding-web-parts-with-preconfigured-entries.md) — ship multiple ready-made configurations of the same web part.
- [Configure web part icon](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/basics/configure-web-part-icon.md) / [hide from the toolbox](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/basics/hide-web-part-from-toolbox.md) (useful for web parts only meant to be provisioned programmatically, e.g. dashboard templates).

## Isolated web parts & Single Part App Pages

- **Isolated web parts** run in a separate origin/context for stronger isolation — note there is a retirement announcement for this feature; check `spfx/web-parts/isolated-web-parts-retirement.md` before recommending it for new work.
- **[Single Part App Pages](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/single-part-app-pages)** — a page that is entirely one web part, useful for SPA-style experiences and Teams tabs.

## Microsoft Teams tabs from the same web part

A web part with `TeamsTab`/`TeamsPersonalApp` in `supportedHosts` can run as-is inside Microsoft Teams (channel/group tab or personal app) with no code changes — see [Building Microsoft Teams Tabs using SharePoint Framework](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/integrate-with-teams-introduction) and `teams-viva-copilot.md`. `this.context.sdks.microsoftTeams` is populated when running inside Teams — use it to branch UI (the Hello World scaffold already checks `!!this.context.sdks.microsoftTeams` to toggle a `styles.teams` class).

## Migrating older customizations into web parts

If the user is modernizing legacy SharePoint customizations, point them at these guidance docs (all under `spfx/web-parts/guidance/`):
- Script Editor web part → SPFx web part.
- jQuery + DataTables / jQuery + FullCalendar scripts → SPFx.
- AngularJS (1.x) applications → SPFx.
- Working with `__REQUESTDIGEST` in client code.
- Connecting to SharePoint via JSOM (legacy) vs modern REST/PnPjs.

## Tutorials worth pointing to for deeper scenarios

- Add jQueryUI Accordion to a web part.
- Use Microsoft Graph in web parts / Microsoft Graph Toolkit in web parts.
- Provision SharePoint assets from a web part package.
- Build Microsoft Teams channel/group tabs with SPFx.
