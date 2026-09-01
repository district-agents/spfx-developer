# SPFx Extensions

Source: [Overview of SPFx Extensions](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/overview-extensions), [Build your first Extension (Hello World)](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/build-a-hello-world-extension), [Tenant-wide deployment](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/basics/tenant-wide-deployment-extensions).

## The five extension types

| Type | Purpose |
|---|---|
| **Application Customizer** | Injects script/UI into well-known page placeholders (e.g. top, bottom) and runs on every page load. Base class: `BaseApplicationCustomizer`. |
| **Field Customizer** | Overrides how a specific column's values render in list views. |
| **Command Set** | Adds custom actions/buttons to the command surfaces (toolbar and item context menus) for lists/libraries. |
| **Form Customizer** | Replaces the default New/Edit/Display form for a list or content type. |
| **Search Query Modifier** | Modifies SharePoint search queries before execution (inject filters, refine results, adjust behavior). |

All five are available in every Microsoft 365 subscription for production use. Extensions **cannot** use plain iframes for isolation — like web parts, they run embedded in the page.

## Scaffolding & first Application Customizer

```console
md app-extension && cd app-extension
yo @microsoft/sharepoint
```
Prompts: solution name → **Which type of client-side component to create?** → `Extension` → **Which type of client-side extension to create?** → `Application Customizer` → name.

Manifest (`HelloWorldApplicationCustomizer.manifest.json`):
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-extension-manifest.schema.json",
  "id": "<GUID>",
  "alias": "HelloWorldApplicationCustomizer",
  "componentType": "Extension",
  "extensionType": "ApplicationCustomizer",
  "version": "*",
  "manifestVersion": 2,
  "requiresCustomScript": false
}
```

Component code:
```typescript
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';

export default class HelloWorldApplicationCustomizer
  extends BaseApplicationCustomizer<IHelloWorldApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    // this.context and this.properties are available here (NOT in the constructor)
    let message = this.properties.testMessage ?? '(No properties were provided.)';
    Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`);
    return Promise.resolve();
  }
}
```
`onInit()` is the entry point (constructor runs too early — `this.context`/`this.properties` are undefined there).

## Testing extensions — the Workbench does NOT work

Extensions must be tested against a **real, live modern SharePoint page** via `config/serve.json`:
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/spfx-serve.schema.json",
  "port": 4321,
  "https": true,
  "serveConfigurations": {
    "default": {
      "pageUrl": "https://contoso.sharepoint.com/sites/mySite/SitePages/myPage.aspx",
      "customActions": {
        "<extension-GUID-from-manifest>": {
          "location": "ClientSideExtension.ApplicationCustomizer",
          "properties": { "testMessage": "Test message" }
        }
      }
    }
  }
}
```
Run `heft start` (or `gulp serve`), then in the browser append the debug query string (SPFx generates/shows it in the console) or select **Load debug scripts** when prompted. Update `serve.json` whenever you add components or change their properties/GUIDs.

Use `SPFX_SERVE_TENANT_DOMAIN` environment variable to templatize `{tenantDomain}` across serve configurations without hardcoding a tenant in source control.

## Deploying / activating extensions

- Deploy the `.sppkg` to an app catalog like any SPFx solution (`references/toolchain-and-deployment.md`).
- Extensions need a `ClientSideComponentId` associated with a site/list/content type via a SharePoint **feature/element** (the Yeoman scaffold generates this) — see [Deploy your extension to SharePoint](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/serving-your-extension-from-sharepoint).
- `ClientSideComponentProperties` on that association supplies the extension's runtime properties (equivalent of `serve.json`'s `properties` in production).
- **[Tenant-wide deployment](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/basics/tenant-wide-deployment-extensions)**: to auto-activate an Application Customizer or Command Set across the *entire tenant* without a per-site feature activation, set `skipFeatureDeployment: true` in `package-solution.json`. This surfaces a **Tenant Wide Extensions** list in the app catalog site that manages activations tenant-wide. Supported for Application Customizers and List View Command Sets.
- [Pre-allocated placeholders](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/basics/preallocated-space-placeholders.md) — reserve layout space for an Application Customizer to avoid layout shift ("flash") when it renders.

## Other extension-specific guidance

- [Use page placeholders](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/using-page-placeholder-with-extensions.md) — the `PlaceholderContent` API (`this.context.placeholderProvider.tryCreateContent(...)`) for Application Customizers.
- [Configure extension icon](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/basics/configure-extension-icon.md).
- [Field Customizer tutorial](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/building-simple-field-customizer.md).
- [ListView Command Set tutorial](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/building-simple-cmdset-with-dialog-api.md), and [grouping of Command Set extensions](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/guidance/list-view-command-set-grouping.md) when multiple command sets target the same list.
- [Search Query extension](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/building-search-extensions.md).
- [Form Customizer tutorial](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/building-form-customizer.md).
- [Using custom dialogs with SPFx](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/guidance/using-custom-dialogs-with-spfx.md) — the `@microsoft/sp-dialog` `Dialog`/custom `BaseDialog` API for modal UI from extensions.

## Migrating legacy customizations to Extensions

Point users at these when modernizing:
- Client-Side Rendering (JSLink) → Field Customizers (`guidance/migrate-jslink-to-spfx-extensions.md`, plus a full tutorial `guidance/migrate-from-jslink-to-spfx-extensions.md`).
- JavaScript embedding via `UserCustomAction` → Application Customizers (`guidance/migrate-user-customactions-to-spfx-extensions.md`, tutorial `guidance/migrate-from-usercustomactions-to-spfx-extensions.md`).
- Edit Control Block (ECB) custom menu items → Command Set extensions (`guidance/migrate-from-ecb-to-spfx-extensions.md`).
