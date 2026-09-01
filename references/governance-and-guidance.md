# Enterprise Guidance, Governance & Publishing

Source: [Enterprise guidance](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/enterprise-guidance), [Team-based development](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/team-based-development-on-sharepoint-framework), [Governance considerations](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/governance-considerations), [Publish to marketplace overview](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview).

## Enterprise & team-based development

- [Enterprise guidance](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/enterprise-guidance.md) — patterns for organizations running many SPFx solutions: shared component libraries, versioning strategy, environment promotion (dev/test/prod tenants), governance of who can deploy to the tenant app catalog.
- [Team-based development on SharePoint Framework](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/team-based-development-on-sharepoint-framework.md) — source control branching, code review, shared component conventions for multiple developers/teams working across SPFx solutions.
- [Governance considerations](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/governance-considerations.md) — organizational policy considerations (who can request AAD permissions, custom-script requirements, etc.) — surface this when the user is planning a rollout, not just writing code.
- [Hyperlinking considerations](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/hyperlinking.md) — correct handling of SharePoint-relative vs absolute links across sites/tenants.
- [Intercepting query string changes in web parts](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/intercepting-query-changes-in-webparts.md) — handling client-side routing/query-string state inside a web part hosted on a modern page.

## External libraries

- [Add an external library](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/basics/add-an-external-library.md) — how to bring in a third-party npm package correctly for the SPFx bundler.
- [Use existing JavaScript libraries](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/use-existing-javascript-libraries.md) — patterns for libraries that don't ship TypeScript types or expect global script tags.
- [Reference third-party CSS styles](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/reference-third-party-css-styles.md).
- [SharePoint Online support for Content Security Policy (CSP) / trusted script sources](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/content-securty-policy-trusted-script-sources.md) — external origins may need tenant CSP allow-listing.

## Localization

[Localize web parts](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/localize-web-parts.md) — the `mystrings.d.ts` / `*.strings` resource-file pattern SPFx scaffolds by default for locale-aware string resources; extend this pattern rather than hardcoding UI text when the user needs multi-language support.

## Tenant properties & preview capabilities

- [Tenant properties](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties.md) — reading/writing tenant-wide custom properties usable across SPFx solutions.
- [Use preview capabilities](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/try-preview-capabilities.md) — how to opt into not-yet-GA SPFx features for early testing (set expectations that these are unsupported for production use).
- [Provision SharePoint assets](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/provision-sharepoint-assets.md) — packaging and provisioning supporting SharePoint artifacts (lists, content types) alongside an SPFx solution at install time.

## Library components

For sharing pure logic (no UI) across multiple web parts/extensions without duplicating code:
- [Library component overview](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/library-component-overview.md).
- [Library component tutorial](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/library-component-tutorial.md).

## Publishing to the Microsoft 365 / SharePoint Marketplace

If the user is building an ISV/commercial SPFx app for the marketplace (not just internal deployment):
- [Publish to marketplace overview](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview.md).
- [App validation checklist](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-checklist.md) — run through this before submission; covers manifest completeness, permissions justification, branding assets, etc.
- [Common validation errors to avoid](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-common-validation-errors.md).
- [After publishing your app](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-after-publishing.md) — update/versioning process post-publish.
- FAQ: `spfx/publish-to-marketplace-faq.yml` (fetch directly if the user has marketplace-specific edge-case questions).

## Maintenance mode

[Maintenance mode for client-side web parts](https://learn.microsoft.com/en-us/sharepoint/dev/general-development/client-side-web-parts-maintenance-mode.md) — the mechanism/UX SharePoint uses to flag a deployed web part as broken/unsupported; useful context when diagnosing why a previously working web part suddenly shows a warning banner in production.

## Business process automation adjacent to SPFx

SPFx solutions frequently sit alongside **Power Automate** flows and **Power Apps** customized forms on the same lists/libraries. If the user's ask blends SPFx with flows/forms rather than pure custom code, these are the relevant docs to route to (outside the core SPFx doc tree but linked from it):
- [Power Automate + SharePoint overview](https://learn.microsoft.com/en-us/sharepoint/dev/business-apps/power-automate/sharepoint-connector-actions-triggers.md) and guidance on `Get items`/`Get files` actions, list/file permissions, the `Send an HTTP request to SharePoint` action, page approvals, document approvals, and migrating classic workflows to Power Automate.
- [Power Apps custom list forms](https://learn.microsoft.com/en-us/sharepoint/dev/business-apps/power-apps/get-started/create-your-first-custom-form.md) as an alternative/complement to an SPFx Form Customizer when the need is closer to no-code.
