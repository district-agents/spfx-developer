# Microsoft Teams, Viva Connections & Copilot Apps with SPFx

Source: [Build for Teams overview](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview), [Viva Connections overview](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/overview-viva-connections), [Copilot Apps overview](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/copilot/overview-copilot-apps), [Office/Outlook overview](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/office/overview).

## Microsoft Teams

SPFx web parts can be exposed directly as Teams tabs/personal apps — same code, no rewrite:
- Set `supportedHosts: ["TeamsTab", "TeamsPersonalApp", ...]` in the web part manifest.
- [Expose web parts in Microsoft Teams](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-expose-webparts-teams.md) / [configure web parts in Teams](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-configure-in-teams.md).
- `this.context.sdks.microsoftTeams` is populated at runtime when hosted in Teams — branch UI/behavior on its presence.
- Scenario walkthrough: [Build a "Me-experience" in Microsoft Teams](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-me-experience.md).
- [Considerations](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-considerations.md) — Teams-specific caveats (theming differences, host restrictions).
- Deployment specifics: [Deployment options for SPFx solutions & Microsoft Teams](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/deployment-spfx-teams-solutions.md) — packaging a Teams app manifest around the SPFx component, publishing via Teams admin center or the org app catalog.
- Also: [Hosting Microsoft Teams Tabs in SharePoint](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/using-teams-solutions-in-sharepoint.md) (the inverse — a Teams-first tab surfaced back into SharePoint), and [Build meeting apps for Microsoft Teams with SPFx](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-meeting-app.md).

## Microsoft Viva Connections — Adaptive Card Extensions (ACE)

Viva Connections' dashboard relies **entirely** on SPFx for extensibility — ACEs are the only customization mechanism (desktop and mobile both).

- Component type: **Adaptive Card Extension**, distinct from web parts/extensions — scaffolded via the Yeoman generator's Copilot/ACE option or `yo @microsoft/sharepoint` → ACE flow.
- Two primary views per ACE: **Card View** (compact dashboard tile) and **Quick View** (expanded detail panel) — see [Card View Design](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/design/designing-card.md) and [Quick View Design](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/design/designing-quick-view.md) (with [Quick View Samples](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/design/quick-view-samples.md)).
- Tutorial path: [Build your first SharePoint Adaptive Card Extension](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/get-started/build-first-sharepoint-adaptive-card-extension.md) → [Advanced Card View functionality](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/get-started/advanced-card-view-functionality.md) → [Advanced Quick View functionality](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/get-started/advanced-quick-view-functionality.md).
- Feature-specific tutorials: geolocation actions ([tutorial](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/get-started/actions/geolocation/GeolocationTutorial.md), [property pane](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/get-started/actions/geolocation/GeolocationPropertyPane.md)), media upload actions ([tutorial](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/get-started/actions/media-upload/MediaUploadTutorial.md), [property pane](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/get-started/actions/media-upload/MediaUploadPropertyPane.md)), [People Search ACE](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/get-started/build-people-search-adaptive-card-extension.md), [Data Visualization ACE](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/get-started/build-data-visualization-adaptive-card-extension.md), [HTML Quick View ACE](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/get-started/build-html-quickview-adaptive-card-extension.md).
- The **Focus feature** (attention-drawing UI pattern for cards) has its own [documentation](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/features/focus-feature/FocusFeatureDocumentation.md) and [tutorial](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/features/focus-feature/FocusFeatureTutorial.md).
- [Card Designer advanced API features](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/features/card-designer/card-designer-api-support.md).
- Known limitation: [Adaptive Card Extensions iconography limitations](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/get-started/fluent-icons-limitations.md) (not all Fluent icon glyphs are supported in cards).
- [Making Quick View compatible with dark mode](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/get-started/making-quickview-compatable-darkmode-mobile.md).
- [ACEs and Teams apps](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/get-started/adaptive-card-extensions-and-teams.md) — how ACEs relate to Teams app packaging.
- Upgrading older ACEs: [Migrate existing ACEs to SPFx v1.18.0](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/get-started/migrate-to-spfx-1-18.md).
- Check [Known issues](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/known-issues-adaptive-cards.md) before debugging odd ACE rendering.

## Copilot Apps

Newer SPFx component type for building agent/Copilot-surfaced experiences.
- [Overview of Copilot Apps](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/copilot/overview-copilot-apps.md).
- [Build your first Copilot App](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/copilot/get-started/build-your-first-copilot-app.md).
- [Display Mode](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/copilot/displayMode.md) — controls how the app's UI is presented within a Copilot surface.
- Because this is a newer/actively evolving area, **verify current API shape with a fresh Microsoft Learn lookup** before giving code-level guidance — don't rely solely on this summary.

## Outlook, Office & other Microsoft 365 hosts

- [Teams apps for SharePoint, Teams, Outlook, and Office (overview)](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/office/overview.md) — the same SPFx component can, in principle, be surfaced across multiple M365 hosts ("write once" story referenced in the SPFx overview, with the `contoso-retail-demo` PnP reference solution as a worked example).
- [Extend Outlook and Office](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/office/overview) is the entry point the main SPFx overview page links to for this scenario — check it directly for current specifics, as host support surface area changes over time.

## When to reach for the Microsoft Agents Toolkit / SPFx Toolkit instead of raw Yeoman

The main SPFx overview mentions three ways to scaffold SPFx-powered solutions:
- **Yeoman generator** (`yo @microsoft/sharepoint`) — the default, framework-agnostic path covered throughout this skill.
- **[Microsoft Agents Toolkit](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/agents-toolkit-fundamentals)** (VS Code extension) — when the target is primarily Teams/agents-centric development.
- **[SharePoint Framework Toolkit](https://pnp.github.io/vscode-viva/)** (community VS Code extension, formerly "Viva Connections Toolkit") — adds SPFx-specific VS Code tooling, especially useful for ACE/Viva Connections work.
