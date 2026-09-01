# SPFx Platform Support, Versions & Compatibility

Source: [Supported extensibility platforms overview](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/supported-extensibility-platforms-overview), [Compatibility reference](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/compatibility), [SharePoint Server 2019 & SE support](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-2019-and-subscription-edition-support), [SharePoint Server 2016 support](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-2016-support), [Roadmap](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/roadmap).

**Always verify exact current numbers via a fresh Microsoft Learn lookup before quoting a specific SPFx/Node version pairing to the user** — this table summarizes the pattern as of the docs snapshot used to build this skill (through SPFx v1.24, Aug 2026) but Microsoft ships new SPFx versions regularly.

## Supported platforms at a glance

| Platform | SPFx support |
|---|---|
| **SharePoint Online (SPO)** | Always the broadest, most current support; SPO always runs the latest SPFx. This is the primary/recommended target. |
| **SharePoint Server Subscription Edition (SE)** | Same SPFx ceiling as SharePoint Server 2019. |
| **SharePoint Server 2019** | SPFx up to **v1.4.1** only. |
| **SharePoint Server 2016** | SPFx up to **v1.1.0** only (introduced in Feature Pack 2). |
| **Microsoft Teams** | Full support via web parts (as tabs) and Copilot Apps. |
| **Microsoft Viva Connections** | Entirely dependent on SPFx (ACEs) — no other extensibility path exists, desktop and mobile alike. |

**Implication for the user's questions**: if they mention on-prem SharePoint (2016/2019/SE), immediately narrow the SPFx feature set you recommend — no AadHttpClient/MSGraphClient beyond v1.4.1's capabilities, no Heft toolchain (Heft only exists from v1.22+, far beyond on-prem's ceiling — on-prem projects always use the legacy Gulp toolchain), and check the dedicated on-prem support pages for Node.js/npm version pinning quirks specific to that SPFx version (e.g., historically SPFx v1.4.1 needed workarounds like pinning `graceful-fs`/`sass` to run on newer Node versions than officially listed).

## Toolchain generations

| SPFx version range | Toolchain | Node.js |
|---|---|---|
| v1.0 – v1.21.1 | Gulp-based (legacy) | Version pinned per release; confirm per-version |
| **v1.22.0+** | **Heft-based** (RushStack) | v22 LTS as of the v1.22 line |

The Heft transition (SPFx v1.22, Dec 2025) is a **build-system change**, not a change to the deployed runtime artifact — webpack is still the bundler either way; only how it's invoked/configured changed. `gulpfile.js` customizations don't carry over — see the migration guide referenced in `toolchain-and-deployment.md`.

## Compatibility research checklist

Because SPFx pins tightly to specific Node.js LTS releases (mismatches throw a hard error at build time, e.g. *"...does not meet the requirements for running this tool"*), when the user reports a build/tooling error:
1. Ask (or infer from `package.json`) the exact `@microsoft/sp-core-library` / `@microsoft/generator-sharepoint` version.
2. Check `node --version` against what that SPFx line requires.
3. Cross-reference the **[Platform & platform dependency compatibility reference](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/compatibility)** — it's the authoritative, actively maintained table of SPFx version ↔ Node.js version ↔ TypeScript version ↔ React version. Search/fetch it fresh rather than relying on memorized numbers, since this table changes with every SPFx release.
4. Also check the specific version's **release notes** page (`spfx/release-<version>.md`) for that release's dependency updates and breaking changes — there is a dedicated release-notes page per SPFx version going back to v1.0.0 (Feb 2017).

## Release notes / roadmap

The [Roadmap](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/roadmap.md) page indexes every SPFx release since GA (Feb 2017) with links to per-version release notes. Recent major milestones (verify currency before citing to the user):
- **v1.22.0** (Dec 2025) — Heft-based toolchain became the default for new projects.
- **v1.21.x** (Apr/May 2025) — Node.js 22 LTS support, TypeScript 5.x.
- Ongoing point releases continue through v1.23.x and v1.24.x — **search Microsoft Learn for the latest release notes rather than assuming which is "current."**

For learning-path style onboarding, Microsoft also publishes structured training: [Extend SharePoint with SPFx – Learning Path](https://learn.microsoft.com/en-us/training/paths/m365-sharepoint-associate), [Extend Viva Connections – Learning Path](https://learn.microsoft.com/en-us/training/paths/m365-extend-viva-connections), and the [Create ACEs for Viva Connections module](https://learn.microsoft.com/en-us/training/modules/sharepoint-spfx-adaptive-card-extension-card-types).

## Migrating from SharePoint Add-ins

SharePoint Add-ins (the older SharePoint-hosted/provider-hosted add-in model) are in retirement; **SPFx is the primary replacement technology** and is explicitly unaffected by the add-in retirement. If the user is planning a migration off Add-ins, point them to SPFx web parts/extensions as the target architecture and flag that Add-in-specific APIs (cross-domain library, add-in webs, ACS-based auth) have no direct SPFx equivalent — they map to modern REST/Graph + AadHttpClient patterns instead (`apis-and-data.md`).
