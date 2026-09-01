# spfx-developer
[![skills.sh](https://skills.sh/b/district-agents/spfx-developer)](https://skills.sh/district-agents/spfx-developer)

An AI-agent **skill** for developing, debugging, packaging, and deploying [SharePoint Framework (SPFx)](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview) solutions — web parts, all five extension types, Viva Connections Adaptive Card Extensions, Copilot Apps, and Teams/Outlook/Office apps built on SPFx.

It bakes in an opinionated, current stack:
- **[PnPjs](https://pnp.github.io/pnpjs/) (`@pnp/sp` / `@pnp/graph`) v4** for SharePoint and Microsoft Graph data access
- **[@pnp/spfx-controls-react](https://pnp.github.io/sp-dev-fx-controls-react/)** for canvas/body UI (pickers, forms, list views, placeholders, charts, ...)
- **[@pnp/spfx-property-controls](https://pnp.github.io/sp-dev-fx-property-controls/)** for property pane fields (list/site/term/color/date/people pickers, collection-data editor, ...)

so an agent reaches for these by default instead of hand-rolling REST calls or rebuilding common controls from scratch.

## Why this exists

General-purpose AI agents tend to give stale or blended SPFx guidance — mixing the legacy Gulp toolchain with the current Heft one, mixing PnPjs v3 and v4 syntax, or not knowing the PnP community controls exist at all. This skill is a curated, current reference distilled from the primary sources so an agent gives version-correct, idiomatic answers instead.

## Sources

Everything in this skill is derived directly from the official documentation, not from general training knowledge:

| Area | Source |
|---|---|
| SPFx core (web parts, extensions, toolchain, deployment, design, Teams/Viva/Copilot, platform support, governance) | [learn.microsoft.com/sharepoint/dev/spfx](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview) and the [`SharePoint/sp-dev-docs`](https://github.com/SharePoint/sp-dev-docs) GitHub source |
| PnPjs (`@pnp/sp`, `@pnp/graph`) | [`pnp/pnpjs`, `version-4` branch, `/docs`](https://github.com/pnp/pnpjs/tree/version-4/docs) |
| `@pnp/spfx-controls-react` | [`pnp/sp-dev-fx-controls-react`, `/docs/documentation`](https://github.com/pnp/sp-dev-fx-controls-react/tree/master/docs/documentation) |
| `@pnp/spfx-property-controls` | [`pnp/sp-dev-fx-property-controls`, `/docs/documentation`](https://github.com/pnp/sp-dev-fx-property-controls/tree/master/docs/documentation) |

Snapshot current as of **SPFx v1.24 (August 2026)**. SPFx, PnPjs, and the PnP controls libraries all ship regularly — treat this skill as a strong, well-organized starting point and re-verify version-sensitive specifics (exact Node/SPFx compatibility numbers, latest release notes) against the live docs for anything that matters for a production decision.

## Structure

```
spfx-developer/
├── SKILL.md                              orientation, decision tree, quick-start workflow, cheat sheets, pitfalls
└── references/
    ├── web-parts.md                      web part anatomy, property pane, theming hooks, Teams tabs, legacy migration
    ├── extensions.md                     all 5 extension types, manifests, tenant-wide deployment, custom dialogs
    ├── pnpjs.md                          @pnp/sp / @pnp/graph v4: setup, CRUD, batching, error handling, v3→v4 migration
    ├── react-controls.md                 @pnp/spfx-controls-react: canvas/body UI decision guide + usage patterns
    ├── property-controls.md              @pnp/spfx-property-controls: property pane field decision guide + usage patterns
    ├── apis-and-data.md                  native SPHttpClient/AadHttpClient/MSGraphClient, webhooks, permission approval
    ├── toolchain-and-deployment.md       Heft vs Gulp, packaging, app catalog deployment, CI/CD, debugging, CSP
    ├── design-and-theming.md             Fluent UI, theme tokens, ThemeProvider, accessibility, empty states
    ├── teams-viva-copilot.md             Teams tabs, Viva Connections ACEs, Copilot Apps, Outlook/Office
    ├── platforms-and-versions.md         SPFx version history, Node.js compatibility, on-prem support ceilings
    └── governance-and-guidance.md        enterprise rollout, external libraries, localization, marketplace publishing
```

`SKILL.md` is intentionally kept short (~100 lines) and orienting; it links out to the `references/` files, which the agent is expected to open only when the task actually needs that depth (progressive disclosure) rather than loading the entire skill on every turn.

## Usage

This follows the [Claude Skills](https://docs.claude.com/en/docs/agents-and-tools/agent-skills/overview) format — a folder with a `SKILL.md` (YAML frontmatter + Markdown body) plus supporting reference files. To use it:

1. **Claude (claude.ai / Claude Code / Claude Cowork / API)**: add the `spfx-developer/` folder as a skill. See [Anthropic's skills documentation](https://docs.claude.com/en/docs/agents-and-tools/agent-skills/overview) for how to install a skill for your surface.
2. **Other agent tooling**: the folder is plain Markdown with no Claude-specific syntax — point any agent that supports file/context loading at `SKILL.md` as the entry point, with `references/*.md` as supplementary docs it can load on demand.

## Scope notes for anyone extending this

- **PnPjs**: this covers `@pnp/sp` in depth per the intended use case (SharePoint on Microsoft 365); `@pnp/graph` setup is covered where it intersects (auth context, common patterns) but not exhaustively — extend `references/pnpjs.md` if your solutions lean heavily on Graph.
- **On-premises SharePoint (2016/2019/Subscription Edition)**: covered at the level of "what's different and where the ceiling is" (`references/platforms-and-versions.md`), not as a fully first-class parallel track — PnPjs v4 and the current PnP controls libraries target SharePoint Online only (on-prem needs PnPjs v2 and controls v1).
- **Copilot Apps**: this is the newest, fastest-moving SPFx component type; `references/teams-viva-copilot.md` flags it explicitly as an area to re-verify against live docs before giving code-level guidance.

## Contributing / updating

If you fork or extend this: keep the "read the reference file before giving code-level guidance" pattern in `SKILL.md`, keep each `references/*.md` scoped to one coherent area, and update the source table above if you add material derived from a new upstream doc set.
