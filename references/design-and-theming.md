# Design, Theming & Fluent UI in SPFx

Source: [Design guidance overview](https://learn.microsoft.com/en-us/sharepoint/dev/design/design-guidance-overview), [Use theme colors](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-theme-colors-in-your-customizations), [Fluent UI integration](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/fluent-ui-integration), [Work with CSS](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/css-recommendations), [Accessibility](https://learn.microsoft.com/en-us/sharepoint/dev/design/accessibility).

## Theme tokens (never hardcode colors)

SPFx ships `@microsoft/load-themed-styles`, which scans CSS/SCSS for theme-slot syntax and swaps in the current site theme's color at runtime, falling back to the given default if the token isn't defined:

```scss
// _theming.scss or component .module.scss
$themePrimary: '[theme:themePrimary, default:#0078d4]';

.myButton {
  background-color: $themePrimary;
}
```
- Always include the `default:` fallback — themes may not define every token.
- New scaffolded web parts default to a fixed blue palette in generated SCSS; **update them to theme tokens** so the component matches whatever theme/palette the site actually uses (red, green, custom, etc.).
- Full token list and where each is used across SharePoint chrome is in [Use theme colors in your SharePoint Framework customizations](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-theme-colors-in-your-customizations.md) — reference it directly rather than guessing token names.

## `ThemeProvider` — theme-variant-aware components (section backgrounds)

Since SPFx v1.8, each page section can use one of **4 theme variants**: main, neutral, soft, strong. A theme-aware component should react to the variant of the section it's placed in:

```typescript
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base';

private _themeProvider: ThemeProvider;
private _themeVariant: IReadonlyTheme | undefined;

protected onInit(): Promise<void> {
  this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
  this._themeVariant = this._themeProvider.tryGetTheme();
  this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
  return super.onInit();
}
```
Pass `this._themeVariant.semanticColors` (an `ISemanticColors` object — background, body text, buttons, borders, etc.) down into your rendering (e.g., as a React prop) instead of hardcoded CSS colors. See [Supporting section backgrounds](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/supporting-section-backgrounds.md) and [Semantic slots](https://learn.microsoft.com/en-us/sharepoint/dev/design/semantic_slots.md).

## Fluent UI

- Use **Fluent UI React** components for anything UI-heavy so the component visually matches the rest of Microsoft 365 — buttons, panels, DetailsList, Persona, etc. Referenced throughout the docs as the standard component library for SPFx web parts and extensions.
- [Fluent UI integration](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/fluent-ui-integration.md) covers version considerations and how SPFx wires Fluent UI styling into the theme system.
- [Use Office UI Fabric React components](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/use-fabric-react-components.md) — the older-named tutorial (Office UI Fabric was Fluent UI's previous name); still the relevant walkthrough for wiring it into a web part.
- [Using Brand center fonts](https://learn.microsoft.com/en-us/sharepoint/dev/design/use-brand-center-fonts-in-spfx-components.md) — pull the org's configured brand fonts into a component.

## CSS/SASS recommendations

[Work with CSS](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/css-recommendations.md):
- Scope styles to the component (CSS Modules — SPFx generates `*.module.scss` → typed `styles` object) to avoid leaking into/from the host page.
- Don't rely on the page's own CSS classes/DOM structure (unsupported, breaks on any SharePoint UI update).
- [Reference third-party CSS styles](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/reference-third-party-css-styles.md) — how to safely bring in an external stylesheet/library's CSS without polluting the host page.
- [Configure SASS processing](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/configure-sass.md) if custom SASS options are needed.

## Design system guidance (`design/` docs)

These cover UX patterns Microsoft expects for well-behaved web parts, useful when the user asks "how should this look/behave":
- [Design a web part](https://learn.microsoft.com/en-us/sharepoint/dev/design/design-a-web-part.md) (parent overview) →
  - Reactive vs nonreactive web parts (property pane update model).
  - Responsive/grid design.
  - Titles & descriptions conventions.
  - Commanding within a web part (top actions/menus).
  - Web part layout patterns.
  - "Web part levels" (complexity tiers Microsoft defines for web parts).
  - UI text guidance (tone, capitalization, terminology).
  - Placeholders & fallbacks (loading/empty/error states).
  - **Empty states** — dedicated guidance for what to show when a web part has no data yet.
- [Key web part examples](https://learn.microsoft.com/en-us/sharepoint/dev/design/key-web-part-examples.md) / [Web part showcase](https://learn.microsoft.com/en-us/sharepoint/dev/design/showcase-web-part.md) — reference implementations.
- [Designing a web part icon](https://learn.microsoft.com/en-us/sharepoint/dev/design/designing-a-web-part-icon.md).
- [Authoring pages](https://learn.microsoft.com/en-us/sharepoint/dev/design/authoring-pages.md) — page-composition guidance for solutions that provision whole pages.
- [Design considerations for web parts](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/basics/design-considerations-for-web-parts.md) (SPFx-specific companion to the general design docs).

## Accessibility

[Accessibility](https://learn.microsoft.com/en-us/sharepoint/dev/design/accessibility.md) — keyboard navigation, ARIA roles/labels, color contrast, screen-reader behavior. Bring this up proactively whenever building interactive controls (custom dropdowns, modals, drag targets) rather than only when asked — Fluent UI components handle most of this automatically, custom HTML does not.

## Image helper & Graph Toolkit

- [Image Helper API](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/image-helper-api.md) — SPFx's built-in helper for cropping/positioning images consistently with SharePoint's own image treatment.
- [Microsoft Graph Toolkit](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-microsoft-graph-toolkit.md) — prebuilt, theme-aware web components for common M365-connected UI (people picker, person card, agenda, file list) usable inside SPFx to avoid rebuilding common Graph UI from scratch.
