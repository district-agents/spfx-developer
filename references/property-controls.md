# @pnp/spfx-property-controls — Reusable Property Pane Controls

Source: [sp-dev-fx-property-controls docs](https://github.com/pnp/sp-dev-fx-property-controls/tree/master/docs/documentation) (`docs/index.md`, `docs/getting-started.md`, `docs/controls/*.md`).

**Prefer these over the built-in `PropertyPaneTextField`/`PropertyPaneDropdown`/etc. whenever the property pane needs something richer** — a list/site/term/color/date picker, a people picker, a collection-data editor, a code editor, or a callout-annotated field. This is the sibling library to `@pnp/spfx-controls-react` (see `references/react-controls.md`) but specifically for **property pane fields**, not canvas/body UI — don't confuse the two. A property-pane need (configuring the web part) → this file. An on-page/body UI need (what the web part renders for end users) → `react-controls.md`.

## Install

```console
npm install @pnp/spfx-property-controls --save --save-exact
```
Then register the localized strings in `config/config.json`:
```json
"localizedResources": {
  "PropertyControlStrings": "./node_modules/@pnp/spfx-property-controls/lib/loc/{locale}.js"
}
```
- **v3** (current, actively maintained) needs SPFx ≥ 1.13.\*, SharePoint Online only.
- **v2** — deprecated, SPFx ≥ 1.11.0.
- **v1** — SharePoint on-premises (2016/2019) only.
- With v3 on SPFx exactly 1.12.1, cast the web part context to `any` before passing it to a control (same caveat as `@pnp/spfx-controls-react`).
- Every control imports from its own subpath: `@pnp/spfx-property-controls/lib/<ControlName>` — never import the whole package.

## The universal usage shape

Unlike the body-UI controls in `react-controls.md` (which are JSX components), **every control here is a factory function** you call inside `getPropertyPaneConfiguration()` and place directly into a `groupFields` array — it returns an `IPropertyPaneField`, the same shape as the built-in `PropertyPaneTextField`, etc.

```typescript
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages: [{
      header: { description: "Settings" },
      groups: [{
        groupName: "General",
        groupFields: [
          PropertyFieldListPicker('listId', {
            label: 'Select a list',
            selectedList: this.properties.listId,
            includeHidden: false,
            orderBy: PropertyFieldListPickerOrderBy.Title,
            disabled: false,
            onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
            properties: this.properties,
            context: this.context,
            onGetErrorMessage: null,
            deferredValidationTime: 0,
            key: 'listPickerFieldId',
          }),
        ],
      }],
    }],
  };
}
```
Recurring required options across almost every control:
- **First positional argument**: the target property name on `this.properties` (e.g. `'listId'`).
- `key` — a unique string identifying this pane control instance (not the same as the property name; convention is `<name>FieldId`).
- `onPropertyChange` — pass `this.onPropertyPaneFieldChanged.bind(this)` (or `this.onPropertyPaneFieldChanged` if already bound) so the property pane re-renders and the web part updates reactively.
- `properties` — pass `this.properties` so the control can write the new value back onto it.
- `context` — pass `this.context`, required by any control that makes its own SharePoint/Graph calls (list/site/people/term pickers).
- `onGetErrorMessage` / `deferredValidationTime` — standard property-pane validation hook, same contract as native fields; see `references/web-parts.md`'s property-pane validation note.

## Decision guide — which control for which need

| Need | Control |
|---|---|
| Pick a SharePoint list/library on the current (or another) site | **PropertyFieldListPicker** |
| Pick a view of a list | **PropertyFieldViewPicker** |
| Pick column(s) from a list | **PropertyFieldColumnPicker** |
| Pick a content type | **PropertyFieldContentTypePicker** |
| Pick a SharePoint site | **PropertyFieldSitePicker** |
| Pick a folder | **PropertyFieldFolderPicker** |
| Pick a file (OneDrive/SharePoint/upload/web search/link) | **PropertyFieldFilePicker** |
| Pick user(s)/group(s) | **PropertyFieldPeoplePicker** |
| Pick a Microsoft Team | **PropertyFieldTeamPicker** |
| Pick managed-metadata term(s) | **PropertyFieldTermPicker** (standard) or **PropertyFieldEnterpriseTermPicker** (enterprise/cross-tenant scenarios) |
| Pick a SharePoint role definition (permission level) | **PropertyFieldRoleDefinitionPicker** |
| Pick a date/time | **PropertyFieldDateTimePicker** |
| Pick a color (full picker) | **PropertyFieldColorPicker** |
| Pick a color from a fixed swatch set | **PropertyFieldSwatchColorPicker** |
| Pick a theme/brand palette | **PropertyPanePalettePicker** |
| Pick a Brand Center font | **PropertyFieldBrandFontPicker** |
| Pick an icon | **PropertyFieldIconPicker** |
| Enter/edit a repeating table of structured data (e.g. multiple locations, links, rows) | **PropertyFieldCollectionData** |
| Multi-select dropdown | **PropertyFieldMultiSelect** |
| Reorderable list of items | **PropertyFieldOrder** |
| Numeric input | **PropertyFieldNumber** |
| Spin button (numeric with +/- steppers) | **PropertyFieldSpinButton** |
| Loading/spinner indicator inside the pane | **PropertyFieldSpinner** |
| Free-text search box | **PropertyFieldSearch** |
| Password-style masked input | **PropertyFieldPassword** |
| GUID input/generator | **PropertyFieldGuid** |
| Code snippet editor (syntax highlighted) | **PropertyFieldCodeEditor** |
| Full Monaco editor (VS Code's editor) | **PropertyFieldMonacoEditor** |
| Editable data grid | **PropertyFieldGrid** |
| A styled message/banner inside the pane (info/warning/error/success) | **PropertyFieldMessage** |
| Rich descriptive/help content authored in Markdown | **PropertyPaneMarkdownContent** |
| Description + "read more" link + optional embedded video, standard info panel | **PropertyPaneWebPartInformation** |
| Raw JSON edit / export / import of all web part properties | **PropertyPanePropertyEditor** |
| A styled button | **PropertyFieldButton** |
| Any of the above native-style fields (text/checkbox/choice-group/dropdown/link/label/slider/toggle) but **with an attached info callout** | **PropertyField<X>WithCallout** family — `PropertyFieldButtonWithCallout`, `PropertyFieldCheckboxWithCallout`, `PropertyFieldChoiceGroupWithCallout`, `PropertyFieldDropdownWithCallout`, `PropertyFieldLabelWithCallout`, `PropertyFieldLinkWithCallout`, `PropertyFieldSliderWithCallout`, `PropertyFieldTextWithCallout`, `PropertyFieldToggleWithCallout` |

## Usage examples for the most common controls

### PropertyFieldPeoplePicker
```typescript
import { PropertyFieldPeoplePicker, PrincipalType, IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

export interface IMyWebPartProps { people: IPropertyFieldGroupOrPerson[]; }

PropertyFieldPeoplePicker('people', {
  label: 'Select people',
  initialData: this.properties.people,
  allowDuplicate: false,
  principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
  onPropertyChange: this.onPropertyPaneFieldChanged,
  context: this.context,
  properties: this.properties,
  key: 'peopleFieldId',
})
```

### PropertyFieldCollectionData — repeating structured rows
```typescript
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

PropertyFieldCollectionData("collectionData", {
  key: "collectionData",
  label: "Collection data",
  panelHeader: "Manage rows",
  manageBtnLabel: "Manage collection data",
  value: this.properties.collectionData,
  fields: [
    { id: "Title", title: "Title", type: CustomCollectionFieldType.string, required: true },
    { id: "Age", title: "Age", type: CustomCollectionFieldType.number },
    { id: "Sign", title: "Confirmed", type: CustomCollectionFieldType.boolean },
  ],
  onPropertyChange: this.onPropertyPaneFieldChanged,
  properties: this.properties,
})
```
Result is a plain array of objects matching the field definitions — good fit whenever a web part needs "a user-editable list of N similar things" (locations, links, quotes, FAQ entries) configured entirely in the property pane rather than via a separate list.

### PropertyFieldColorPicker
```typescript
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

PropertyFieldColorPicker('color', {
  label: 'Color',
  selectedColor: this.properties.color,
  onPropertyChange: this.onPropertyPaneFieldChanged,
  properties: this.properties,
  style: PropertyFieldColorPickerStyle.Full,
  key: 'colorFieldId',
})
```

### PropertyFieldTermPicker
```typescript
import { PropertyFieldTermPicker, IPickerTerms } from '@pnp/spfx-property-controls/lib/PropertyFieldTermPicker';

export interface IMyWebPartProps { terms: IPickerTerms; }

PropertyFieldTermPicker('terms', {
  label: 'Select terms',
  panelTitle: 'Select terms',
  initialValues: this.properties.terms,
  onPropertyChange: this.onPropertyPaneFieldChanged,
  properties: this.properties,
  context: this.context,
  key: 'termPickerFieldId',
})
```
Note: `PropertyFieldTermPicker` relies on the legacy `ProcessQuery` (CSOM-over-HTTP) endpoint under the hood, not modern REST — functional but worth knowing when debugging odd network calls. If the user needs the modern REST-based experience, point them at the **ModernTaxonomyPicker** body-UI control in `references/react-controls.md` instead (note: that one is for canvas UI, not property-pane placement — there isn't yet a modern-REST property-pane term picker equivalent in this library as of this doc snapshot).

### PropertyFieldMessage — inline banner in the pane
```typescript
import { PropertyFieldMessage } from '@pnp/spfx-property-controls/lib/PropertyFieldMessage';
import { MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

PropertyFieldMessage("", {
  key: "MessageKey",
  text: "This list has no items yet.",
  messageType: MessageBarType.warning,
  isVisible: true,
})
```
Useful for surfacing validation/state feedback (e.g., "no list selected yet") directly in the pane rather than only in the web part body.

### PropertyPaneWebPartInformation — standard info/help panel
```typescript
import { PropertyPaneWebPartInformation } from '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation';

PropertyPaneWebPartInformation({
  description: `This is a <strong>demo web part</strong>.`,
  moreInfoLink: `https://contoso.sharepoint.com/help`,
  videoProperties: { embedLink: `https://www.youtube.com/embed/...`, properties: { allowFullScreen: true } },
  key: 'webPartInfoId',
})
```

### PropertyPanePropertyEditor — raw JSON export/import of properties
```typescript
import { PropertyPanePropertyEditor } from '@pnp/spfx-property-controls/lib/PropertyPanePropertyEditor';

PropertyPanePropertyEditor({
  webpart: this, // pass the web part instance itself
  key: 'propertyEditor',
})
```
Adds an "export"/"import" affordance (downloads/uploads `webpartproperties.json`) — handy for admins copying a fully-configured web part's settings to another page/site without rebuilding it by hand. Good to suggest when a user wants configuration portability.

## Telemetry

Same mechanism as `@pnp/spfx-controls-react` — anonymous usage telemetry (control name + related metadata). Opt out per solution:
```typescript
import PnPTelemetry from "@pnp/telemetry-js";
PnPTelemetry.getInstance().optOut();
```
If the user already opted out for `@pnp/spfx-controls-react`, mention this needs to be done once — the same `PnPTelemetry` singleton instance covers both packages, so a single `optOut()` call anywhere in the solution's startup path is sufficient.

## Migrating from v1

If the user has an older on-prem-era (v1) solution moving to v3 (SPO), point them at the dedicated [Migrating from V1](https://github.com/pnp/sp-dev-fx-property-controls/blob/master/docs/documentation/docs/guides/migrate-from-v1.md) guide — several controls changed their option shapes between v1 and v2/v3, same caveat as the `react-controls.md` sibling library.
