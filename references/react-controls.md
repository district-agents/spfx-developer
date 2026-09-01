# @pnp/spfx-controls-react — Reusable React Controls

Source: [sp-dev-fx-controls-react docs](https://github.com/pnp/sp-dev-fx-controls-react/tree/master/docs/documentation) (`docs/index.md`, `docs/controls/*.md`, `docs/controls/charts/*.md`, `docs/controls/fields/*.md`).

**Prefer these controls over hand-rolling the same UI.** When a web part/extension needs a people picker, a list/item picker, a taxonomy picker, a file/folder picker, a rich text editor, a permission-trimmed section, an unconfigured-state placeholder, or a themed title — reach for the matching control here before writing custom Fluent UI composition from scratch. Only build custom UI when no control below fits or the design requires something materially different.

## Install

```console
npm install @pnp/spfx-controls-react --save --save-exact
```
Then register the localized strings in `config/config.json`:
```json
"localizedResources": {
  "ControlStrings": "node_modules/@pnp/spfx-controls-react/lib/loc/{locale}.js"
}
```
- **v3** (current, actively maintained) needs SPFx ≥ 1.13.\* and targets SharePoint Online only.
- **v2** — legacy, SPFx ≥ 1.11.0, no longer actively developed.
- **v1** — SharePoint on-premises (2016/2019) only; use this if the target is on-prem, not v3.
- If pairing v3 with SPFx exactly 1.12.1, the web part context must be cast to `any` before passing it to a control.
- Every control imports from its own subpath: `@pnp/spfx-controls-react/lib/<ControlName>` — never import the whole package.

## Decision guide — which control for which need

| Need | Control | Notes |
|---|---|---|
| Pick user(s)/group(s) | **PeoplePicker** | Needs `IPeoplePickerContext` (absoluteUrl, msGraphClientFactory, spHttpClient) |
| Pick item(s) from a specific list by a text column | **ListItemPicker** | Type-ahead suggestions against one column |
| Pick one or more lists/libraries on the current site | **ListPicker** | Filter by `baseTemplate`/`contentTypeId` |
| Pick a content type | **ContentTypePicker** | |
| Pick field(s) from a list or site | **FieldPicker** | |
| Pick a file from OneDrive/SharePoint/upload/web search/link | **FilePicker** | Multi-source, tabbed UI |
| Pick a folder | **FolderPicker** | |
| Browse folders/subfolders from a root | **FolderExplorer** | |
| Upload one or more files | **UploadFiles** / **DragDropFiles** | |
| Pick term(s) from a managed metadata term set | **ModernTaxonomyPicker** (preferred, REST-based, handles large term sets) or **TaxonomyPicker** (legacy, more feature parity with OOTB control) | |
| Pick a date/time | **DateTimePicker** | |
| Pick an icon | **IconPicker** | |
| Pick an image (SharePoint/OneDrive/Stock/search) | **ImagePicker** | |
| Pick a SharePoint site | **SitePicker** | |
| Pick a Team / Team channel | **TeamPicker** / **TeamChannelPicker** | |
| Pick a location on a map | **LocationPicker** | |
| Render a full CRUD-ready list/document-library form automatically from list schema | **DynamicForm** | Reads the list's fields/settings; supports file selection for doc libraries |
| Render a filterable/sortable/groupable list of items (not tied to a live SP list) | **ListView** | Supports selection, drag-drop, grouping, per-column filtering |
| Add a contextual/right-click menu to ListView rows | **ListView.ContextualMenu** | |
| Render list item comments UI (matches OOTB) | **ListItemComments** | |
| Manage list item attachments | **ListItemAttachments** | |
| Rich text editing/display | **RichText** | |
| Render field values the way modern list views do (user, lookup, date, taxonomy, url/picture, file-type icon, etc.) | **Field*Renderer** family (`fields/`) | See "Field renderer controls" below |
| Show/hide UI based on the current user's permissions | **SecurityTrimmedControl** | Works for current site/list or a remote site/list |
| Show an "unconfigured web part" call-to-action | **Placeholder** | Standard SPFx empty-state pattern |
| Editable web part title (edit mode aware) | **WebPartTitle** | Wires into `this.properties.title` + `this.displayMode` |
| Charts | **ChartControl** (Chart.js wrapper) + specific chart docs under `charts/` (Bar, Line, Pie, Doughnut, Radar, Bubble, Scatter, PolarArea) | |
| Breadcrumb of site hierarchy | **SiteBreadcrumb** | |
| Navigate/select a term from a term set (tree UI) | **TermSetNavigation** | |
| Accordion / collapsible sections | **Accordion**, **AccessibleAccordion** | |
| Carousel | **Carousel** | |
| Dialog with animation | **AnimatedDialog** | |
| Dialog/panel hosting an iframe | **IFrameDialog**, **IFramePanel** | |
| KPI tile | **KPIControl** | |
| Multi-step progress indicator | **ProgressStepsIndicator**; sequential task progress | **Progress** |
| Paginate a data set | **Pagination** | |
| Responsive grid layout for web part content | **GridLayout** | |
| Filter bar matching modern list filtering UX | **FilterBar** | |
| Monaco code editor | **MonacoEditor** | |
| Persona with live presence | **LivePersona** | |
| Theme-aware wrapper improving Teams/isolated-web-part theming | **EnhancedThemeProvider**, **VariantThemeProvider** | Use when building Teams tabs/personal apps or isolated web parts that need reliable theme propagation |
| Emoji/hover reaction bar | **HoverReactionsBar** | |
| Audio player | **ModernAudio** | |
| Tree view | **TreeView** | |
| File type icon by extension | **FileTypeIcon** | |
| List of the user's Teams | **MyTeams** | |
| Adaptive Card rendering / designer | **AdaptiveCardHost**, **AdaptiveCardDesignerHost** | Relevant for Viva Connections ACE work — see `teams-viva-copilot.md` |
| World map / generic map rendering | **Map**, **WorldMap** | |
| Native SharePoint Share dialog | **ShareDialog** | |
| Toolbar / dashboard chrome for Teams | **Toolbar**, **Dashboard** | Teams-specific chrome controls |
| View picker (choose a list view) | **ViewPicker** | |

## Usage pattern (consistent across nearly every control)

1. `npm install @pnp/spfx-controls-react --save --save-exact` (once per project).
2. Import from the control's own subpath.
3. Pass `context={this.props.context}` (or the equivalent `BaseComponentContext`) — almost every data-aware control needs this to make its own SharePoint/Graph calls.
4. Wire an `onChange`/`onSelectedItem`/`onSave` callback to receive the selected value(s).

### PeoplePicker
```typescript
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

const peoplePickerContext: IPeoplePickerContext = {
  absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
  msGraphClientFactory: this.props.context.msGraphClientFactory,
  spHttpClient: this.props.context.spHttpClient,
};

<PeoplePicker
  context={peoplePickerContext}
  titleText="People Picker"
  personSelectionLimit={3}
  groupName="Team Site Owners" // omit to search all users
  principalTypes={[PrincipalType.User]}
  required={true}
  onChange={this._getPeoplePickerItems}
/>
```
`groupId` as a `string` (M365 Group) needs Graph scope `GroupMember.Read.All` + `User.ReadBasic.All`, or `Directory.Read.All` — remember to add these via `webApiPermissionRequests` (see `apis-and-data.md`).

### ListItemPicker
```typescript
import { ListItemPicker } from '@pnp/spfx-controls-react/lib/ListItemPicker';

<ListItemPicker
  listId="da8daf15-d84f-4ab1-9800-7568f82fed3f"
  columnInternalName="Title"
  keyColumnInternalName="Id"
  filter="Title eq 'SPFx'"
  itemLimit={2}
  onSelectedItem={this.onSelectedItem}
  context={this.props.context}
/>
```

### FilePicker
```typescript
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';

<FilePicker
  accepts={[".gif", ".jpg", ".jpeg", ".png", ".svg"]}
  buttonIcon="FileImage"
  onSave={(result: IFilePickerResult[]) => this.setState({ result })}
  context={this.props.context}
/>
```

### ModernTaxonomyPicker
```typescript
import { ModernTaxonomyPicker } from "@pnp/spfx-controls-react/lib/ModernTaxonomyPicker";

<ModernTaxonomyPicker
  allowMultipleSelections={true}
  termSetId="f233d4b7-68fb-41ef-8b58-2af0bafc0d38"
  panelTitle="Select Term"
  label="Taxonomy Picker"
  context={this.props.context}
  onChange={this.onTaxPickerChange}
/>
```

### DynamicForm — auto-generate a full list/library form
```typescript
import { DynamicForm } from "@pnp/spfx-controls-react/lib/DynamicForm";

<DynamicForm
  context={this.props.context}
  listId="3071c058-549f-461d-9d73-8b9a52049a80"
  listItemId={1} // omit for a "new item" form
  onSubmitted={async (item) => console.log(item)}
  onSubmitError={(item, error) => alert(error.message)}
  onCancelled={() => console.log("cancelled")}
/>
```
Reaches for this before building a custom property-pane-driven form when the requirement is "let users create/edit a list item" — it saves rebuilding field-type-aware rendering (people, lookup, taxonomy, choice, etc.) by hand.

### ListView — render/filter/group a set of items
```typescript
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";

<ListView
  items={items}
  viewFields={viewFields}
  iconFieldName="FileRef"
  selectionMode={SelectionMode.multiple}
  selection={this._getSelection}
  showFilter={true}
  groupByFields={groupByFields}
/>
```
`items` here is a plain array you already fetched (e.g., via PnPjs) — `ListView` is a presentation control, not a data-fetching one. Filtering supports `<ColumnName>:<value>` syntax for column-scoped search.

### Placeholder — unconfigured web part state
```typescript
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

<Placeholder
  iconName="Edit"
  iconText="Configure your web part"
  description="Please configure the web part."
  buttonLabel="Configure"
  onConfigure={() => this.props.context.propertyPane.open()}
/>
```
Use this as the default render output whenever a required property pane setting (e.g., a list selection) hasn't been set yet, rather than rendering a blank or broken component.

### WebPartTitle — editable title matching OOTB behavior
```typescript
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";

<WebPartTitle
  displayMode={this.props.displayMode}
  title={this.props.title}
  updateProperty={this.props.updateProperty}
/>
```
Requires passing `title`, `displayMode` (from `this.displayMode` in the web part class), and an `updateProperty` callback (that sets `this.properties.title = value`) down from the web part's `render()`.

### SecurityTrimmedControl — permission-based UI gating
```typescript
import { SecurityTrimmedControl, PermissionLevel } from "@pnp/spfx-controls-react/lib/SecurityTrimmedControl";
import { SPPermission } from '@microsoft/sp-page-context';

<SecurityTrimmedControl
  context={this.props.context}
  level={PermissionLevel.currentWeb}
  permissions={[SPPermission.viewPages]}
>
  {/* rendered only if the current user has the permission(s) */}
</SecurityTrimmedControl>
```
Also supports `PermissionLevel.currentList`, `.remoteWeb`, and `.remoteListOrLib` (with `remoteSiteUrl`/`relativeLibOrListUrl`).

## Field renderer controls (`fields/`)

For building a **custom Field Customizer** or any UI that needs to render column values the same way modern SharePoint list views do: `FieldTextRenderer`, `FieldDateRenderer`, `FieldUserRenderer`, `FieldLookupRenderer`, `FieldTaxonomyRenderer`, `FieldUrlRenderer` (hyperlink or picture), `FieldFileTypeRenderer` (doc/folder icon), `FieldNameRenderer` (linked file/item name), `FieldTitleRenderer`, `FieldAttachmentsRenderer` (paperclip + count). Start with `fields/main.md` (`FieldRendererHelper`) which picks the right renderer for a given field type automatically — usually less work than choosing renderers manually.

## Telemetry

All controls report anonymous usage telemetry (control name + related metadata only). Opt out per-solution if required:
```typescript
import PnPTelemetry from "@pnp/telemetry-js";
PnPTelemetry.getInstance().optOut();
```
Mention this to the user if they're building for a privacy-sensitive tenant/customer.

## Migrating from v1

If the user has an older solution on v1 (on-prem-era) moving to v3 (SPO), point them at the dedicated [Migrating from V1](https://github.com/pnp/sp-dev-fx-controls-react/blob/master/docs/documentation/docs/guides/migrate-from-v1.md) guide before assuming prop shapes are unchanged — several controls changed their prop signatures between v1 and v2/v3.
