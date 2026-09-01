# PnPjs v4 (@pnp/sp) — Preferred Data Access Library

Source: [PnPjs docs, version-4 branch](https://github.com/pnp/pnpjs/tree/version-4/docs) (`getting-started.md`, `concepts/*`, `sp/*`, `queryable/*`, `transition-guide.md`).

**This is the default, preferred way for this skill's agent to talk to SharePoint.** Reach for the raw `SPHttpClient` (see `apis-and-data.md`) only for a single trivial one-off GET, or when the user explicitly asks to avoid adding a dependency. For anything with CRUD, filtering, batching, or more than one call, use PnPjs.

## Install & minimum requirements

```console
npm install @pnp/sp @pnp/graph --save
```
- Node.js ≥ 18, TypeScript 4.x, **ESM-only** package output.
- **PnPjs v4 requires SPFx ≥ 1.18.0** (which itself requires Node 18). For SPFx 1.15.0–1.17.4 or 1.12.1–1.14.0, use **PnPjs v3** instead — v4 will not work. For SharePoint on-premises (2016/2019) use **PnPjs v2**. Always confirm the project's SPFx version before recommending v4-specific syntax.

## Setting up context in SPFx (do this once, in `onInit`)

```typescript
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs"; // + whatever sub-modules you need, see Selective Imports below

export default class MyWebPart extends BaseClientSideWebPart<IMyWebPartProps> {
  private sp: SPFI;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this.sp = spfi().using(SPFx(this.context));
  }
}
```
If also using Graph in the same component, **alias the SPFx import** for each package to avoid a name collision:
```typescript
import { spfi, SPFx as spSPFx } from "@pnp/sp";
import { graphfi, SPFx as graphSPFx } from "@pnp/graph";

const sp = spfi().using(spSPFx(this.context));
const graph = graphfi().using(graphSPFx(this.context));
```
`onInit()` — not the constructor — is the correct place; `this.context` isn't populated yet in the constructor.

### Connecting to a different web
```typescript
import { spfi } from "@pnp/sp";
import { AssignFrom } from "@pnp/core";

// new spfi instance pointed at another web, reusing the current auth/observers
const spWebB = spfi("https://tenant.sharepoint.com/sites/other").using(AssignFrom(sp.web));
```

## Selective imports — always required, easy to forget

Importing `@pnp/sp` alone gives you almost nothing; each sub-area must be imported for its methods to "attach" to the fluent chain:
```typescript
import "@pnp/sp/webs";
import "@pnp/sp/lists";      // or "@pnp/sp/lists/web" for a narrower/smaller import
import "@pnp/sp/items";      // or "@pnp/sp/items/list"
import "@pnp/sp/files";
import "@pnp/sp/folders";
```
**Gotcha**: a type-only import (`import { IList } from "@pnp/sp/lists"`) does NOT attach the runtime functionality — TypeScript treats it as type-only and drops it at build time, causing a runtime error (`sp.web.lists is not a function`). You need **both**:
```typescript
import { IList } from "@pnp/sp/lists"; // type
import "@pnp/sp/lists";                 // runtime attachment
```
For quick prototypes/tests where bundle size doesn't matter, `import "@pnp/sp/presets/all"` attaches everything (no tree-shaking) — **do not use this in production web parts**, only for scratch/testing.

## Core CRUD patterns

```typescript
// GET all + GET by id + $select/$top/$orderBy
const items = await sp.web.lists.getByTitle("My List").items();
const item = await sp.web.lists.getByTitle("My List").items.getById(1)();
const items2 = await sp.web.lists.getByTitle("My List").items
  .select("Title", "Description").top(5).orderBy("Modified", true)();

// ADD
const newItem = await sp.web.lists.getByTitle("My List").items.add({ Title: "Title", Description: "..." });
// v4: add/update return the raw server response directly (NOT wrapped in .data / .item like v3)
console.log(newItem.Id); // not newItem.data.Id

// UPDATE (many endpoints return 204/void — re-fetch if you need the updated object back)
await sp.web.lists.getByTitle("My List").items.getById(1).update({ Title: "New Title" });

// DELETE / RECYCLE
await sp.web.lists.getByTitle("My List").items.getById(1).delete();
await sp.web.lists.getByTitle("My List").items.getById(1).recycle(); // soft delete, goes to recycle bin

// Lookup fields — need expand + select
const withLookup = await sp.web.lists.getByTitle("LookupList").items
  .select("Title", "Lookup/Title", "Lookup/ID").expand("Lookup")();

// Ensure a list exists (create-if-missing) — NOT batchable
const ensured = await sp.web.lists.ensure("My List", "description", 100 /* generic list template */);
if (ensured.created) { /* it was just created */ }
```
Full method catalogs: `sp/items.md`, `sp/lists.md`, `sp/webs.md`, `sp/folders.md`, `sp/files.md`, `sp/fields.md`, `sp/content-types.md`, `sp/views.md`.

### User & lookup field values on add/update
- Single-value user field: `User1Id: 9`. Multi-value: `User2Id: [16, 45]`.
- Single-value lookup: `LookupFieldId: 2`. Multi-value lookup: `MultiLookupFieldId: [1, 56]`.
- Internal field name convention: append `Id` to the field's `EntityPropertyName`.

### Paging (lists >5000 items)
v4 removed the old `items/get-all` endpoint in favor of the **async iterator** pattern:
```typescript
for await (const pageOfItems of sp.web.lists.getByTitle("BigList").items.top(2000)) {
  console.log(pageOfItems); // one page (array) per iteration
  // `break;` here stops after the first page
}
```

### Filtering: fluent filter vs CAML
- **Fluent filter** (preview feature) gives strongly-typed filters without hand-building OData strings:
  ```typescript
  interface ListItem extends IListItem { FirstName: string; Age: number; }
  const r = await sp.web.lists.getByTitle("L").items
    .filter<ListItem>(f => f.text("FirstName").equals("John").and().number("Age").greaterThan(30))();
  ```
  Supports `and()`/`or()` nesting; operators vary by field type (text: `startsWith`/`contains`; numeric/date: comparisons; boolean: `isTrue`/`isFalse`).
- **Metadata/managed-metadata (taxonomy) fields cannot be filtered with `$filter`** — use `getItemsByCAMLQuery` instead:
  ```typescript
  await sp.web.lists.getByTitle("TaxonomyList").getItemsByCAMLQuery({
    ViewXml: `<View><Query><Where><Eq><FieldRef Name="MetaData"/><Value Type="TaxonomyFieldType">Term 2</Value></Eq></Where></Query></View>`,
  });
  ```

## Files & folders

```typescript
import "@pnp/sp/files";
import "@pnp/sp/folders";

// small file (<10MB) — addUsingPath
const result = await sp.web.getFolderByServerRelativePath("Shared Documents")
  .files.addUsingPath(encodeURI(file.name), file, { Overwrite: true });

// large file — addChunked, with progress callback
const result2 = await sp.web.getFolderByServerRelativePath("Shared Documents")
  .files.addChunked(encodeURI(file.name), file, {
    progress: (data) => console.log("progress", data),
    Overwrite: true,
  });

// overwrite existing file content
await sp.web.getFileByServerRelativePath("/sites/dev/documents/test.txt").setContent("new content");
```
Both `addUsingPath` and `addChunked` **do not support batching**. `EnsureUniqueFileName` and `Overwrite` are mutually exclusive — omit `Overwrite` when using `EnsureUniqueFileName`.

## Batching — batch multiple calls into one HTTP round-trip

```typescript
import "@pnp/sp/batching";

const [batchedSP, execute] = sp.batched();
const results: any[] = [];

// use .then(), NOT await, on each call inside a batch — awaiting would block before the batch executes
batchedSP.web.lists.getByTitle("MyList").items.add({ Title: "1" }).then(r => results.push(r));
batchedSP.web.lists.getByTitle("MyList").items.add({ Title: "2" }).then(r => results.push(r));

await execute(); // now results is populated
```
Rules:
- All calls in one batch must target **the same web**.
- **Never reuse the same queryable instance twice** in a batch (throws "This instance is already part of a batch"). Start a fresh fluent chain (`batchedSP.web.lists...`) for each call, or use the `Users(...)`/factory-style pattern to branch off an existing reference.
- A batch cannot be reused after `execute()` resolves — create a new one for further batched calls.
- `list.ensure()` and file `addUsingPath`/`addChunked` are **not batchable**.

## Error handling

All request failures surface as `HttpRequestError` (extends `Error`), with `isHttpRequestError: true`, `status`, `statusText`, and an **unread** `response` you can `.json()` yourself for the SharePoint OData error payload:
```typescript
import { HttpRequestError } from "@pnp/queryable";

try {
  await sp.web.lists.getByTitle("DoesNotExist")();
} catch (e) {
  if (e?.isHttpRequestError) {
    const data = await (e as HttpRequestError).response.json();
    const message = data["odata.error"]?.message?.value ?? e.message;
    if ((e as HttpRequestError).status === 404) { /* handle not-found */ }
  }
}
```
The library includes **built-in retry logic for 429/503/504** responses — don't hand-roll your own retry wrapper for throttling.

## Search

```typescript
import "@pnp/sp/search";
import { SearchQueryBuilder } from "@pnp/sp/search";

const results = await sp.search("test"); // simple text query
const results2 = await sp.search(SearchQueryBuilder("test").rowLimit(10).enableInterleaving);
console.log(results.PrimarySearchResults);
```
Caching search results: `.using(Caching())` from `@pnp/queryable`.

## Current user & site users

```typescript
import "@pnp/sp/site-users/web";

const me = await sp.web.currentUser();
const user = await sp.web.getUserById(6)();
const ensured = await sp.web.ensureUser("someone@contoso.com");
```

## Architecture notes (only go this deep if the user needs custom extension behavior)

- PnPjs is built on a `Queryable`/`Timeline` lifecycle: `construct → init → pre → auth → send → parse → post → data → dispose`, plus `log`/`error` which can fire at any point. Custom cross-cutting behavior (custom auth, custom error handling, telemetry) is added by registering **observers** on these moments via `.on.<moment>(...)`, and packaged for reuse as **behaviors** applied with `.using(...)`.
- Full details: `concepts/timeline`, `concepts/observers`, `concepts/behaviors`, `queryable/queryable.md` in the PnPjs docs tree.

## v3 → v4 migration highlights (if the user has existing v3 code)

1. **Add/update no longer return `{ data, <entity> }`** — they return the raw response directly.
   ```typescript
   // v3: const newItemId = (await items.add({...})).data.Id;
   // v4: const newItemId = (await items.add({...})).Id;
   ```
2. **SharePoint Taxonomy moved to `@pnp/graph`** — `@pnp/sp` no longer has taxonomy; re-implement against the Graph taxonomy endpoints.
3. **Paging**: `items/get-all` was removed — use the async iterator pattern shown above.

## Node.js / server-side usage (Azure Functions, scripts — not SPFx)

```console
npm install @pnp/sp @pnp/nodejs
```
```typescript
import { spfi } from "@pnp/sp";
import { SPDefault } from "@pnp/nodejs";

const sp = spfi().using(SPDefault({
  baseUrl: "https://contoso.sharepoint.com/sites/dev",
  msal: { config: msalConfig, scopes: ["https://contoso.sharepoint.com/.default"] },
}));
```
Only relevant if the user is building an Azure Function or script alongside their SPFx solution, not for in-browser SPFx code (which always uses `SPFx(this.context)`).

## Full sp/* module index (for looking up anything not covered above)

`webs`, `lists`, `items`, `files`, `folders`, `fields`, `views`, `content-types`, `attachments`, `search`, `sharing`, `site-users`, `site-groups`, `permissions`, `security`, `subscriptions`, `navigation`, `hubsites`, `recycle-bin`, `regional-settings`, `related-items`, `comments-likes`, `context-info`, `clientside-pages` (provisioning/authoring modern pages programmatically), `site-designs`, `site-scripts`, `user-custom-actions`, `profiles`, `social`, `alm` (app lifecycle / tenant app catalog operations), `column-defaults`, `features`, `forms`, `groupSiteManager`, `sp-utilities-utility`, `publishing-sitepageservice`. Fetch `https://github.com/pnp/pnpjs/blob/version-4/docs/sp/<name>.md` for any of these when a task needs it.
