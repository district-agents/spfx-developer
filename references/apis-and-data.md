# Connecting to APIs and Data from SPFx

Source: [Connect to SharePoint APIs](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/connect-to-sharepoint), [Connect to Microsoft Graph](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-msgraph), [Connect to Entra ID-secured APIs](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-aadhttpclient), [Connect to anonymous APIs](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/connect-to-anonymous-apis), [Working with the AADTokenProvider](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-aadtokenprovider), [Subscribe to list notifications](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/subscribe-to-list-notifications), [Webhooks overview](https://learn.microsoft.com/en-us/sharepoint/dev/apis/webhooks/overview-sharepoint-webhooks).

**Default to PnPjs (`@pnp/sp`/`@pnp/graph`) for SharePoint/Graph data access — see `references/pnpjs.md` for the full, preferred API.** The sections below cover the *native* SPFx HTTP clients that PnPjs itself is built on top of; reach for them directly only for a single trivial one-off call, or when the user explicitly wants to avoid the dependency.

## 1. SharePoint REST — `SPHttpClient` (native, no dependency)

Built in, available on `this.context.spHttpClient` in any web part/extension.

```typescript
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

this.context.spHttpClient
  .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`,
       SPHttpClient.configurations.v1)
  .then((response: SPHttpClientResponse) => response.json());
```
- Handles SharePoint auth cookies automatically; defaults to OData v4.
- Fine for a single simple GET; gets cumbersome fast for batching, complex filters, or strongly-typed responses — use **PnPjs** (`references/pnpjs.md`) for anything beyond that, which is what this skill defaults to.
- POST example (create list item) needs `X-RequestDigest` handled automatically by `spHttpClient.post`, but if hand-rolling `fetch`, see [Work with __REQUESTDIGEST](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/basics/working-with-requestdigest.md).

## 2. Microsoft Graph — `MSGraphClient`

```typescript
this.context.msGraphClientFactory
  .getClient('3') // MSGraphClientV3
  .then((client: MSGraphClientV3) => {
    client.api('/me').get((error, user) => { /* ... */ });
  });
```
- Introduced in SPFx v1.4.1; **do not** use MSAL.js directly with SPFx — unsupported since v1.4.1. `MSGraphClient`/`AadHttpClient` are the sanctioned wrappers around the implicit OAuth flow, and PnPjs's `graphfi().using(SPFx(this.context))` wraps this same mechanism — see `references/pnpjs.md`.
- Common permission scopes need to be requested and approved (see section 4).
- The [Microsoft Graph Toolkit](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-microsoft-graph-toolkit.md) provides pre-built Graph-connected web components (people picker, agenda, person card, etc.); also check `references/react-controls.md` — several `@pnp/spfx-controls-react` controls (PeoplePicker, TeamPicker, TeamChannelPicker) cover similar ground and may fit better in a component-based UI.

## 3. Any other Entra ID (Azure AD)-secured API — `AadHttpClient`

Introduced in SPFx v1.4.1. Handles the OAuth flow via the **SharePoint Online Client Extensibility** service principal (provisioned automatically in every M365 tenant's Entra ID).

```typescript
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

this.context.aadHttpClientFactory
  .getClient('https://contoso.azurewebsites.net')
  .then((client: AadHttpClient) => {
    return client.get('https://contoso.azurewebsites.net/api/orders', AadHttpClient.configurations.v1);
  })
  .then((response: HttpClientResponse) => response.json());
```
One client instance is bound to one resource — create a new instance per resource.

### Requesting permissions (required before AadHttpClient/MSGraphClient calls will succeed)
Add to `config/package-solution.json`:
```json
{
  "solution": {
    "webApiPermissionRequests": [
      { "resource": "Microsoft Graph", "scope": "Calendars.Read" },
      { "resource": "Microsoft Graph", "scope": "User.ReadBasic.All" }
    ]
  }
}
```
- `resource` must be the app's **displayName**, not its objectId (using objectId errors on approval).
- For an Entra ID app registered in a **different** tenant (multi-tenant app), also include `appId` and `replyUrl` (SPFx v1.15.2+) so SharePoint can register a service principal for it locally.
- Deploying the `.sppkg` to the app catalog creates a **pending permission request** — a Global or SharePoint admin must approve it (new SharePoint admin center's API access page, or PowerShell `Approve-SPOTenantServicePrincipalPermissionRequest`, or CLI for M365 `spo serviceprincipal permissionrequest approve`).
- **Deployment always succeeds regardless of approval** — never assume a requested permission has been granted; code defensively.
- Permissions are **tenant-wide**, not per-app; removing the solution does **not** revoke previously granted permissions (admin must revoke manually via `Revoke-SPOTenantServicePrincipalPermission`).
- Admins can disable the whole web-API-request mechanism tenant-wide via `Disable-SPOTenantServicePrincipal` (re-enable with `Enable-SPOTenantServicePrincipal`).

### AadTokenProvider (lower-level)
For scenarios needing a raw access token rather than a wrapped HTTP client, see [Working with the AADTokenProvider](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-aadtokenprovider.md).

## 4. Anonymous / public APIs — `HttpClient`

```typescript
this.context.httpClient.get('https://api.example.com/public', HttpClient.configurations.v1);
```
No auth handling — for public endpoints only. See [Connect to anonymous APIs](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/connect-to-anonymous-apis.md).

## 5. Real-time / change notifications

- **[List webhooks](https://learn.microsoft.com/en-us/sharepoint/dev/apis/webhooks/lists/overview-sharepoint-list-webhooks.md)** — server push notifications on list changes to an external endpoint (e.g., Azure Function); see get-started guide and create/update/delete/get-subscription REST references.
- **[Subscribe to list notifications](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/subscribe-to-list-notifications.md)** from within an SPFx component for lighter-weight change detection.
- Webhook reference implementation and an Azure Functions-based approach (including an `azd` template) are documented under `apis/webhooks/`.

## Tutorials worth pointing to

- [Consume Microsoft Graph](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-aad-tutorial.md) — full walkthrough building a people-search web part against Graph.
- [Consume enterprise APIs](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-aadhttpclient-enterpriseapi.md) — build+secure an Azure Functions API with Azure App Service auth and call it from SPFx.
- [Consume multi-tenant enterprise APIs](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-aadhttpclient-enterpriseapi-multitenant.md).
