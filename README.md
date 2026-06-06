# Copilot API Explorer — SPFx Web Part

An interactive **demo web part** that lets you explore and call five
Microsoft 365 Copilot / Microsoft Graph APIs live from a SharePoint page:

| # | API | Method | Endpoint |
|---|-----|--------|----------|
| 1 | Retrieval API | `POST` | `/copilot/retrieval` |
| 2 | Interaction Export | `GET` | `/copilot/users/{userId}/interactionHistory/getAllEnterpriseInteractions` |
| 3 | Usage – User Count Summary | `GET` | `/reports/getMicrosoft365CopilotUserCountSummary(period='D7')` |
| 4 | Usage – User Count Trend | `GET` | `/reports/getMicrosoft365CopilotUserCountTrend(period='D30')` |
| 5 | Usage – Usage User Detail | `GET` | `/reports/getMicrosoft365CopilotUsageUserDetail(period='D7')` |

Each API gets its own tab (Fluent UI `Pivot`) with four sections: **API Info**,
**Request Builder**, **Response Viewer** and a shared **Permission Matrix**.

![version](https://img.shields.io/badge/SPFx-1.23.1-green.svg)
![react](https://img.shields.io/badge/React-17-blue.svg)
![node](https://img.shields.io/badge/Node-22.14%2B-green.svg)

> ⚠️ Several of these APIs are in **beta** (`/beta`). The endpoint URL is editable
> in every tab so you can switch to `/v1.0` where a GA surface exists.

---

## Authentication

The explorer supports two auth modes (switchable at the top):

1. **Delegated — `MSGraphClientV3`** (`context.msGraphClientFactory.getClient('3')`).
   Calls run as the **signed-in SharePoint user**. The raw token can be acquired
   on demand via the AAD token provider purely to **decode and display its
   claims** (upn, scopes, expiry countdown).
2. **Manual bearer token** (App Registration / client-credentials flow). You
   paste an access token and the web part uses it in a raw `fetch()` with an
   `Authorization: Bearer` header.
   **DEMO ONLY — never hardcode or paste tokens in production.**

In both modes the JWT is decoded **client-side with plain base64url** (no
external library, signature **not** verified) to surface `upn`, `scp`/`roles`,
audience, tenant and a live expiry countdown. A green/red **permission check**
shows whether the current token carries the scopes the selected API needs.

---

## Required API permissions

Register the following on your Azure AD **App Registration** (for the manual
app-token flow) and/or approve them in the **SharePoint API access** page (for
the delegated `MSGraphClientV3` flow).

| API | Delegated scope | Application scope | Admin consent |
|-----|-----------------|-------------------|---------------|
| Retrieval | `Files.Read.All`, `Sites.Read.All` (or `ExternalItem.Read.All` for connectors) | *not supported* | ✅ |
| Interaction Export | *not supported* | `AiEnterpriseInteraction.Read.All` | ✅ |
| Usage – Summary | `Reports.Read.All` | `Reports.Read.All` | ✅ |
| Usage – Trend | `Reports.Read.All` | `Reports.Read.All` | ✅ |
| Usage – User Detail | `Reports.Read.All` | `Reports.Read.All` | ✅ |

Notes worth knowing on stage:

- **Retrieval is delegated-only** — there is no app-only flow. It is permission
  trimmed to whatever the signed-in user can see.
- **Interaction Export is application-only** — `MSGraphClientV3` (delegated)
  **cannot** call it. Use the Manual token mode with an app token carrying
  `AiEnterpriseInteraction.Read.All`. It also requires a valid Microsoft 365
  Copilot license on the target user.
- **Usage reports** are tenant-wide analytics: `Reports.Read.All` + admin
  consent. There is no per-user delegated story because the data spans the org.

### Granting the delegated scopes for SPFx

This solution declares the delegated Graph scopes in
[`config/package-solution.json`](config/package-solution.json) via
`webApiPermissionRequests`:

```json
"webApiPermissionRequests": [
  { "resource": "Microsoft Graph", "scope": "Files.Read.All" },
  { "resource": "Microsoft Graph", "scope": "Sites.Read.All" },
  { "resource": "Microsoft Graph", "scope": "Reports.Read.All" }
]
```

After deploying the package, a tenant admin must **approve** these in
**SharePoint Admin Center → Advanced → API access**.

---

## Run in the local workbench

Because the explorer calls real Microsoft Graph endpoints, use the **hosted**
workbench (the local `localhost` workbench has no Graph token):

```powershell
npm install            # restores existing dependencies (no new packages added)
heft start             # builds, serves the bundle on https://localhost:4321 and opens the hosted workbench
```

Useful Heft commands:

```powershell
heft start                      # dev loop: clean, build, serve, watch (equivalent to legacy `gulp serve`)
heft test                       # one-shot dev build (equivalent to legacy `gulp bundle`)
heft test --production          # production bundle
heft package-solution           # build the .sppkg (debug)
heft package-solution --production  # build the .sppkg (ship)
heft clean                      # wipe lib / temp / dist / sharepoint
```

Then open:

```
https://<your-tenant>.sharepoint.com/_layouts/15/workbench.aspx
```

Add the **Copilot API Explorer** web part from the *Advanced* group.

> Delegated calls require the `webApiPermissionRequests` scopes to be approved
> first, otherwise Graph returns `403`. The Manual token mode works without the
> SPFx grant because it uses your pasted token directly.

---

## Deploy to the SharePoint App Catalog

```powershell
heft test --clean --production
heft package-solution --production
```

1. Upload the generated package from `sharepoint/solution/hands-on-copilot-apis-demo.sppkg`
   to your **tenant App Catalog**.
2. Choose **Deploy** when prompted.
3. In **SharePoint Admin Center → Advanced → API access**, approve the pending
   `Files.Read.All`, `Sites.Read.All` and `Reports.Read.All` requests.
4. Add the web part to a modern page.

---

## Demo script (the 5 steps)

Toggle **Demo Mode** (top-right) to get a floating, step-by-step guide. Each step
highlights the relevant UI section with a coloured border and switches to the
right tab.

1. **Auth** — acquire a token, decode it, and contrast **Delegated** scopes
   (`scp`) vs **Application** roles (`roles`).
2. **Retrieval POST** — build the body (`queryString` + `dataSource`), fire it,
   and explain the `relevanceScore` (cosine similarity, normalised 0–1).
3. **Retrieval + KQL** — add a `filterExpression` (e.g. `FileType:"pptx"` or a
   SharePoint `path:`), re-run, and show the narrower, scoped results.
4. **Swap token (different user)** — paste a token for another user, re-run
   Retrieval, and show permission trimming with the callout *"This user sees
   different results"*.
5. **Usage APIs** — explain why these need **Application** permissions + admin
   consent (tenant-wide data, no delegated per-user story).

---

## Code structure

```
src/webparts/demoLauncher/
  DemoLauncherWebPart.ts          # passes WebPartContext into React
  components/
    DemoLauncher.tsx              # root orchestrator (all interactive state)
    IDemoLauncherProps.ts         # props (incl. WebPartContext)
    types.ts                      # all request/response & auth interfaces
    copilotApi.ts                 # API catalog + JWT decode + exec + cURL/Postman
    AuthPanel.tsx                 # auth modes, JWT decode, permission check
    ApiInfoPanel.tsx              # method badge, editable URL, perms, docs link
    RequestBuilder.tsx            # JSON body / query params + Copy cURL/Postman
    ResponseViewer.tsx            # execute + status/time/size + rich renderers
    PermissionMatrix.tsx          # static 5-API permission table
    DemoModeOverlay.tsx           # floating step-by-step guide
```

### Notable implementation choices

- **Data-driven tabs.** Every API is described once in `API_CATALOG`
  (`copilotApi.ts`); the UI components are purely presentational. Adding a 6th
  API is a *data* change.
- **No external libraries added.** JSON syntax highlighting is done by tokenising
  the string into coloured React `<span>`s (no `dangerouslySetInnerHTML`, so no
  XSS surface). The bar charts are pure CSS (`width: %`). JWT decoding is plain
  base64url.
- **Correct Retrieval body.** The real API uses
  `queryString` / `dataSource` / `filterExpression` / `resourceMetadata` /
  `maximumNumberOfResults` (the KQL field maps to `filterExpression`). The older
  `entityTypes`/`from`/`size` shape would return `400`, so it is **not** used.

---

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR
IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE,
MERCHANTABILITY, OR NON-INFRINGEMENT.** The manual bearer-token feature exists
for live demos only — never embed real access tokens in production code.

## References

- [Retrieval API overview](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/retrieval/overview)
- [Retrieval API reference](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/retrieval/copilotroot-retrieval)
- [Interaction Export — getAllEnterpriseInteractions](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions)
- [Microsoft 365 Copilot usage report APIs](https://learn.microsoft.com/en-us/graph/api/resources/microsoft365-copilot-usage-report-api)
- [Use Microsoft Graph in SPFx](https://learn.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)