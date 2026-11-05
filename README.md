# api-demo

Demo Azure Functions API used by SPFx for elevated SharePoint list operations.

Written in **TypeScript** (Node v4 programming model, ESM) and talks to SharePoint
through **PnPjs** (`@pnp/sp`) using **app-only** auth via the Function App's
**managed identity** (`@azure/identity`'s `ManagedIdentityCredential`). No secret or
certificate is stored. Build output lands in `dist/`.

## What this project provides

Two auth-protected route families, each exposing **full CRUD** over a SharePoint
list via app-only (elevated) permissions:

- **`/api/domain/{itemId?}`**
  - Accepts calls only from origins in `ALLOWED_CALLER_DOMAINS`.
  - **Not real security.** Origin and Referer are headers any client can set, so a
    script outside the browser can fake them. It only blocks casual cross-site calls
    from browsers. Use the Entra route for real protection.
- **`/api/entra/{itemId?}`**
  - Requires and validates an Entra ID bearer token
  - Optionally restricts callers to allowed client app IDs (`ALLOWED_CLIENT_APP_IDS`)
  - If that list is empty, any app in your tenant with a token for this API is
    accepted. Scopes and app roles are not checked.
- **`GET /api/health`** for readiness checks

### CRUD operations (both route families)

| Operation | Method & route | Where params go |
| --------- | -------------- | --------------- |
| **Create** | `POST /api/{mode}` | `siteUrl`, `listTitle`, `fields` in JSON body |
| **Read (list)** | `GET /api/{mode}` | `siteUrl`, `listTitle` (+ optional OData `$filter`, `$select`, `$top`, `$orderby`) in query string |
| **Read (one)** | `GET /api/{mode}/{itemId}` | `siteUrl`, `listTitle` in query string |
| **Update** | `PATCH /api/{mode}/{itemId}` | `siteUrl`, `listTitle`, `fields` in JSON body |
| **Delete** | `DELETE /api/{mode}/{itemId}` | `siteUrl`, `listTitle` in query string (DELETE has no body) |

`{mode}` is `domain` or `entra`. There are no action words in the URL — the
**HTTP method selects the operation**. Update and delete perform an
unconditional `IF-MATCH: *` write (last-write-wins).

### Target list: pinned vs. caller-supplied

- **Locked-down (recommended):** set `SHAREPOINT_SITE_URL` and `SHAREPOINT_LIST_TITLE`.
  Every request then targets that one site/list and any `siteUrl`/`listTitle` sent
  by the caller is **ignored**.
- **Any site/any list:** leave both unset. Callers then supply `siteUrl` and
  `listTitle` per request (in the JSON body or query string). Only operate this
  way once the endpoint is fully locked down, since it lets an authorized caller
  reach any list the app registration can access. The `siteUrl` must be on a
  `*.sharepoint.com` host. Any other host is rejected, so the managed-identity token
  is never sent elsewhere.
- `itemId` in the route must be a whole number. Anything else returns 400.
- A request body that is not valid JSON returns 400.

## Request payload (create / update)

```json
{
  "siteUrl": "https://contoso.sharepoint.com/sites/Finance",
  "listTitle": "Requests",
  "fields": {
    "Title": "Created by API",
    "RequestType": "Demo"
  }
}
```

Example read with OData filtering:

```
GET /api/entra?siteUrl=https://contoso.sharepoint.com/sites/Finance&listTitle=Requests&$top=50&$select=Id,Title
```

## Required configuration

Set these app settings locally (`local.settings.json`) or in Azure:

- `AZURE_TENANT_ID`
- `SHAREPOINT_RESOURCE` (example: `https://contoso.sharepoint.com/.default`)
- `SHAREPOINT_SITE_URL` (optional — pins all requests to one site, e.g. `https://contoso.sharepoint.com/sites/Finance`)
- `SHAREPOINT_LIST_TITLE` (optional — pins all requests to one list, e.g. `Requests`)
- `ALLOWED_CALLER_DOMAINS` (comma-separated origins)
- `ENTRA_AUDIENCE` (your API App ID URI)
- `ALLOWED_CLIENT_APP_IDS` (optional comma-separated app IDs)
- `ENTRA_ISSUER` and `ENTRA_JWKS_URI` (optional overrides). By default both are
  derived from `AZURE_TENANT_ID`, so you normally don't set them.

`local.settings.sample.json` is included as a template.

> **Local runs:** a managed identity only exists inside Azure, so SharePoint calls
> will fail when running `npm start` on a dev machine. Test SharePoint behavior
> against the deployed Function App. Local storage for the Functions host uses
> Azurite (started by the VS Code `tasks.json`, a dev dependency).

### App registration permissions

The Function App's **system-assigned managed identity** needs the SharePoint Online
**application** permission `Sites.ReadWrite.All`, granted with admin consent:

- **`Sites.ReadWrite.All`** — sufficient for this API, which **reads, creates, updates, and deletes items in an existing list**.
- The target list must already exist. Creating lists/site structure would require the
  higher `Sites.Manage.All` (or `Sites.FullControl.All`), which this API intentionally does not need.

Grant it once after the infrastructure deployment (replace `<principalId>`; get it with
`az functionapp identity show -g <rg> -n <functionAppName> --query principalId -o tsv`):

```bash
SPO_APP=00000003-0000-0ff1-ce00-000000000000   # Office 365 SharePoint Online
SPO_SP=$(az ad sp show --id $SPO_APP --query id -o tsv)
ROLE_ID=$(az ad sp show --id $SPO_APP --query "appRoles[?value=='Sites.ReadWrite.All'].id | [0]" -o tsv)

az rest --method POST \
  --headers "Content-Type=application/json" \
  --uri "https://graph.microsoft.com/v1.0/servicePrincipals/<principalId>/appRoleAssignments" \
  --body "{\"principalId\":\"<principalId>\",\"resourceId\":\"$SPO_SP\",\"appRoleId\":\"$ROLE_ID\"}"
```

No certificate or client secret is created or rotated.

## Run locally

1. Install dependencies:
   ```bash
   npm install
   ```
2. Copy config template:
   ```bash
   cp local.settings.sample.json local.settings.json
   ```
3. Start Functions host (`npm start` runs `tsc` first via the `prestart` script,
   then `func start`):
   ```bash
   npm start
   ```

Other scripts:

```bash
npm run build   # compile TypeScript (src/*.ts -> dist/)
npm run watch   # incremental tsc -w during development
npm test        # run the unit tests (node --test via tsx, no build needed)
```

## Deploy from your machine

```bash
# 1. Provision/update infrastructure
az deployment group create \
  --resource-group <rg> \
  --template-file infra/main.bicep \
  --parameters location=<region> functionAppName=<name> storageAccountName=<name> ...

# 2. Publish the code (remote build is REQUIRED)
func azure functionapp publish <functionAppName> --build remote
```

> **`--build remote` is required.** `.funcignore` excludes `node_modules` and `dist`, so
> server-side Oryx both installs dependencies and runs `npm run build` to transpile the
> TypeScript into `dist/` (`main` in `package.json` points at `dist/functions/*.js`). The
> Flex Consumption SKU does **not** support the `SCM_DO_BUILD_DURING_DEPLOYMENT` app
> setting, so the build must be requested at deploy time via `--build remote` (CLI) or
> `remote-build: true` (GitHub Actions). A plain publish will deploy uncompiled sources
> with no dependencies and register zero functions.

## Automated Azure deployment

Infrastructure and deployment automation files are included:

Two workflows, each with one job:

- `.github/workflows/infra.yml` – **manual only** (Actions tab, *Run workflow*).
  Deploys `infra/main.bicep`: storage, Flex Consumption plan (Node 22), Function App,
  App Insights and app settings. Run it when the infrastructure changes.
- `.github/workflows/deploy.yml` – **runs on push to `main`** when app code changes
  (`src/`, `tests/`, `package*.json`, and so on), and can also be run manually. It
  installs, builds, runs the tests, and only then deploys the code with a remote build.
  It deploys code only and never touches infrastructure.

`infra/main.parameters.json` is a template for `infra/main.bicep` parameters. Fill it
in and pass it with `--parameters @infra/main.parameters.json` if you deploy from the CLI.

### GitHub configuration

Set repository **Secrets** (used by both workflows):

- `AZURE_CLIENT_ID`
- `AZURE_TENANT_ID`
- `AZURE_SUBSCRIPTION_ID`

Set repository **Variables** for `infra.yml`:

- `AZURE_RESOURCE_GROUP`
- `AZURE_LOCATION`
- `FUNCTION_APP_NAME`
- `STORAGE_ACCOUNT_NAME`
- `SHAREPOINT_RESOURCE`
- `ALLOWED_CALLER_DOMAINS`
- `ENTRA_AUDIENCE`
- `ALLOWED_CLIENT_APP_IDS`

`deploy.yml` also needs `FUNCTION_APP_NAME`, which is listed above.

### Why these settings

- **`excludedTypes: "Request"` in `host.json`:** stops Application Insights from
  sampling away request telemetry. Each HTTP call stays visible in the logs.
- **`typescript` in `dependencies` (not `devDependencies`):** the remote build runs
  `npm run build` (`tsc`) on the server, and this keeps `typescript` installed there.
- **`azurite` and `tsx` are dev dependencies:** only needed for local runs and tests.
