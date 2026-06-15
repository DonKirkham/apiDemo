# apiDemo

Demo Azure Functions API used by SPFx for elevated SharePoint list operations.

## What this project provides

- **`POST /api/elevated/domain`**
  - Accepts calls from configured SPFx origins only (`ALLOWED_CALLER_DOMAINS`)
  - Then performs app-only SharePoint list item creation with elevated permissions
- **`POST /api/elevated/entra`**
  - Requires and validates an Entra ID bearer token
  - Optionally restricts callers to allowed client app IDs (`ALLOWED_CLIENT_APP_IDS`)
  - Then performs the same app-only elevated SharePoint list item creation
- **`GET /api/health`** for readiness checks

## Request payload (both POST endpoints)

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

## Required configuration

Set these app settings locally (`local.settings.json`) or in Azure:

- `AZURE_TENANT_ID`
- `SHAREPOINT_APP_CLIENT_ID`
- **App-only credential — one of:**
  - `SHAREPOINT_CERT_PEM` (**recommended**): full PEM (private key + certificate). SharePoint REST only honors **certificate-based** Entra app-only tokens; a client secret is rejected with `401 Unsupported app only token`.
  - `SHAREPOINT_APP_CLIENT_SECRET`: works for token acquisition but the SharePoint call will fail — kept only as a fallback / for demonstration.
- `SHAREPOINT_RESOURCE` (example: `https://contoso.sharepoint.com/.default`)
- `ALLOWED_CALLER_DOMAINS` (comma-separated origins)
- `ENTRA_AUDIENCE` (your API App ID URI)
- `ALLOWED_CLIENT_APP_IDS` (optional comma-separated app IDs)

`local.settings.sample.json` is included as a template.

### App registration permissions

The app registration (`SHAREPOINT_APP_CLIENT_ID`) needs a SharePoint Online **application**
permission with admin consent:

- **`Sites.ReadWrite.All`** — sufficient for this API, which **creates items in an existing list**.
- The target list must already exist. Creating lists/site structure would require the
  higher `Sites.Manage.All` (or `Sites.FullControl.All`), which this API intentionally does not need.

Generate and upload the certificate in one step:

```bash
az ad app credential reset --id <SHAREPOINT_APP_CLIENT_ID> --create-cert --append
# put the contents of the emitted .pem file into SHAREPOINT_CERT_PEM
```

## Run locally

1. Install dependencies:
   ```bash
   npm install
   ```
2. Copy config template:
   ```bash
   cp local.settings.sample.json local.settings.json
   ```
3. Start Functions host:
   ```bash
   npm start
   ```

## Deploy from your machine

```bash
# 1. Provision/update infrastructure
az deployment group create \
  --resource-group <rg> \
  --template-file infra/main.bicep \
  --parameters location=<region> functionAppName=<name> storageAccountName=<name> ...

# 2. Publish the code (remote build is REQUIRED)
func azure functionapp publish <functionAppName> --javascript --build remote
```

> **`--build remote` is required.** `.funcignore` excludes `node_modules`, so dependencies
> are installed server-side by Oryx. The Flex Consumption SKU does **not** support the
> `SCM_DO_BUILD_DURING_DEPLOYMENT` app setting, so the build must be requested at deploy
> time via `--build remote` (CLI) or `remote-build: true` (GitHub Actions). A plain publish
> will deploy without dependencies and register zero functions.

## Automated Azure deployment

Infrastructure and deployment automation files are included:

- `infra/main.bicep` – deploys storage account, Flex Consumption plan (Node 22), Linux Function App, App Insights, and app settings
- `.github/workflows/deploy.yml` – provisions/updates Azure resources and deploys the function app (remote build) on manual dispatch

### GitHub configuration for workflow

Set repository **Secrets**:

- `AZURE_CLIENT_ID`
- `AZURE_TENANT_ID`
- `AZURE_SUBSCRIPTION_ID`
- `SHAREPOINT_APP_CLIENT_ID`
- `SHAREPOINT_APP_CLIENT_SECRET`

Set repository **Variables**:

- `AZURE_RESOURCE_GROUP`
- `AZURE_LOCATION`
- `FUNCTION_APP_NAME`
- `STORAGE_ACCOUNT_NAME`
- `SHAREPOINT_RESOURCE`
- `ALLOWED_CALLER_DOMAINS`
- `ENTRA_AUDIENCE`
- `ALLOWED_CLIENT_APP_IDS`
