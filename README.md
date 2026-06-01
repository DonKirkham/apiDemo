# apiDemo2

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
- `SHAREPOINT_APP_CLIENT_SECRET`
- `SHAREPOINT_RESOURCE` (example: `https://contoso.sharepoint.com/.default`)
- `ALLOWED_CALLER_DOMAINS` (comma-separated origins)
- `ENTRA_AUDIENCE` (your API App ID URI)
- `ALLOWED_CLIENT_APP_IDS` (optional comma-separated app IDs)

`local.settings.sample.json` is included as a template.

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

## Automated Azure deployment

Infrastructure and deployment automation files are included:

- `infra/main.bicep` – deploys storage account, consumption plan, Linux Function App, and app settings
- `.github/workflows/deploy.yml` – provisions/updates Azure resources and deploys the function app on push to `main` (or manually)

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
