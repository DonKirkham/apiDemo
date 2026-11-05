const required = ['AZURE_TENANT_ID', 'SHAREPOINT_RESOURCE'] as const;

export function getMissingRequired(): string[] {
  return required.filter((key) => !process.env[key]) as string[];
}

export function parseCsv(value: string | undefined): string[] {
  return (value || '')
    .split(',')
    .map((part) => part.trim())
    .filter(Boolean);
}

export const allowedCallerDomains = parseCsv(process.env.ALLOWED_CALLER_DOMAINS);
export const allowedClientAppIds = parseCsv(process.env.ALLOWED_CLIENT_APP_IDS);
export const tenantId = process.env.AZURE_TENANT_ID;
export const sharePointResource = process.env.SHAREPOINT_RESOURCE;
// When set, these pin every request to a single site/list and the
// client-supplied siteUrl/listTitle are ignored (locked-down mode). Leave
// them unset to allow callers to target any site/list per request.
export const sharePointSiteUrl = process.env.SHAREPOINT_SITE_URL;
export const sharePointListTitle = process.env.SHAREPOINT_LIST_TITLE;
export const entraAudience = process.env.ENTRA_AUDIENCE;
export const entraIssuer =
  process.env.ENTRA_ISSUER ||
  (process.env.AZURE_TENANT_ID
    ? `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/v2.0`
    : undefined);
export const entraJwksUri =
  process.env.ENTRA_JWKS_URI ||
  (process.env.AZURE_TENANT_ID
    ? `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/discovery/v2.0/keys`
    : undefined);
