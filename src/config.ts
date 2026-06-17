const required = ['AZURE_TENANT_ID', 'SHAREPOINT_APP_CLIENT_ID', 'SHAREPOINT_RESOURCE'] as const;

export function getMissingRequired(): string[] {
  const missing = required.filter((key) => !process.env[key]) as string[];

  // App-only auth needs EITHER a certificate (preferred) OR a client secret.
  if (!process.env.SHAREPOINT_CERT_PEM && !process.env.SHAREPOINT_APP_CLIENT_SECRET) {
    missing.push('SHAREPOINT_CERT_PEM or SHAREPOINT_APP_CLIENT_SECRET');
  }

  return missing;
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
export const sharePointClientId = process.env.SHAREPOINT_APP_CLIENT_ID;
export const sharePointClientSecret = process.env.SHAREPOINT_APP_CLIENT_SECRET;
// Full PEM (certificate + private key). Newlines may be escaped as \n in app settings.
export const sharePointCertPem = process.env.SHAREPOINT_CERT_PEM
  ? process.env.SHAREPOINT_CERT_PEM.replace(/\\n/g, '\n')
  : undefined;
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
