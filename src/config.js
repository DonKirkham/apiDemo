const required = [
  'AZURE_TENANT_ID',
  'SHAREPOINT_APP_CLIENT_ID',
  'SHAREPOINT_APP_CLIENT_SECRET',
  'SHAREPOINT_RESOURCE'
];

function getMissingRequired() {
  return required.filter((key) => !process.env[key]);
}

function parseCsv(value) {
  return (value || '')
    .split(',')
    .map((part) => part.trim())
    .filter(Boolean);
}

module.exports = {
  parseCsv,
  getMissingRequired,
  allowedCallerDomains: parseCsv(process.env.ALLOWED_CALLER_DOMAINS),
  allowedClientAppIds: parseCsv(process.env.ALLOWED_CLIENT_APP_IDS),
  tenantId: process.env.AZURE_TENANT_ID,
  sharePointClientId: process.env.SHAREPOINT_APP_CLIENT_ID,
  sharePointClientSecret: process.env.SHAREPOINT_APP_CLIENT_SECRET,
  sharePointResource: process.env.SHAREPOINT_RESOURCE,
  entraAudience: process.env.ENTRA_AUDIENCE,
  entraIssuer:
    process.env.ENTRA_ISSUER ||
    (process.env.AZURE_TENANT_ID
      ? `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/v2.0`
      : undefined),
  entraJwksUri:
    process.env.ENTRA_JWKS_URI ||
    (process.env.AZURE_TENANT_ID
      ? `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/discovery/v2.0/keys`
      : undefined)
};
