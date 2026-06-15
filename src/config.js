const required = [
  'AZURE_TENANT_ID',
  'SHAREPOINT_APP_CLIENT_ID',
  'SHAREPOINT_RESOURCE'
];

function getMissingRequired() {
  const missing = required.filter((key) => !process.env[key]);

  // App-only auth needs EITHER a certificate (preferred) OR a client secret.
  if (!process.env.SHAREPOINT_CERT_PEM && !process.env.SHAREPOINT_APP_CLIENT_SECRET) {
    missing.push('SHAREPOINT_CERT_PEM or SHAREPOINT_APP_CLIENT_SECRET');
  }

  return missing;
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
  // Full PEM (certificate + private key). Newlines may be escaped as \n in app settings.
  sharePointCertPem: process.env.SHAREPOINT_CERT_PEM
    ? process.env.SHAREPOINT_CERT_PEM.replace(/\\n/g, '\n')
    : undefined,
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
