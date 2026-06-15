const crypto = require('crypto');
const jwt = require('jsonwebtoken');
const {
  tenantId,
  sharePointClientId,
  sharePointClientSecret,
  sharePointCertPem,
  sharePointResource
} = require('./config');

let cachedToken;
let cachedCredential;

// Parse the PEM once into an X509 certificate (for the x5t thumbprint) and a
// private key (for signing). SharePoint REST only honors certificate-based
// Entra app-only tokens, so this is the path that actually works.
function loadCertificateCredential() {
  if (cachedCredential) {
    return cachedCredential;
  }

  const certMatch = sharePointCertPem.match(
    /-----BEGIN CERTIFICATE-----[\s\S]+?-----END CERTIFICATE-----/
  );
  const keyMatch = sharePointCertPem.match(
    /-----BEGIN (?:RSA |ENCRYPTED )?PRIVATE KEY-----[\s\S]+?-----END (?:RSA |ENCRYPTED )?PRIVATE KEY-----/
  );

  if (!certMatch || !keyMatch) {
    throw new Error('SHAREPOINT_CERT_PEM must contain both a CERTIFICATE and a PRIVATE KEY block.');
  }

  const certificate = new crypto.X509Certificate(certMatch[0]);
  // x5t = base64url(SHA-1 thumbprint of the DER-encoded certificate)
  const x5t = crypto.createHash('sha1').update(certificate.raw).digest('base64url');

  cachedCredential = {
    privateKey: crypto.createPrivateKey(keyMatch[0]),
    x5t
  };
  return cachedCredential;
}

function buildClientAssertion(tokenUrl) {
  const { privateKey, x5t } = loadCertificateCredential();
  const now = Math.floor(Date.now() / 1000);

  return jwt.sign(
    {
      aud: tokenUrl,
      iss: sharePointClientId,
      sub: sharePointClientId,
      jti: crypto.randomUUID(),
      nbf: now,
      iat: now,
      exp: now + 10 * 60
    },
    privateKey,
    { algorithm: 'RS256', header: { alg: 'RS256', typ: 'JWT', x5t } }
  );
}

function buildTokenRequestBody(tokenUrl) {
  const body = new URLSearchParams({
    client_id: sharePointClientId,
    grant_type: 'client_credentials',
    scope: sharePointResource
  });

  if (sharePointCertPem) {
    body.set(
      'client_assertion_type',
      'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
    );
    body.set('client_assertion', buildClientAssertion(tokenUrl));
  } else {
    body.set('client_secret', sharePointClientSecret);
  }

  return body;
}

async function getAppOnlyToken() {
  if (cachedToken && cachedToken.expiresAt > Date.now() + 60 * 1000) {
    return cachedToken.accessToken;
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const tokenResponse = await fetch(tokenUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    body: buildTokenRequestBody(tokenUrl)
  });

  if (!tokenResponse.ok) {
    const body = await tokenResponse.text();
    throw new Error(`Failed to get SharePoint token (${tokenResponse.status}): ${body}`);
  }

  const tokenPayload = await tokenResponse.json();
  cachedToken = {
    accessToken: tokenPayload.access_token,
    expiresAt: Date.now() + (tokenPayload.expires_in || 3600) * 1000
  };

  return cachedToken.accessToken;
}

async function createSharePointListItem({ siteUrl, listTitle, fields }) {
  const token = await getAppOnlyToken();
  const authHeaderValue = 'Bearer '.concat(token);
  const sanitizedListTitle = (listTitle || '').replace(/'/g, "''");
  const endpoint = `${siteUrl.replace(/\/$/, '')}/_api/web/lists/GetByTitle('${encodeURIComponent(
    sanitizedListTitle
  )}')/items`;

  const response = await fetch(endpoint, {
    method: 'POST',
    headers: {
      Authorization: authHeaderValue,
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata'
    },
    body: JSON.stringify(fields)
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`SharePoint request failed (${response.status}): ${body}`);
  }

  return response.json();
}

module.exports = {
  createSharePointListItem
};
