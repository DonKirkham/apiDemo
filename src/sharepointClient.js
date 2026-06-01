const {
  tenantId,
  sharePointClientId,
  sharePointClientSecret,
  sharePointResource
} = require('./config');

let cachedToken;

async function getAppOnlyToken() {
  if (cachedToken && cachedToken.expiresAt > Date.now() + 60 * 1000) {
    return cachedToken.accessToken;
  }

  const tokenResponse = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      body: new URLSearchParams({
        client_id: sharePointClientId,
        client_secret: sharePointClientSecret,
        grant_type: 'client_credentials',
        scope: sharePointResource
      })
    }
  );

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
