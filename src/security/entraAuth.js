const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');
const {
  allowedClientAppIds,
  entraAudience,
  entraIssuer,
  entraJwksUri
} = require('../config');

let client;

function getClient() {
  if (!client) {
    if (!entraJwksUri) {
      throw Object.assign(new Error('ENTRA_JWKS_URI is required for Entra validation.'), { status: 500 });
    }

    client = jwksClient({
      jwksUri: entraJwksUri,
      cache: true,
      cacheMaxEntries: 5,
      cacheMaxAge: 10 * 60 * 1000
    });
  }

  return client;
}

function getSigningKey(header, callback) {
  getClient()
    .getSigningKey(header.kid)
    .then((key) => callback(null, key.getPublicKey()))
    .catch((err) => callback(err));
}

function getBearerToken(request) {
  const authHeader = request.headers.get('authorization') || '';
  const match = authHeader.match(/^Bearer\s+(.+)$/i);
  return match ? match[1] : null;
}

function ensureAllowedClientApplication(payload) {
  if (!allowedClientAppIds.length) {
    return;
  }

  const appId = payload.azp || payload.appid;
  if (!appId || !allowedClientAppIds.includes(appId)) {
    throw Object.assign(new Error('Calling application is not allowed.'), { status: 403 });
  }
}

async function validateEntraToken(request) {
  const token = getBearerToken(request);
  if (!token) {
    throw Object.assign(new Error('Authorization bearer token is required.'), { status: 401 });
  }

  if (!entraAudience || !entraIssuer) {
    throw Object.assign(new Error('ENTRA_AUDIENCE and tenant settings are required.'), { status: 500 });
  }

  const payload = await new Promise((resolve, reject) => {
    jwt.verify(
      token,
      getSigningKey,
      {
        algorithms: ['RS256'],
        audience: entraAudience,
        issuer: entraIssuer
      },
      (error, decoded) => {
        if (error) {
          reject(Object.assign(new Error(`Invalid token: ${error.message}`), { status: 401 }));
          return;
        }

        resolve(decoded);
      }
    );
  });

  ensureAllowedClientApplication(payload);
  return payload;
}

module.exports = {
  validateEntraToken
};
