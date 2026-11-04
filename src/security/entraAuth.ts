import { HttpRequest } from '@azure/functions';
import jwt, { JwtHeader, SigningKeyCallback, JwtPayload } from 'jsonwebtoken';
import jwksClient from 'jwks-rsa';
import { allowedClientAppIds, entraAudience, entraIssuer, entraJwksUri } from '../config.js';
import { HttpError } from '../httpError.js';

// Entra signs tokens with keys that Microsoft rotates. The JWKS endpoint
// publishes the current public keys, and this client fetches and caches them.
// Caching avoids a network call on every request.
let client: ReturnType<typeof jwksClient> | undefined;

function getClient(): ReturnType<typeof jwksClient> {
  if (!client) {
    if (!entraJwksUri) {
      throw new HttpError(500, 'AZURE_TENANT_ID (or ENTRA_JWKS_URI) is required for Entra validation.');
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

// jsonwebtoken calls this to find the public key for a token. The token's
// header names the key it was signed with (`kid`), and we look that key up.
function getSigningKey(header: JwtHeader, callback: SigningKeyCallback): void {
  if (!header.kid) {
    callback(new Error('Token is missing a key id (kid).'));
    return;
  }

  getClient()
    .getSigningKey(header.kid)
    .then((key) => callback(null, key.getPublicKey()))
    .catch((err) => callback(err as Error));
}

function getBearerToken(request: HttpRequest): string | null {
  const authHeader = request.headers.get('authorization') || '';
  const match = authHeader.match(/^Bearer\s+(.+)$/i);
  return match ? match[1] : null;
}

// A valid token from the tenant is not enough by itself. If
// ALLOWED_CLIENT_APP_IDS is set, the calling app must also be on that list.
// When the list is empty, any app in the tenant with a token for this API is
// accepted. This demo does not check scopes or app roles.
function ensureAllowedClientApplication(payload: JwtPayload): void {
  if (!allowedClientAppIds.length) {
    return;
  }

  const appId = (payload.azp as string | undefined) || (payload.appid as string | undefined);
  if (!appId || !allowedClientAppIds.includes(appId)) {
    throw new HttpError(403, 'Calling application is not allowed.');
  }
}

export async function validateEntraToken(request: HttpRequest): Promise<JwtPayload> {
  const token = getBearerToken(request);
  if (!token) {
    throw new HttpError(401, 'Authorization bearer token is required.');
  }

  if (!entraAudience || !entraIssuer) {
    throw new HttpError(500, 'ENTRA_AUDIENCE and tenant settings are required.');
  }

  const payload = await new Promise<JwtPayload>((resolve, reject) => {
    // jwt.verify checks the signature, expiry, audience and issuer. The
    // algorithm is pinned to RS256. If we accepted whatever algorithm the token
    // names, an attacker could forge tokens (the "algorithm confusion" attack).
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
          reject(new HttpError(401, `Invalid token: ${error.message}`));
          return;
        }

        resolve(decoded as JwtPayload);
      }
    );
  });

  ensureAllowedClientApplication(payload);
  return payload;
}
