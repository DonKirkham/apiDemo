import { HttpRequest } from '@azure/functions';
import jwt, { JwtHeader, SigningKeyCallback, JwtPayload } from 'jsonwebtoken';
import jwksClient from 'jwks-rsa';
import { allowedClientAppIds, entraAudience, entraIssuer, entraJwksUri } from '../config.js';
import { httpError } from '../httpError.js';

let client: ReturnType<typeof jwksClient> | undefined;

function getClient(): ReturnType<typeof jwksClient> {
  if (!client) {
    if (!entraJwksUri) {
      throw httpError(500, 'ENTRA_JWKS_URI is required for Entra validation.');
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

function ensureAllowedClientApplication(payload: JwtPayload): void {
  if (!allowedClientAppIds.length) {
    return;
  }

  const appId = (payload.azp as string | undefined) || (payload.appid as string | undefined);
  if (!appId || !allowedClientAppIds.includes(appId)) {
    throw httpError(403, 'Calling application is not allowed.');
  }
}

export async function validateEntraToken(request: HttpRequest): Promise<JwtPayload> {
  const token = getBearerToken(request);
  if (!token) {
    throw httpError(401, 'Authorization bearer token is required.');
  }

  if (!entraAudience || !entraIssuer) {
    throw httpError(500, 'ENTRA_AUDIENCE and tenant settings are required.');
  }

  const payload = await new Promise<JwtPayload>((resolve, reject) => {
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
          reject(httpError(401, `Invalid token: ${error.message}`));
          return;
        }

        resolve(decoded as JwtPayload);
      }
    );
  });

  ensureAllowedClientApplication(payload);
  return payload;
}
