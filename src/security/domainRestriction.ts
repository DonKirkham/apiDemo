import { HttpRequest } from '@azure/functions';
import { allowedCallerDomains } from '../config.js';
import { HttpError } from '../httpError.js';

// WARNING: this is NOT real authentication. Origin and Referer are plain HTTP
// headers that any client can set to any value, so a script outside the
// browser can fake them. This check only stops casual cross-site calls from
// browsers. For real protection use the Entra tier (entraAuth.ts).
function getIncomingOrigin(request: HttpRequest): string | null {
  const origin = request.headers.get('origin');
  if (origin) {
    return origin;
  }

  const referer = request.headers.get('referer');
  if (!referer) {
    return null;
  }

  try {
    return new URL(referer).origin;
  } catch {
    return null;
  }
}

export function ensureAllowedDomain(request: HttpRequest): void {
  if (!allowedCallerDomains.length) {
    throw new HttpError(500, 'No ALLOWED_CALLER_DOMAINS configured.');
  }

  const incomingOrigin = getIncomingOrigin(request);
  if (!incomingOrigin) {
    throw new HttpError(403, 'Request origin or referer is required.');
  }

  if (!allowedCallerDomains.includes(incomingOrigin)) {
    throw new HttpError(403, 'Caller domain is not allowed.');
  }
}
