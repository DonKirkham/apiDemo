import { HttpRequest } from '@azure/functions';
import { allowedCallerDomains } from '../config.js';
import { httpError } from '../httpError.js';

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
    throw httpError(500, 'No ALLOWED_CALLER_DOMAINS configured.');
  }

  const incomingOrigin = getIncomingOrigin(request);
  if (!incomingOrigin) {
    throw httpError(403, 'Request origin or referer is required.');
  }

  if (!allowedCallerDomains.includes(incomingOrigin)) {
    throw httpError(403, 'Caller domain is not allowed.');
  }
}
