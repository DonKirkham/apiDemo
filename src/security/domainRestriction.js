const { allowedCallerDomains } = require('../config');

function getIncomingOrigin(request) {
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

function ensureAllowedDomain(request) {
  if (!allowedCallerDomains.length) {
    throw Object.assign(new Error('No ALLOWED_CALLER_DOMAINS configured.'), { status: 500 });
  }

  const incomingOrigin = getIncomingOrigin(request);
  if (!incomingOrigin) {
    throw Object.assign(new Error('Request origin or referer is required.'), { status: 403 });
  }

  if (!allowedCallerDomains.includes(incomingOrigin)) {
    throw Object.assign(new Error('Caller domain is not allowed.'), { status: 403 });
  }
}

module.exports = {
  ensureAllowedDomain
};
