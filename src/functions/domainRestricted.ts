import { ensureAllowedDomain } from '../security/domainRestriction.js';
import { registerCrudRoutes } from './helpers.js';

// /api/domain  and  /api/domain/{itemId}  — the HTTP method selects the CRUD
// action (GET=read, POST=create, PATCH=update, DELETE=delete).
// Accepts calls from configured SPFx origins only (ALLOWED_CALLER_DOMAINS).
registerCrudRoutes('domain', (request) => {
  ensureAllowedDomain(request);
  return { mode: 'domain-restricted' };
});
