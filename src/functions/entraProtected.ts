import { validateEntraToken } from '../security/entraAuth.js';
import { registerCrudRoutes } from './helpers.js';

// /api/entra  and  /api/entra/{itemId}  — the HTTP method selects the CRUD
// action (GET=read, POST=create, PATCH=update, DELETE=delete).
// Requires and validates an Entra ID bearer token.
registerCrudRoutes('entra', async (request) => {
  const tokenPayload = await validateEntraToken(request);
  return {
    mode: 'entra-protected',
    caller: tokenPayload.azp || tokenPayload.appid || tokenPayload.sub
  };
});
