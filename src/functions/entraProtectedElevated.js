const { app } = require('@azure/functions');
const { validateEntraToken } = require('../security/entraAuth');
const { createSharePointListItem } = require('../sharepointClient');
const { parseJsonBody, runWithErrorHandling, validatePayload } = require('./helpers');

app.http('entraProtectedElevated', {
  route: 'elevated/entra',
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request) =>
    runWithErrorHandling(async () => {
      const tokenPayload = await validateEntraToken(request);

      const body = await parseJsonBody(request);
      validatePayload(body);

      const item = await createSharePointListItem(body);

      return {
        status: 200,
        jsonBody: {
          mode: 'entra-protected',
          caller: tokenPayload.azp || tokenPayload.appid || tokenPayload.sub,
          item
        }
      };
    })
});
