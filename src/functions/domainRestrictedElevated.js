const { app } = require('@azure/functions');
const { ensureAllowedDomain } = require('../security/domainRestriction');
const { createSharePointListItem } = require('../sharepointClient');
const { parseJsonBody, runWithErrorHandling, validatePayload } = require('./helpers');

app.http('domainRestrictedElevated', {
  route: 'elevated/domain',
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request) =>
    runWithErrorHandling(async () => {
      ensureAllowedDomain(request);

      const body = await parseJsonBody(request);
      validatePayload(body);

      const item = await createSharePointListItem(body);

      return {
        status: 200,
        jsonBody: {
          mode: 'domain-restricted',
          item
        }
      };
    })
});
