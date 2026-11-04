import { app } from '@azure/functions';
import { buildInfo } from '../buildInfo.js';

// Process start time: changes on every cold start or restart.
const startedAt = new Date().toISOString();

app.http('health', {
  route: 'health',
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async () => ({
    status: 200,
    jsonBody: {
      ok: true,
      commit: buildInfo.commit,
      builtAt: buildInfo.builtAt,
      startedAt,
      sharePointAuth: 'managed-identity'
    }
  })
});
