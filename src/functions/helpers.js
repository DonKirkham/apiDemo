const { getMissingRequired } = require('../config');

function parseJsonBody(request) {
  return request.json();
}

async function runWithErrorHandling(operation) {
  try {
    const missing = getMissingRequired();
    if (missing.length) {
      return {
        status: 500,
        jsonBody: {
          error: `Missing required app settings: ${missing.join(', ')}`
        }
      };
    }

    return await operation();
  } catch (error) {
    return {
      status: error.status || 500,
      jsonBody: {
        error: error.message || 'Unexpected error.'
      }
    };
  }
}

function validatePayload(payload) {
  if (!payload || typeof payload !== 'object') {
    throw Object.assign(new Error('JSON body is required.'), { status: 400 });
  }

  if (!payload.siteUrl || !payload.listTitle || !payload.fields || typeof payload.fields !== 'object') {
    throw Object.assign(
      new Error('siteUrl, listTitle, and fields object are required in the request body.'),
      { status: 400 }
    );
  }
}

module.exports = {
  parseJsonBody,
  runWithErrorHandling,
  validatePayload
};
