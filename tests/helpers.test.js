const test = require('node:test');
const assert = require('node:assert/strict');

const { validatePayload } = require('../src/functions/helpers');

test('validatePayload accepts expected request body shape', () => {
  assert.doesNotThrow(() =>
    validatePayload({
      siteUrl: 'https://contoso.sharepoint.com/sites/hr',
      listTitle: 'Requests',
      fields: { Title: 'Test' }
    })
  );
});

test('validatePayload rejects missing fields object', () => {
  assert.throws(
    () =>
      validatePayload({
        siteUrl: 'https://contoso.sharepoint.com/sites/hr',
        listTitle: 'Requests'
      }),
    /required/
  );
});
