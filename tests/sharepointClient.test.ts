import test from 'node:test';
import assert from 'node:assert/strict';

// config.ts reads process.env on first import, so set SHAREPOINT_RESOURCE first.
// getSp() needs it, and the host check runs after that check.
process.env.SHAREPOINT_RESOURCE = 'https://contoso.sharepoint.com/.default';

const { getSp } = await import('../src/sharepointClient.js');
const { HttpError } = await import('../src/httpError.js');

test('getSp rejects a site URL that is not on sharepoint.com', () => {
  assert.throws(
    () => getSp('https://evil.example.com/sites/hr'),
    (error: unknown) =>
      error instanceof HttpError && error.status === 400 && /sharepoint\.com/.test(error.message)
  );
});

test('getSp rejects a site URL that merely contains sharepoint.com', () => {
  assert.throws(
    () => getSp('https://contoso.sharepoint.com.evil.example.com/sites/hr'),
    (error: unknown) => error instanceof HttpError && error.status === 400
  );
});

test('getSp rejects a value that is not a URL', () => {
  assert.throws(
    () => getSp('not a url'),
    (error: unknown) => error instanceof HttpError && error.status === 400
  );
});
