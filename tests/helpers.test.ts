import test from 'node:test';
import assert from 'node:assert/strict';
import type { HttpRequest } from '@azure/functions';
import { resolveTarget, dispatchCrudOperation } from '../src/functions/helpers.js';
import { HttpError } from '../src/httpError.js';

function fakeRequest({
  query = {},
  params = {}
}: { query?: Record<string, string>; params?: Record<string, string> } = {}): HttpRequest {
  return {
    query: new URLSearchParams(query),
    params
  } as unknown as HttpRequest;
}

test('resolveTarget reads siteUrl and listTitle from the body', () => {
  const target = resolveTarget(fakeRequest({ params: { itemId: '5' } }), {
    siteUrl: 'https://contoso.sharepoint.com/sites/hr',
    listTitle: 'Requests'
  });

  assert.deepEqual(target, {
    siteUrl: 'https://contoso.sharepoint.com/sites/hr',
    listTitle: 'Requests',
    itemId: '5'
  });
});

test('resolveTarget falls back to the query string', () => {
  const target = resolveTarget(
    fakeRequest({
      query: {
        siteUrl: 'https://contoso.sharepoint.com/sites/hr',
        listTitle: 'Requests'
      }
    }),
    undefined
  );

  assert.equal(target.siteUrl, 'https://contoso.sharepoint.com/sites/hr');
  assert.equal(target.listTitle, 'Requests');
  assert.equal(target.itemId, undefined);
});

test('resolveTarget throws when siteUrl or listTitle is missing', () => {
  assert.throws(() => resolveTarget(fakeRequest(), {}), /siteUrl and listTitle are required/);
});

test('resolveTarget rejects a non-numeric itemId with 400', () => {
  assert.throws(
    () =>
      resolveTarget(fakeRequest({ params: { itemId: 'abc' } }), {
        siteUrl: 'https://contoso.sharepoint.com/sites/hr',
        listTitle: 'Requests'
      }),
    (error: unknown) => error instanceof HttpError && error.status === 400
  );
});

test('dispatchCrudOperation rejects a malformed JSON body with 400', async () => {
  const request = {
    method: 'POST',
    query: new URLSearchParams(),
    params: {},
    json: async () => {
      throw new SyntaxError('Unexpected token');
    }
  } as unknown as HttpRequest;

  await assert.rejects(dispatchCrudOperation(request), (error: unknown) => {
    return error instanceof HttpError && error.status === 400;
  });
});
