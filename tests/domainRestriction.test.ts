import test from 'node:test';
import assert from 'node:assert/strict';
import type { HttpRequest } from '@azure/functions';

process.env.ALLOWED_CALLER_DOMAINS = 'https://contoso.sharepoint.com';

const { ensureAllowedDomain } = await import('../src/security/domainRestriction.js');

function buildRequest({ origin, referer }: { origin?: string; referer?: string }): HttpRequest {
  return {
    headers: {
      get(name: string): string | null {
        if (name === 'origin') return origin || null;
        if (name === 'referer') return referer || null;
        return null;
      }
    }
  } as unknown as HttpRequest;
}

test('allows requests from configured origin header', () => {
  assert.doesNotThrow(() =>
    ensureAllowedDomain(buildRequest({ origin: 'https://contoso.sharepoint.com' }))
  );
});

test('rejects requests from unapproved origin', () => {
  assert.throws(
    () => ensureAllowedDomain(buildRequest({ origin: 'https://evil.example.com' })),
    /not allowed/
  );
});

test('falls back to referer origin', () => {
  assert.doesNotThrow(() =>
    ensureAllowedDomain(
      buildRequest({ referer: 'https://contoso.sharepoint.com/sites/portal/SitePages/home.aspx' })
    )
  );
});
