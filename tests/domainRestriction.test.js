const test = require('node:test');
const assert = require('node:assert/strict');

process.env.ALLOWED_CALLER_DOMAINS = 'https://contoso.sharepoint.com';

const { ensureAllowedDomain } = require('../src/security/domainRestriction');

function buildRequest({ origin, referer }) {
  return {
    headers: {
      get(name) {
        if (name === 'origin') return origin || null;
        if (name === 'referer') return referer || null;
        return null;
      }
    }
  };
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
