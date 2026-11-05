import { spfi, SPFI } from '@pnp/sp';
import '@pnp/sp/webs/index.js';
import '@pnp/sp/lists/index.js';
import '@pnp/sp/items/index.js';
import { SPDefault } from '@pnp/nodejs';
import { AzureIdentity } from '@pnp/azidjsclient';
import { ManagedIdentityCredential } from '@azure/identity';
import { sharePointResource } from './config.js';
import { HttpError } from './httpError.js';

let cachedCredential: ManagedIdentityCredential | undefined;
const spCache = new Map<string, SPFI>();

// App-only credential via the Function App's managed identity. No secret or
// certificate is stored anywhere: Azure issues the token to the app's identity
// at runtime. The identity needs an application permission on SharePoint
// (e.g. Sites.ReadWrite.All), granted once with admin consent; see README.
function getCredential(): ManagedIdentityCredential {
  if (!cachedCredential) {
    cachedCredential = new ManagedIdentityCredential();
  }
  return cachedCredential;
}

// One SPFI instance per site. PnPjs binds an instance to a base URL, so when
// callers may target any site we build (and cache) an instance per site URL.
// `using(SPDefault(), AzureIdentity(...))` wires up the Node fetch/parse
// behaviors and acquires the app-only token from the credential above.
export function getSp(siteUrl: string): SPFI {
  if (!sharePointResource) {
    throw new HttpError(500, 'SHAREPOINT_RESOURCE is required.');
  }

  // The identity's token is sent to this URL, so only SharePoint hosts are
  // allowed. Without this check a caller could point the token at a server
  // they control.
  let host: string;
  try {
    host = new URL(siteUrl).hostname;
  } catch {
    throw new HttpError(400, 'siteUrl must be a valid URL.');
  }
  if (!host.endsWith('.sharepoint.com')) {
    throw new HttpError(400, 'siteUrl must be on a *.sharepoint.com host.');
  }

  const normalized = siteUrl.replace(/\/$/, '');
  const existing = spCache.get(normalized);
  if (existing) {
    return existing;
  }

  const sp = spfi(normalized).using(
    SPDefault(),
    AzureIdentity(getCredential(), [sharePointResource])
  );
  spCache.set(normalized, sp);
  return sp;
}

export interface ListItemTarget {
  siteUrl: string;
  listTitle: string;
}

export interface ItemTarget extends ListItemTarget {
  itemId: string | number;
}

export async function createSharePointListItem({
  siteUrl,
  listTitle,
  fields
}: ListItemTarget & { fields: Record<string, unknown> }): Promise<unknown> {
  return getSp(siteUrl).web.lists.getByTitle(listTitle).items.add(fields);
}

// Read a single item when `itemId` is provided, otherwise list items. Any OData
// system query options on the request ($select, $expand, $filter, $top,
// $orderby, …) are passed straight through to SharePoint rather than parsed and
// re-applied — the caller owns the query. PnPjs's fluent builders just set these
// same `$` params under the hood, so forwarding them is equivalent.
export async function readSharePointListItems({
  siteUrl,
  listTitle,
  itemId,
  query
}: ListItemTarget & { itemId?: string | number; query?: URLSearchParams }): Promise<unknown> {
  const items = getSp(siteUrl).web.lists.getByTitle(listTitle).items;
  const target =
    itemId !== undefined && itemId !== null && itemId !== ''
      ? items.getById(Number(itemId))
      : items;

  // Forward every OData system query option verbatim. (On a single-item read
  // SharePoint only honors $select/$expand; collection-only options are simply
  // ignored there.)
  if (query) {
    for (const [key, value] of query) {
      if (key.startsWith('$')) {
        target.query.set(key, value);
      }
    }
  }

  return target();
}

// PnPjs `update` is a MERGE (partial update) and defaults to IF-MATCH: *
// (overwrites regardless of the current ETag — last-write-wins). `delete`
// likewise defaults to IF-MATCH: *.
export async function updateSharePointListItem({
  siteUrl,
  listTitle,
  itemId,
  fields
}: ItemTarget & { fields: Record<string, unknown> }): Promise<void> {
  await getSp(siteUrl).web.lists.getByTitle(listTitle).items.getById(Number(itemId)).update(fields);
}

export async function deleteSharePointListItem({
  siteUrl,
  listTitle,
  itemId
}: ItemTarget): Promise<void> {
  await getSp(siteUrl).web.lists.getByTitle(listTitle).items.getById(Number(itemId)).delete();
}
