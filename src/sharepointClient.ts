import { spfi, SPFI } from '@pnp/sp';
import '@pnp/sp/webs/index.js';
import '@pnp/sp/lists/index.js';
import '@pnp/sp/items/index.js';
import { SPDefault } from '@pnp/nodejs';
import { AzureIdentity } from '@pnp/azidjsclient';
import { ClientCertificateCredential, ClientSecretCredential } from '@azure/identity';
import {
  tenantId,
  sharePointClientId,
  sharePointClientSecret,
  sharePointCertPem,
  sharePointResource
} from './config.js';
import { httpError } from './httpError.js';

type AppOnlyCredential = ClientCertificateCredential | ClientSecretCredential;

let cachedCredential: AppOnlyCredential | undefined;
const spCache = new Map<string, SPFI>();

// App-only credential. SharePoint REST only honors *certificate-based* Entra
// app-only tokens, so the certificate path is the one that actually works; a
// client secret acquires a token fine but the SharePoint call returns
// "401 Unsupported app only token". The secret path is kept only as a fallback
// for demonstrating that failure.
function getCredential(): AppOnlyCredential {
  if (cachedCredential) {
    return cachedCredential;
  }

  if (!tenantId || !sharePointClientId) {
    throw httpError(500, 'AZURE_TENANT_ID and SHAREPOINT_APP_CLIENT_ID are required.');
  }

  if (sharePointCertPem) {
    // `certificate` takes the full PEM (private key + certificate blocks),
    // exactly the format stored in the SHAREPOINT_CERT_PEM app setting.
    cachedCredential = new ClientCertificateCredential(tenantId, sharePointClientId, {
      certificate: sharePointCertPem
    });
  } else if (sharePointClientSecret) {
    cachedCredential = new ClientSecretCredential(tenantId, sharePointClientId, sharePointClientSecret);
  } else {
    throw httpError(500, 'SHAREPOINT_CERT_PEM or SHAREPOINT_APP_CLIENT_SECRET is required.');
  }

  return cachedCredential;
}

// One SPFI instance per site. PnPjs binds an instance to a base URL, so when
// callers may target any site we build (and cache) an instance per site URL.
// `using(SPDefault(), AzureIdentity(...))` wires up the Node fetch/parse
// behaviors and acquires the app-only token from the credential above.
export function getSp(siteUrl: string): SPFI {
  if (!sharePointResource) {
    throw httpError(500, 'SHAREPOINT_RESOURCE is required.');
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

// Read a single item when `itemId` is provided, otherwise list items. `odata`
// carries optional query options (select/filter/top/orderby) parsed from the
// request, applied the PnPjs way.
export async function readSharePointListItems({
  siteUrl,
  listTitle,
  itemId,
  odata
}: ListItemTarget & { itemId?: string | number; odata?: ODataOptions }): Promise<unknown> {
  const items = getSp(siteUrl).web.lists.getByTitle(listTitle).items;

  if (itemId !== undefined && itemId !== null && itemId !== '') {
    let query = items.getById(Number(itemId));
    if (odata?.select?.length) {
      query = query.select(...odata.select);
    }
    if (odata?.expand?.length) {
      query = query.expand(...odata.expand);
    }
    return query();
  }

  let query = items;
  if (odata?.select?.length) {
    query = query.select(...odata.select);
  }
  if (odata?.expand?.length) {
    query = query.expand(...odata.expand);
  }
  if (odata?.filter) {
    query = query.filter(odata.filter);
  }
  if (odata?.orderby) {
    query = query.orderBy(odata.orderby.field, odata.orderby.ascending);
  }
  if (odata?.top) {
    query = query.top(odata.top);
  }
  return query();
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

export interface ODataOptions {
  select?: string[];
  expand?: string[];
  filter?: string;
  top?: number;
  orderby?: { field: string; ascending: boolean };
}
