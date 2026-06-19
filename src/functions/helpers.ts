import { app, HttpRequest, HttpResponseInit } from '@azure/functions';
import { getMissingRequired, sharePointSiteUrl, sharePointListTitle } from '../config.js';
import {
  createSharePointListItem,
  readSharePointListItems,
  updateSharePointListItem,
  deleteSharePointListItem
} from '../sharepointClient.js';
import { httpError, statusOf } from '../httpError.js';

export interface CrudResult {
  status: number;
  operation: string;
  item?: unknown;
  itemId?: string;
}

export async function runWithErrorHandling(
  operation: () => Promise<HttpResponseInit>
): Promise<HttpResponseInit> {
  try {
    const missing = getMissingRequired();
    if (missing.length) {
      return {
        status: 500,
        jsonBody: { error: `Missing required app settings: ${missing.join(', ')}` }
      };
    }

    return await operation();
  } catch (error) {
    return {
      status: statusOf(error),
      jsonBody: { error: (error as Error).message || 'Unexpected error.' }
    };
  }
}

interface RequestBody {
  siteUrl?: string;
  listTitle?: string;
  fields?: Record<string, unknown>;
}

// Resolve the target list. When SHAREPOINT_SITE_URL / SHAREPOINT_LIST_TITLE are
// configured they are authoritative (locked-down mode) and any client-supplied
// siteUrl/listTitle is ignored. With them unset, callers may target any
// site/list per request via the body or query string. itemId comes from the
// route (e.g. .../domain/5).
export function resolveTarget(
  request: HttpRequest,
  body: RequestBody | undefined
): { siteUrl: string; listTitle: string; itemId?: string } {
  const query = request.query;
  const siteUrl = sharePointSiteUrl || body?.siteUrl || query.get('siteUrl') || undefined;
  const listTitle = sharePointListTitle || body?.listTitle || query.get('listTitle') || undefined;
  const itemId = request.params?.itemId;

  if (!siteUrl || !listTitle) {
    throw httpError(400, 'siteUrl and listTitle are required (in the body or query string).');
  }

  return { siteUrl, listTitle, itemId };
}

// Dispatches one CRUD operation based on the HTTP method. Shared by every
// auth-protected endpoint so the create/read/update/delete behavior stays
// identical regardless of how the caller was authorized.
export async function dispatchCrudOperation(request: HttpRequest): Promise<CrudResult> {
  const method = request.method.toUpperCase();
  const hasBody = method === 'POST' || method === 'PATCH';
  const body = hasBody ? ((await request.json()) as RequestBody) : undefined;
  const { siteUrl, listTitle, itemId } = resolveTarget(request, body);

  switch (method) {
    case 'POST': {
      if (!body?.fields || typeof body.fields !== 'object') {
        throw httpError(400, 'A fields object is required in the request body to create an item.');
      }
      const item = await createSharePointListItem({ siteUrl, listTitle, fields: body.fields });
      return { status: 201, operation: 'create', item };
    }

    case 'GET': {
      const item = await readSharePointListItems({
        siteUrl,
        listTitle,
        itemId,
        query: request.query
      });
      return { status: 200, operation: 'read', item };
    }

    case 'PATCH': {
      if (!itemId) {
        throw httpError(400, 'An itemId route parameter is required to update an item.');
      }
      if (!body?.fields || typeof body.fields !== 'object') {
        throw httpError(400, 'A fields object is required in the request body to update an item.');
      }
      await updateSharePointListItem({ siteUrl, listTitle, itemId, fields: body.fields });
      return { status: 200, operation: 'update', itemId };
    }

    case 'DELETE': {
      if (!itemId) {
        throw httpError(400, 'An itemId route parameter is required to delete an item.');
      }
      await deleteSharePointListItem({ siteUrl, listTitle, itemId });
      return { status: 200, operation: 'delete', itemId };
    }

    default:
      throw httpError(405, `Unsupported method: ${method}`);
  }
}

// Runs before every operation. Throws to reject the request; returns any extra
// fields to merge into the JSON response (e.g. the auth mode and caller).
export type Authorizer = (
  request: HttpRequest
) => Promise<Record<string, unknown>> | Record<string, unknown>;

// Registers the CRUD endpoint for one auth tier at two routes:
//   {prefix}            — the collection
//   {prefix}/{itemId}   — a single item
// There are no action words in the URL; the HTTP method maps to the CRUD
// action (dispatchCrudOperation does the switch):
//   GET    -> read    (list when no itemId, one item when itemId is present)
//   POST   -> create
//   PATCH  -> update  (requires itemId)
//   DELETE -> delete  (requires itemId)
// `authorize` runs before every operation; it throws to reject and returns any
// extra fields to merge into the JSON response (e.g. the auth mode and caller).
export function registerCrudRoutes(prefix: string, authorize: Authorizer): void {
  app.http(prefix, {
    route: `${prefix}/{itemId?}`,
    methods: ['GET', 'POST', 'PATCH', 'DELETE'],
    authLevel: 'anonymous',
    handler: (request: HttpRequest) =>
      runWithErrorHandling(async () => {
        const extra = await authorize(request);
        const { status, ...result } = await dispatchCrudOperation(request);

        return {
          status,
          jsonBody: { ...extra, ...result }
        };
      })
  });
}
