import { app, HttpRequest, HttpResponseInit } from '@azure/functions';
import { getMissingRequired, sharePointSiteUrl, sharePointListTitle } from '../config.js';
import {
  createSharePointListItem,
  readSharePointListItems,
  updateSharePointListItem,
  deleteSharePointListItem
} from '../sharepointClient.js';
import { HttpError } from '../httpError.js';

export interface CrudResult {
  status: number;
  operation: string;
  item?: unknown;
  itemId?: string;
}

// Turns any thrown error into an HTTP response. HttpError carries its own
// status; anything else is a 500. The error message is returned to the caller
// as-is. That is handy when learning, but a production API would hide it.
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
    const status = error instanceof HttpError ? error.status : 500;
    return {
      status,
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
// configured they win (locked-down mode) and any siteUrl/listTitle from the
// caller is ignored. Otherwise the caller picks the site and list per request.
// itemId comes from the route (e.g. .../domain/5) and must be a whole number.
export function resolveTarget(
  request: HttpRequest,
  body: RequestBody | undefined
): { siteUrl: string; listTitle: string; itemId?: string } {
  const query = request.query;
  const siteUrl = sharePointSiteUrl || body?.siteUrl || query.get('siteUrl') || undefined;
  const listTitle = sharePointListTitle || body?.listTitle || query.get('listTitle') || undefined;
  const itemId = request.params?.itemId;

  if (!siteUrl || !listTitle) {
    throw new HttpError(400, 'siteUrl and listTitle are required (in the body or query string).');
  }

  if (itemId !== undefined && !/^\d+$/.test(itemId)) {
    throw new HttpError(400, `itemId must be a whole number, got "${itemId}".`);
  }

  return { siteUrl, listTitle, itemId };
}

// Reads the JSON body for POST and PATCH. A malformed body is the caller's
// mistake, so it gets a 400 rather than a 500.
async function readJsonBody(request: HttpRequest): Promise<RequestBody> {
  try {
    return (await request.json()) as RequestBody;
  } catch {
    throw new HttpError(400, 'Request body must be valid JSON.');
  }
}

// Runs one CRUD operation. The HTTP method picks the action. Both route
// families use this, so create/read/update/delete behave the same no matter
// how the caller was authorized.
export async function dispatchCrudOperation(request: HttpRequest): Promise<CrudResult> {
  const method = request.method.toUpperCase();
  const hasBody = method === 'POST' || method === 'PATCH';
  const body = hasBody ? await readJsonBody(request) : undefined;
  const { siteUrl, listTitle, itemId } = resolveTarget(request, body);

  switch (method) {
    case 'POST': {
      if (!body?.fields || typeof body.fields !== 'object') {
        throw new HttpError(400, 'A fields object is required in the request body to create an item.');
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
        throw new HttpError(400, 'An itemId route parameter is required to update an item.');
      }
      if (!body?.fields || typeof body.fields !== 'object') {
        throw new HttpError(400, 'A fields object is required in the request body to update an item.');
      }
      await updateSharePointListItem({ siteUrl, listTitle, itemId, fields: body.fields });
      return { status: 200, operation: 'update', itemId };
    }

    case 'DELETE': {
      if (!itemId) {
        throw new HttpError(400, 'An itemId route parameter is required to delete an item.');
      }
      await deleteSharePointListItem({ siteUrl, listTitle, itemId });
      return { status: 200, operation: 'delete', itemId };
    }

    default:
      throw new HttpError(405, `Unsupported method: ${method}`);
  }
}

// Runs before every operation. It throws to reject the request. It returns
// extra fields to merge into the JSON response, such as the auth mode.
export type Authorizer = (
  request: HttpRequest
) => Promise<Record<string, unknown>> | Record<string, unknown>;

// Registers one auth tier's CRUD endpoint at two routes:
//   {prefix}           the collection (GET lists, POST creates)
//   {prefix}/{itemId}  one item (GET reads, PATCH updates, DELETE deletes)
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
