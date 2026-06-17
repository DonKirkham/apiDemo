# Demo Script — "Protect Your SPFx Solutions: Build Secure Backends with Azure Functions"

**Session length:** 70 minutes · **Speaker:** Don Kirkham · **Repo:** `github.com/donkirkham/apiDemo`
**Demo philosophy:** Everything is pre-built. Demos are *reveals*, not live coding. Talk to the **why**; let the running code prove it. Each demo below is timed and has a fallback for when the network/SharePoint misbehaves.

---

## ⚠️ Read this first — code↔slide alignment

The repo has been brought in line with the deck: it's now **TypeScript** (Node v4 model, ESM)
and talks to SharePoint through **PnPjs** (`@pnp/sp`). Two things still need attention:

| Topic | Status | Action before presenting |
|---|---|---|
| TypeScript on screen | ✅ Repo is TS now ([src/sharepointClient.ts](../src/sharepointClient.ts)) | none |
| PnPjs (`spfi().using(...)`) in the Function | ✅ Real — `SPDefault()` + `AzureIdentity(ClientCertificateCredential)` | none |
| **Client secret vs. certificate** | ⚠️ Slides (esp. **slide 11**) say *client secret*; code uses a **certificate** | **Edit the slides** — see "Slide changes" at the bottom. This is the cert-vs-secret aha (D2) |
| Separate function per verb (slide 6) **vs** combined (slide 7) | ✅ Code uses **one combined handler per tier** at `/api/{tier}/{itemId?}`, dispatched by HTTP method in `dispatchCrudOperation` ([helpers.ts:98](../src/functions/helpers.ts#L98)); both tiers are wired by one shared `registerCrudRoutes(...)` ([helpers.ts:161](../src/functions/helpers.ts#L161)) | Keep slide 7, delete slide 6 |
| Tier 1 / Tier 2 middleware | ✅ Concrete files: [domainRestriction.ts](../src/security/domainRestriction.ts), [entraAuth.ts](../src/security/entraAuth.ts) | Small enough to read on stage |

---

## Pre-flight checklist (do this before you walk on stage)

- [ ] `npm install` already run; `node_modules` present.
- [ ] `local.settings.json` populated from `local.settings.sample.json` with **real** values for the `pdslabs2` tenant (cert PEM in `SHAREPOINT_CERT_PEM`, `SHAREPOINT_SITE_URL`/`SHAREPOINT_LIST_TITLE` pinned to the root site's `Requests` list).
- [ ] `func start` works locally → `http://localhost:7071/api/health` returns `{ "ok": true }`.
- [ ] Live app is warm: open `https://apidemo-func-4r1iq2.azurewebsites.net/api/health` once so the first cold start is paid before the audience is watching.
- [ ] A **valid Entra bearer token** for the Tier-2 endpoint copied somewhere pasteable (for D6/jwt.ms), plus **one deliberately-wrong token** (a Graph token, or a token with the wrong `aud`).
- [ ] Postman/Thunder Client collection open with the four CRUD calls + a "no-origin" call pre-built.
- [ ] `jwt.ms` open in a tab.
- [ ] Azure Portal open and **already signed in** to the `pdslabs2` tenant: the server app registration, the Function App (`apidemo-func-4r1iq2`), and App Insights blades pinned.
- [ ] Font size cranked; terminal + editor at presentation zoom.
- [ ] **Secrets hygiene:** the cert PEM and any real tokens must NOT be visible on screen. Have `local.settings.json` closed; if you must show it, show `local.settings.sample.json` instead.

**If the venue Wi-Fi dies:** D1, D2, D4 (the `npm test` variant), D6 (jwt.ms is the only online bit — have a screenshot ready), and D9 all run **locally or offline**. D3, D5, D8, D10 need the network — fall back to the recorded screen capture / screenshots noted in each.

---

## Time budget (70 min)

| Block | Section | Minutes | Demos | Priority |
|---|---|---:|---|---|
| Intro | Title / About / Roadmap (1–3) | 4 | — | — |
| 1 | Azure Functions Orientation (4–8) | 9 | **D1** | core |
| 2 | App-only auth in Node (9–11) | 10 | **D2** ⭐ | core |
| 3 | Tier 1 — Origin-validated (12–15) | 12 | **D3**, **D4** ⭐ | D4 core, D3 optional |
| 4 | Tier 2 — Entra ID (16–19) | 14 | D5, **D6** ⭐ | D6 core, D5 optional |
| 5 | Calling both from SPFx (20–23) | 12 | D7, **D8** ⭐, D9 | D8 core |
| 6 | Deploy & Harden (24–25) | 6 | **D10** | core |
| Recap | Recap & Resources (26) | 3 | — | — |

⭐ = the four "aha" demos. If you fall behind, **protect D2, D4, D6, D8** and cut D3/D5/D9 to a single sentence.

> SPFx-side demos (D7/D8/D9) consume the API but the **web part code lives in the companion SPFx repo**, not this one. Have that solution open in a second VS Code window / second machine, or use recorded clips.

---

# Section 1 — Azure Functions Orientation

## D1 — The local dev loop · slide 8 · ~3 min · risk: LOW (all local)

**Goal:** prove the whole inner loop runs on the laptop before a dollar reaches Azure.

**Pre-state:** terminal at repo root; [src/functions/health.ts](../src/functions/health.ts) and [helpers.ts](../src/functions/helpers.ts) open; a breakpoint set on the `switch (method)` line in `dispatchCrudOperation` ([helpers.ts:104](../src/functions/helpers.ts#L104)).

**Steps**
1. `npm start` (this is just `func start`). Point at the terminal as it prints the function map:
   - `health: [GET] http://localhost:7071/api/health`
   - `domain: [GET,POST,PATCH,DELETE] .../api/domain/{itemId?}`
   - `entra: [GET,POST,PATCH,DELETE] .../api/entra/{itemId?}`
2. Hit `http://localhost:7071/api/health` in the browser → `{ "ok": true }`.
3. F5 (or "Attach to Node Functions") → fire a `GET /api/domain` from Postman → VS Code **breaks inside the handler**. Hover `request.method`, step once, show it landing in the `GET` case.
4. Stop. `npm test` → the three [domainRestriction.test.ts](../tests/domainRestriction.test.ts) cases go green in ~1s.

**Say:** "Same machine, full debugger, real breakpoints, unit tests — and Azure hasn't been billed a cent. The Functions extension scaffolds this; `func start` is the local runtime; each function gets its own URL. This is the loop you live in 90% of the time."

**Aha:** the function map printing on startup — "the runtime *discovers* my functions from `app.http(...)` calls; there's no routing config to maintain. Both CRUD endpoints come from two `registerCrudRoutes(...)` calls — one per auth tier."

**Fallback:** if `func start` won't bind, run `npm test` alone — it needs nothing but Node and still demonstrates the inner loop.

**Reset:** `Ctrl+C` the host; clear the breakpoint before D6 so you don't break unexpectedly later.

---

# Section 2 — App-only auth in a Node context

## D2 — App-only auth with PnPjs, and *why a certificate* ⭐ · slide 11 · ~5 min · risk: LOW–MED

**Goal:** show that PnPjs in a Function is the *same API* as in the browser — you only swap in an explicit app credential — and deliver the gotcha: **SharePoint REST rejects secret-based app-only tokens; you need a certificate.**

**Pre-state:** [src/sharepointClient.ts](../src/sharepointClient.ts) open; the `getCredential()` cert-vs-secret branch and the `getSp()` `spfi().using(...)` setup visible; README "Required configuration" cert note handy.

**Steps**
1. Show `getSp()`: `spfi(siteUrl).using(SPDefault(), AzureIdentity(getCredential(), [sharePointResource]))`. "This is the exact `spfi().using(...)` shape from the SPFx slide — the only difference in Node is *which* credential I hand it."
2. Show `getCredential()`: if `SHAREPOINT_CERT_PEM` is set → **`ClientCertificateCredential`**; else fall back to `ClientSecretCredential`. The PEM (key + cert) is read straight from the `SHAREPOINT_CERT_PEM` app setting — no hand-rolled JWT, PnPjs/@azure-identity mint the token.
3. **The reveal:** "The secret branch is there as a fallback — and it *gets a token just fine*. But when that token hits SharePoint REST you get `401 Unsupported app only token`. SharePoint only honors **certificate-based** app-only tokens. That one line cost me an afternoon." (Optional live: set only the secret and show the 401 — needs a safe secret.)
4. Show the CRUD call site in [helpers.ts](../src/functions/helpers.ts) / [sharepointClient.ts](../src/sharepointClient.ts): `getSp(siteUrl).web.lists.getByTitle(listTitle).items.add(fields)` — "identical to what you'd write in a web part."

**Say:** "In the browser PnPjs gets its identity from the page context. In a Function there's no user, so I provide an app identity explicitly. Registration is a daemon app, no redirect URI, **`Sites.ReadWrite.All`** application permission, admin-consented — enough to read/write items in an existing list, not to create lists."

**Aha:** same PnPjs API, explicit credential — and cert > secret for SharePoint app-only.

**Fallback:** pure code read — no network needed. Keep the README cert note on screen as proof.

**Reset:** restore `local.settings.json` to the cert if you flipped it.

---

# Section 3 — Tier 1: Origin-Validated API

## D3 — Tier 1 CRUD against a real list · slide 14 · ~6 min · risk: MED (live SharePoint)

**Goal:** show full CRUD flowing through one combined handler into a real SharePoint list.

**Pre-state:** Postman collection against the **live** app (`https://apidemo-func-4r1iq2.azurewebsites.net`) with `Origin: https://<your-allowed-domain>` set on every request. The `Requests` list open in a browser tab. App is pinned (`SHAREPOINT_SITE_URL`/`SHAREPOINT_LIST_TITLE` set), so the body only needs `fields`.

**Steps** (one combined handler, dispatched by HTTP method — [helpers.ts:98](../src/functions/helpers.ts#L98))
1. **POST** `/api/domain` body `{ "fields": { "Title": "Created on stage", "RequestType": "Demo" } }` → `201`, note the returned item `Id`.
2. Switch to the SharePoint tab, refresh → the row is there.
3. **GET** `/api/domain?$select=Id,Title,RequestType&$top=5` → the new row in JSON.
4. **PATCH** `/api/domain/{Id}` body `{ "fields": { "Title": "Updated on stage" } }` → `200`.
5. **DELETE** `/api/domain/{Id}` → `200`; refresh SharePoint → gone.

**Say:** "One handler, four verbs — no action words in the URL. The HTTP method *is* the operation: `dispatchCrudOperation` switches on it, and one `registerCrudRoutes` call wires up both tiers. Create/read/update/delete behave identically no matter how the caller was authorized, which is exactly why Tier 1 and Tier 2 share it. Updates and deletes are an unconditional `IF-MATCH: *` — last write wins."

**Aha:** the same code object serves both tiers; auth is a thin gate in front of shared logic.

**Fallback:** point at the local `func start` instance instead of Azure; or play a 60s screen capture of the round-trip.

**Reset:** delete any leftover stage rows so the next run is clean.

---

## D4 — Origin validation in action ⭐ · slides 13 & 15 · ~5 min · risk: LOW

**Goal:** Tier 1 is a **bouncer checking where you came from**, not identity — and it's spoofable, which is *the point* of having Tier 2.

**Pre-state:** [src/security/domainRestriction.ts](../src/security/domainRestriction.ts) open (it's ~35 lines — read it on screen). Postman with three pre-built `GET /api/domain` calls: (a) `Origin: https://contoso.sharepoint.com` [allowed], (b) **no** Origin/Referer, (c) `Origin: https://evil.example.com`.

**Steps**
1. Read `ensureAllowedDomain` ([domainRestriction.ts:23](../src/security/domainRestriction.ts#L23)): pulls `origin`, falls back to `referer`'s origin, 403s if not in `ALLOWED_CALLER_DOMAINS`.
2. Run call (a) → `200` + data.
3. Run call (b) no origin → **403** "Request origin or referer is required."
4. Run call (c) wrong origin → **403** "Caller domain is not allowed."
5. **The honest catch:** hand-add `Origin: https://contoso.sharepoint.com` to a curl/Postman call → **200**. "A browser can't lie about Origin. `curl` can. So this protects against *casual* cross-origin calls and random scanners — not a determined client. That's why it's for low-risk reads only."
6. (Optional, offline) `npm test` → [domainRestriction.test.ts](../tests/domainRestriction.test.ts) shows the allow/reject/referer-fallback cases codified.

**Say:** "Zero friction for a real SPFx web part — the browser sends Origin automatically. Zero identity guarantee for anything else. Know which one you're buying."

**Aha:** the spoof. It's what motivates the entire Tier 2 half of the talk — land it deliberately.

**Fallback:** the `npm test` variant proves the same logic with no network.

---

# Section 4 — Tier 2: Entra ID Protected API

## D5 — Expose an API (app registrations) · slide 18 · ~5 min · risk: MED (portal)

**Goal:** show the two-app dance: a **server** app that exposes a scope, a **client** (SPFx) app that requests it.

**Pre-state:** Portal signed into `pdslabs2`. Server app registration open at **Expose an API**; the SPFx client app's **API permissions** blade in a second tab.

**Steps**
1. Server app → **Expose an API**: show the **Application ID URI** `api://{client-id}` and the custom scope (e.g. `SharePoint.ReadWrite`), "Admins and users."
2. Client app → **API permissions** → My APIs → server app → the scope → **admin consent granted** (green check).
3. Tie it to config: this URI is what becomes `ENTRA_AUDIENCE` ([config.ts:36](../src/config.ts#L36)) in the Function, and what SPFx's `getClient('api://…')` asks for (D8).

**Say:** "Two registrations. The server *publishes* a scope; the client *requests* it; an admin *consents* once. After that, SPFx can silently acquire tokens for my API."

**Aha:** the `aud` on the token the API validates is literally this App ID URI — connect it forward to D6.

**Fallback:** screenshots of both blades. Portal nav on stage is slow; pre-screenshot if you're tight.

---

## D6 — Token validation: what claims matter ⭐ · slide 19 · ~6 min · risk: LOW

**Goal:** make token validation concrete — four claims decide everything, and the code checks all four.

**Pre-state:** `jwt.ms` open. A **valid** bearer token and a **wrong** token (Graph token or wrong `aud`) ready to paste. [src/security/entraAuth.ts](../src/security/entraAuth.ts) open.

**Steps**
1. Paste the valid token into jwt.ms. Highlight:
   - **`iss`** — must be your tenant's v2 endpoint (`…/{tenant}/v2.0`).
   - **`aud`** — must equal your API's App ID URI. *A Microsoft Graph token will not pass.*
   - **`azp` / `appid`** — must be your SPFx client app.
   - **`scp`** — the scope you required is present.
2. Map each to the code in `validateEntraToken` ([entraAuth.ts:55](../src/security/entraAuth.ts#L55)): `jwt.verify` enforces `algorithms: ['RS256']`, `audience`, `issuer`; the signing key comes from **JWKS** via `jwks-rsa` (cached, [entraAuth.ts:15](../src/security/entraAuth.ts#L15)); then `ensureAllowedClientApplication` ([entraAuth.ts:44](../src/security/entraAuth.ts#L44)) enforces the `azp`/`appid` allowlist.
3. **The reject:** fire `GET /api/entra` with the wrong-`aud` token → **401 "Invalid token: …"**. "Right signature, real Entra token — wrong audience. Rejected. That one check kills a whole class of token-replay attacks where someone hands my API a token that was minted for something else."
4. Show a **missing** token → 401 "Authorization bearer token is required" ([entraAuth.ts:58](../src/security/entraAuth.ts#L58)).

**Say:** "Validation isn't 'is there a token.' It's: signed by my tenant's keys, minted *for me*, by an app *I* allow, carrying the scope I need. Four claims, four checks, ~50 lines."

**Aha:** the wrong-audience 401. A token can be 100% valid and still not be *for you*.

**Fallback:** jwt.ms is the only online piece — keep a screenshot of a decoded token. The code read + the 401 against the local host need no internet.

---

# Section 5 — Calling Both APIs from SPFx

> Web part code is in the **companion SPFx repo**. Have it open separately or use clips.

## D7 — Tier 1 from SPFx: HttpClient · slide 21 · ~3 min · risk: LOW

**Steps**
1. Show the web part calling `this.context.httpClient.get(API_BASE_URL + '/api/domain', HttpClient.configurations.v1)`.
2. Render the list data in the web part. No token, no manifest entry — the browser attached the Origin header for free.
3. Show graceful handling of `!response.ok` surfacing the 403 from D4 as a real message.

**Say:** "Standard SPFx client, nothing special. The Origin header *is* the credential here — which is exactly as much trust as Tier 1 earns."

**Fallback:** clip of the web part rendering rows.

## D8 — Tier 2 from SPFx: AadHttpClient ⭐ · slide 22 · ~6 min · risk: HIGH (token + admin + CORS)

**Goal:** the headline "one line of difference," plus the governance catch.

**Steps**
1. Manifest: `webApiPermissionRequests: [{ resource: 'api://{server-client-id}', scope: 'SharePoint.ReadWrite' }]`.
2. Code: `this.context.aadHttpClientFactory.getClient('api://{server-client-id}')` then `aadClient.get(SECURE_API_URL + '/api/entra', AadHttpClient.configurations.v1)`.
3. Open the browser **Network tab**, fire the call, expand the request → **`Authorization: Bearer …` added automatically**. Paste that token into jwt.ms to close the loop with D6.
4. **The catch:** show the **SharePoint Admin Center → API access** approval screen. "Until an admin approves this permission post-deploy, *every* call 401s. This is the step people forget."

**Say:** "I never touch a token. SPFx acquires and caches it against the scope I declared. The price of that magic is one admin approval and getting the App ID URI exactly right."

**Aha:** the Bearer header appearing in the Network tab that you *never wrote code to add*.

**Fallback:** this is the riskiest live demo — **have a recorded clip** of (network tab + admin approval). Don't gamble the headline on live token acquisition + CORS.

## D9 — Side-by-side: what actually changes · slide 23 · ~2 min · risk: LOW

**Steps:** put the two web-part call sites side by side (split editor or the slide-23 table). Walk the diff: import, get-client, the call, token (none vs auto-Bearer), manifest entry, admin step. "Six rows in the table; in the code it's about one meaningful line plus a manifest entry."

**Fallback:** the slide-23 table *is* the fallback.

---

# Section 6 — Deploy & Harden

## D10 — Deploy & hardening pass · slide 25 · ~5 min · risk: MED (live deploy)

**Goal:** ship it and point at the three production knobs.

**Pre-state:** terminal at repo root, Azure CLI logged into the `VS Enterprise 2021` subscription; Portal on `apidemo-func-4r1iq2`.

**Steps**
1. Deploy (or show the last deploy's output):
   ```bash
   func azure functionapp publish apidemo-func-4r1iq2 --javascript --build remote
   ```
   **Call out `--build remote` and why:** `.funcignore` excludes `node_modules`, and **Flex Consumption does not support `SCM_DO_BUILD_DURING_DEPLOYMENT`** — so the build must be requested at deploy time, server-side via Oryx. Skip it and the app deploys with no dependencies and **registers zero functions** (everything 404s). This is the #1 first-deploy trap.
2. Verify live: `https://apidemo-func-4r1iq2.azurewebsites.net/api/health` → `{ "ok": true }`.
3. **Key Vault:** in App Settings, show a secret referenced as `@Microsoft.KeyVault(SecretUri=…)` — same env var the code reads, zero code change. (Better yet: "next level is Managed Identity and no secret at all.")
4. **App Insights:** open **Live Metrics**, fire a request, watch it appear — traces, failures, durations, one checkbox at provision time.
5. **CORS:** show the Function App CORS list — SharePoint domains only, **no `*`**. Locally it's the Workbench origin in `host.json`.

**Say:** "Flex Consumption for warm starts, Node 22 to match what SPFx supports, remote build because the SKU demands it. Key Vault + App Insights is my minimum-viable production bar; cert auth + Managed Identity is the next rung."

**Aha:** the `--build remote` requirement — it's non-obvious, costs people their first deploy, and you can save the room from it in one sentence.

**Fallback:** don't deploy live if the room is slow — show the previous deploy's terminal output + the live `/api/health` response, then the Portal blades.

---

## Cut-down order if you're behind

Running long is the default failure mode. Drop in this order, each to a single spoken sentence:
1. D9 (table speaks for itself)
2. D5 (one screenshot)
3. D3 (mention D4 already proved the route works)
4. D7 (D8 implies it)

**Never cut:** D2 (cert gotcha), D4 (the spoof), D6 (wrong-audience 401), D8 (auto-Bearer). Those four are the talk.

## Slide changes to make before presenting

The code now matches the deck's TypeScript + PnPjs story. Two edits remain so the slides
match the code:

1. **Slide 11 — "App-Only Auth: Client Secret Setup" → "Certificate Setup".** This is the
   biggest one. The code authenticates with a **certificate**, not a secret, because
   SharePoint REST rejects secret-based app-only tokens (`401 Unsupported app only token`).
   - Step 3 "Generate client secret" → **"Generate & upload a certificate"**
     (`az ad app credential reset --id <clientId> --create-cert --append`); store the PEM
     in `SHAREPOINT_CERT_PEM`.
   - Code block → the PnPjs + cert shape now in the repo:
     ```ts
     import { spfi } from "@pnp/sp";
     import { SPDefault } from "@pnp/nodejs";
     import { AzureIdentity } from "@pnp/azidjsclient";
     import { ClientCertificateCredential } from "@azure/identity";

     const credential = new ClientCertificateCredential(tenantId, clientId, { certificate: pem });
     export const getSp = (siteUrl: string) =>
       spfi(siteUrl).using(SPDefault(), AzureIdentity(credential, [sharePointResource]));
     ```
   - Add a one-line callout: *"A client secret acquires a token but SharePoint rejects it —
     app-only to SharePoint REST must be certificate-based."* (This is demo **D2**.)

2. **Slides 6 vs 7 — keep one.** The repo uses the **combined-handler** pattern: one route
   per tier, `/api/{tier}/{itemId?}` with `methods: ['GET','POST','PATCH','DELETE']`, and the
   **HTTP method maps to the CRUD action** inside `dispatchCrudOperation`
   ([helpers.ts:98](../src/functions/helpers.ts#L98)) — there are no action words in the URL.
   Both tiers are wired by one shared `registerCrudRoutes(prefix, authorize)`
   ([helpers.ts:161](../src/functions/helpers.ts#L161)). Keep slide 7, delete slide 6, and
   align the annotation to "one handler per tier, dispatch by HTTP method."

Optional polish: slide 10's "AzureIdentityFetchClient or MSAL" line can be simplified to
"`@azure/identity` credential via the PnPjs `AzureIdentity` behavior," which is what the
code uses.
