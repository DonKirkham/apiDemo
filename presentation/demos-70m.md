# Demo Script — "Protect Your SPFx Solutions: Build Secure Backends with Azure Functions"

**Session length:** 70 minutes · **Speaker:** Don Kirkham · **Repo:** `github.com/donkirkham/apiDemo`
**Demo philosophy:** Everything is pre-built. Demos are *reveals*, not live coding. Talk to the **why**; let the running code prove it. Each demo below is timed and has a fallback for when the network/SharePoint misbehaves.

---

## The 5-demo arc (problem → solution)

The talk is built around **five demos** that move from the problem to the working backend:

1. **The Problem** — a real tenant, one list, three users with different permissions. Show the wall each one hits.
2. **The Azure Function** — the deployed pieces in Azure, then the app-only **SharePoint** code in VS Code.
3. **The Domain endpoint** — Tier 1: origin validation, and why it's only half the story.
4. **Entra ID** — Tier 2: token-validation code **plus** the Entra app-registration setup.
5. **The SPFx web part** — the same three users from Demo 1, now *all* succeeding through the Function.

Demos **1 and 5 bookend** the talk: Demo 1 poses the problem (a no-access user can't touch the list), Demo 5 is the payoff (that same user reads/writes it through the Function). Demos 2–4 are the *how*.

⭐ = the four "aha" moments: **D2** (cert, not secret) · **D3** (the origin spoof) · **D4** (the wrong-audience 401) · **D5** (app-only elevation — perms don't matter).

> **Slide alignment:** this script was restructured from a 10-demo outline into these 5. Slide numbers are intentionally **not** referenced per-demo anymore — re-map the deck to these five sections before presenting. The still-true content gotchas (cert vs. secret, combined handler) are preserved in the appendix at the bottom.

---

## Pre-flight checklist (do this before you walk on stage)

- [ ] **Three signed-in browser profiles** ready for the `pdslabs2` tenant, each on the target site, labeled so the audience can tell them apart:
  - **Owner** — Full Control on the `Speaking Events` (or `Requests`) list.
  - **Reader** — Read-only on the list.
  - **Outsider** — **no** access to the list at all.
- [ ] `npm install` already run; `node_modules` present.
- [ ] `local.settings.json` populated from `local.settings.sample.json` with **real** `pdslabs2` values (cert PEM in `SHAREPOINT_CERT_PEM`; `ENTRA_AUDIENCE` set to the real App ID URI; `SHAREPOINT_SITE_URL`/`SHAREPOINT_LIST_TITLE` pinned if you're using locked-down mode).
- [ ] `func start` works locally → `http://localhost:7071/api/health` returns `{ "ok": true }`.
- [ ] Live app is warm: open `https://apidemo-func-4r1iq2.azurewebsites.net/api/health` once so the first cold start is paid before the audience is watching. (Or set `alwaysReadyInstances=1` and redeploy infra.)
- [ ] A **valid Entra bearer token** for the Tier-2 endpoint copied somewhere pasteable (for D4/jwt.ms), plus **one deliberately-wrong token** (a Graph token, or a token with the wrong `aud`).
- [ ] Postman/Thunder Client collection open with the four CRUD calls + a "no-origin" call + a spoofed-Origin call pre-built.
- [ ] `jwt.ms` open in a tab.
- [ ] Azure Portal open and **already signed in** to `pdslabs2`: the app registration, the Function App (`apidemo-func-4r1iq2`), and App Insights blades pinned.
- [ ] Companion **SPFx repo** (`data`) open in a second VS Code window for Demo 5 — `ServiceFactory.ts`, `ApiDomainSpService.ts`, `ApiEntraSpService.ts`, `PnPjsApiSpService.ts` ready.
- [ ] Font size cranked; terminal + editor at presentation zoom.
- [ ] **Secrets hygiene:** the cert PEM and any real tokens must NOT be visible on screen. Keep `local.settings.json` closed; show `local.settings.sample.json` if you must show a settings file.

**If the venue Wi-Fi dies:** the code reads in D2/D3/D4 and `npm test` all run **offline**. jwt.ms (D4) and anything live (the deployed app, the web part, the Portal) need the network — fall back to the recorded clips / screenshots noted in each demo.

---

## Time budget (70 min)

| Block | Section | Minutes | Demo | Priority |
|---|---|---:|---|---|
| Intro | Title / About / Roadmap | 4 | — | — |
| 1 | The Problem | 6 | **D1** | core (sets up D5) |
| 2 | The Azure Function: deployed + SP code | 16 | **D2** ⭐ | core |
| 3 | Tier 1 — the Domain endpoint | 9 | **D3** ⭐ | core |
| 4 | Tier 2 — Entra ID code + setup | 16 | **D4** ⭐ | core |
| 5 | The SPFx web part | 13 | **D5** ⭐ | core |
| Recap | Recap & Resources | 3 | — | — |

All five are core. If you fall behind, **shorten** D2's Azure-tour and D4's portal walk first (talk over screenshots), but never cut a demo whole — the arc breaks without all five.

---

# Demo 1 — The Problem: one list, three users ⭐ sets up the payoff

**~6 min · risk: MED (live SharePoint, multiple sign-ins)**

**Goal:** make the audience *feel* the problem before any code appears. SPFx runs as the **signed-in user**, so the user needs direct permission on the data — which means either over-permissioning people or telling them "no."

**Pre-state:** the `Speaking Events` list open. Three browser profiles ready: **Owner**, **Reader**, **Outsider**.

**Steps**
1. **Owner** opens the list → sees rows, can add/edit/delete. "Great — but this person already has Full Control. Most of your users shouldn't."
2. **Reader** opens the list → sees rows, but Add/Edit/Delete are unavailable. "Read-only. Can look, can't touch."
3. **Outsider** opens the list → **access denied**. "This person can't even see it. In a plain SPFx web part running in *their* context, there is nothing you can do for them without granting them permission you don't want to grant."
4. Pose the question that the rest of the talk answers: *"What if the action — not the user — should decide what's allowed? What if I want every one of these three to safely add an event, without making them all owners?"*

**Say:** "SPFx is the user's identity, full stop. That's perfect until the thing you need to do requires more trust than the user should have. The fix isn't more permissions — it's moving the privileged work to a backend that has its *own* identity. That backend is an Azure Function."

**Aha:** the Outsider's access-denied screen. Hold on it. This is the wall the whole talk knocks down — and D5 comes back to this exact account.

**Fallback:** a 60s screen capture of the three profiles, or three pre-taken screenshots side by side.

**Reset:** leave all three profiles signed in — you need them again in D5.

---

# Demo 2 — The Azure Function: deployed pieces, then the SharePoint code ⭐

**~16 min · risk: MED (live Azure + SharePoint)**

**Goal:** show what's running in Azure, then open the code and land the headline gotcha: **app-only to SharePoint REST must use a *certificate*, not a client secret.**

## Part A — the deployed pieces (Azure Portal, ~6 min)

**Pre-state:** Portal on the Function App `apidemo-func-4r1iq2`.

**Steps**
1. **Overview / Functions:** show the three discovered functions — `health`, `domain`, `entra`. "The runtime *discovers* these from `app.http(...)` calls — no routing config to maintain."
2. **Hit it live:** `https://apidemo-func-4r1iq2.azurewebsites.net/api/health` → `{ "ok": true }`.
3. **Plan:** point out **Flex Consumption (FC1)** — scale-to-zero, warm starts, Node 22. Mention `alwaysReadyInstances` is the cold-start knob ([main.bicep](../infra/main.bicep)).
4. **Configuration → App settings:** show the env vars the code reads (`AZURE_TENANT_ID`, `SHAREPOINT_APP_CLIENT_ID`, `SHAREPOINT_RESOURCE`, `ALLOWED_CALLER_DOMAINS`, `ENTRA_AUDIENCE`, …). Show a secret bound as `@Microsoft.KeyVault(SecretUri=…)` — "same env var name, zero code change."
5. **App Insights → Live Metrics:** fire a request, watch it appear.
6. **CORS:** show the allowed-origins list — SharePoint domains only, **no `*`**. "This is the *browser* preflight, separate from the app's own origin check we'll see in D3."
7. **How it got here (IaC):** flip to [main.bicep](../infra/main.bicep) + the split workflows — [infra.yml](../.github/workflows/infra.yml) (manual provision) vs [deploy.yml](../.github/workflows/deploy.yml) (push-triggered, build+test gated). Call out the **`remote-build` / `--build remote`** requirement: `.funcignore` excludes `node_modules` and **Flex Consumption doesn't support `SCM_DO_BUILD_DURING_DEPLOYMENT`**, so the build is requested server-side at deploy time. Skip it → zero functions registered → everything 404s. The #1 first-deploy trap.

## Part B — the SharePoint code (VS Code, ~10 min)

**Pre-state:** [src/sharepointClient.ts](../src/sharepointClient.ts) open; [helpers.ts](../src/functions/helpers.ts) in a split. Optional: a breakpoint on the `switch (method)` in `dispatchCrudOperation` ([helpers.ts:109](../src/functions/helpers.ts#L109)).

**Steps**
1. **`getSp()`** ([sharepointClient.ts:55](../src/sharepointClient.ts#L55)): `spfi(siteUrl).using(SPDefault(), AzureIdentity(getCredential(), [sharePointResource]))`. "This is the *exact* `spfi().using(...)` shape you'd write in a web part — the only difference in Node is *which* credential I hand it." Note the per-site SPFI cache.
2. **`getCredential()`** ([sharepointClient.ts:27](../src/sharepointClient.ts#L27)): if `SHAREPOINT_CERT_PEM` is set → **`ClientCertificateCredential`**; else fall back to `ClientSecretCredential`. The PEM (key + cert) comes straight from the app setting — no hand-rolled JWT; `@azure/identity` mints the token.
3. **The reveal ⭐:** "The secret branch *gets a token just fine*. But when that token hits SharePoint REST you get `401 Unsupported app only token`. SharePoint only honors **certificate-based** app-only tokens. That one line cost me an afternoon." (Optional live: flip to secret-only and show the 401.)
4. **The shared pipeline:** in [helpers.ts](../src/functions/helpers.ts), show `registerCrudRoutes(prefix, authorize)` ([helpers.ts:169](../src/functions/helpers.ts#L169)) → route `{prefix}/{itemId?}`, methods `GET/POST/PATCH/DELETE`, and `dispatchCrudOperation` ([helpers.ts:103](../src/functions/helpers.ts#L103)) switching the **HTTP method → CRUD action**. "One combined handler per tier. No action words in the URL. Both endpoints share this — auth is the only thing that differs." Note `runWithErrorHandling` (config preflight → 500, central error→status) and `resolveTarget` (locked-down `SHAREPOINT_SITE_URL`/`LIST_TITLE` vs. caller-supplied site/list).
5. The CRUD call site: `getSp(siteUrl).web.lists.getByTitle(listTitle).items.add(fields)` — "identical to a web part." Mention `update` is a MERGE with `IF-MATCH: *` (last-write-wins).

**Say:** "In the browser, PnPjs gets its identity from the page context. In a Function there's no user, so I give it an app identity explicitly — a daemon registration, `Sites.ReadWrite.All` application permission, admin-consented, authenticating with a **certificate**. Same API, explicit credential."

**Aha:** same PnPjs API as SPFx, just a different credential — and **cert > secret** for SharePoint app-only.

**Fallback:** Part A → previous deploy output + screenshots of the Portal blades. Part B → pure code read, no network. `npm test` (the [helpers.test.ts](../tests/helpers.test.ts) cases) proves the pipeline offline.

**Reset:** restore `local.settings.json` to the cert if you flipped it; clear the breakpoint.

---

# Demo 3 — Tier 1: the Domain endpoint ⭐

**~9 min · risk: LOW**

**Goal:** Tier 1 is a **bouncer checking where you came from**, not who you are — and it's spoofable, which is exactly *why* Tier 2 exists.

**Pre-state:** [src/security/domainRestriction.ts](../src/security/domainRestriction.ts) open (~35 lines — read it on screen) alongside [domainRestricted.ts](../src/functions/domainRestricted.ts) (the 4-line wiring). Postman with four pre-built `GET /api/domain` calls: (a) `Origin: https://pdslabs2.sharepoint.com` [allowed], (b) **no** Origin/Referer, (c) `Origin: https://evil.example.com`, (d) a `curl`-style call with a hand-set allowed Origin.

**Steps**
1. Show the wiring: [domainRestricted.ts](../src/functions/domainRestricted.ts) is just `registerCrudRoutes('domain', (req) => { ensureAllowedDomain(req); return { mode: 'domain-restricted' }; })`. "Same shared handler from D2; the authorizer is the only new thing."
2. Read `ensureAllowedDomain` ([domainRestriction.ts:23](../src/security/domainRestriction.ts#L23)): pulls `origin`, falls back to `referer`'s origin, 403s if not in `ALLOWED_CALLER_DOMAINS`.
3. Run (a) → `200` + data. Run (b) no origin → **403** "Request origin or referer is required." Run (c) wrong origin → **403** "Caller domain is not allowed."
4. **The honest catch ⭐:** run (d) — `curl`/Postman with a hand-added `Origin: https://pdslabs2.sharepoint.com` → **200**. "A browser *cannot* lie about Origin. `curl` can. So this stops casual cross-origin calls and random scanners — not a determined client. That's why Tier 1 is for low-risk reads only."
5. (Optional, offline) `npm test` → [domainRestriction.test.ts](../tests/domainRestriction.test.ts) codifies the allow / reject / referer-fallback cases.

**Say:** "Zero friction for a real SPFx web part — the browser sends Origin for free. Zero identity guarantee for anything else. Know which one you're buying."

**Aha:** the spoof. It motivates the entire Tier 2 half of the talk — land it deliberately.

**Fallback:** the `npm test` variant proves the same logic with no network.

---

# Demo 4 — Tier 2: Entra ID — token-validation code + Azure setup ⭐

**~16 min · risk: MED (portal + jwt.ms online)**

**Goal:** show the two-app registration dance, then make token validation concrete — four claims decide everything, and the code checks all four.

## Part A — the Azure setup (app registration, ~6 min)

**Pre-state:** Portal signed into `pdslabs2`, the app registration open.

**Steps**
1. **One app, two roles.** Call out up front that in this demo the *same* registration (`99b202da-…`) is both the **app-only-to-SharePoint** identity (its cert + `Sites.ReadWrite.All` from D2) **and** the **API the SPFx web part calls**. You could split them; here they're one.
2. **Certificates & secrets / API permissions:** show the uploaded certificate (the credential D2's code uses) and the admin-consented `Sites.ReadWrite.All` *application* permission.
3. **Expose an API:** show the **Application ID URI** `api://99b202da-…` and the custom scope (e.g. `SharePoint.ReadWrite`). "This URI is what becomes `ENTRA_AUDIENCE` ([config.ts:36](../src/config.ts#L36)) in the Function, and what SPFx's `getClient('api://…')` asks for in D5."
4. **Client consent:** the SPFx client's API permission → My APIs → this scope → **admin consent granted** (green check). "Until an admin approves this post-deploy, every Tier-2 call 401s. The step people forget."

## Part B — the validation code (VS Code + jwt.ms, ~10 min)

**Pre-state:** `jwt.ms` open. A **valid** bearer token and a **wrong** token (Graph token or wrong `aud`) ready to paste. [src/security/entraAuth.ts](../src/security/entraAuth.ts) open alongside [entraProtected.ts](../src/functions/entraProtected.ts).

**Steps**
1. Show the wiring: [entraProtected.ts](../src/functions/entraProtected.ts) is `registerCrudRoutes('entra', async (req) => { const p = await validateEntraToken(req); return { mode: 'entra-protected', caller: p.azp || p.appid || p.sub }; })`. "Same handler again — different authorizer."
2. Paste the valid token into jwt.ms. Highlight the four claims that matter:
   - **`iss`** — your tenant's v2 endpoint (`…/{tenant}/v2.0`).
   - **`aud`** — your API's App ID URI. *A Microsoft Graph token will not pass.*
   - **`azp` / `appid`** — your SPFx client app.
   - **`scp`** — the scope you required.
3. Map each to `validateEntraToken` ([entraAuth.ts:55](../src/security/entraAuth.ts#L55)): `jwt.verify` enforces `algorithms: ['RS256']`, `audience`, `issuer`; the signing key comes from **JWKS** via `jwks-rsa` (cached, [entraAuth.ts:15](../src/security/entraAuth.ts#L15)); then `ensureAllowedClientApplication` ([entraAuth.ts:44](../src/security/entraAuth.ts#L44)) enforces the optional `azp`/`appid` allow-list (`ALLOWED_CLIENT_APP_IDS`).
4. **The reject ⭐:** fire `GET /api/entra` with the wrong-`aud` token → **401 "Invalid token: …"**. "Right signature, real Entra token — wrong audience. Rejected. That one check kills a whole class of token-replay attacks." Then a **missing** token → 401 "Authorization bearer token is required" ([entraAuth.ts:58](../src/security/entraAuth.ts#L58)).

**Say:** "Validation isn't 'is there a token.' It's: signed by my tenant's keys, minted *for me*, by an app *I* allow, carrying the scope I need. Four claims, four checks, ~50 lines."

**Aha:** the wrong-audience 401. A token can be 100% valid and still not be *for you*.

**Fallback:** Part A → screenshots of the registration blades (portal nav on stage is slow). Part B → code read + the 401 against the local host need no internet; keep a decoded-token screenshot for the jwt.ms piece.

---

# Demo 5 — The SPFx web part: the payoff ⭐

**~13 min · risk: HIGH (live token + CORS + 3 sign-ins) — have a clip ready**

**Goal:** close the loop from Demo 1. The **same three users** now read and write the list through the Function — because the Function, not the user, holds the permission.

**Pre-state:** the dataDemo web part deployed to the `pdslabs2` site. The three browser profiles (**Owner**, **Reader**, **Outsider**) from D1 still signed in. The companion SPFx repo open in a second VS Code window.

**Steps**
1. **The code (brief):** show how the web part talks to the Function. `ServiceFactory.createSpService` branches by transport and endpoint:
   - **Simple Auth (Tier 1):** SPFx `httpClient` (`ApiDomainSpService`) — the browser attaches Origin for free.
   - **Entra App (Tier 2):** SPFx `aadHttpClientFactory.getClient('api://…')` (`ApiEntraSpService`) — **the `Authorization: Bearer` header is added automatically**; you never write token code.
   - **PnPjs variants:** the same two endpoints driven by PnPjs's `Queryable` pipeline (`PnPjsApiSpService`) instead of the native SPFx clients — same API, same result, different transport.
   Open the **Network tab**, fire an Entra App call, expand the request → the Bearer header you never wrote. Paste it into jwt.ms to tie back to D4.
2. **The payoff — run all three users** against the **Entra App** tab (and the **Simple Auth** tab):
   - **Owner** → reads + writes. Expected.
   - **Reader** → can still **add/edit/delete** through the web part. "They have read-only on the list — but the Function writes app-only, so the button works."
   - **Outsider** → the same. "Remember Demo 1? This account couldn't even *see* the list. Through the Function, they just added an event." Refresh the list as **Owner** to prove the row is really there.
3. Toggle the **SPFx ↔ PnPjs** transport for one of them to show the result is identical regardless of how the call is made.

**Say:** "This is the whole talk in one screen. The user's own permission on the list is now irrelevant to what the *action* is allowed to do — because the action runs with the Function's identity, gated by the tier we chose. Tier 1 trusts the origin; Tier 2 trusts a validated token. Either way, the elevation is centralized, auditable, and out of the browser."

**Aha:** the **Outsider writing a row**. That's Demo 1's denied user, now succeeding — the problem visibly solved.

**Fallback:** this is the riskiest live demo (token + CORS + multiple sign-ins). **Have a recorded clip** of all three accounts succeeding. Don't gamble the payoff on live token acquisition.

**Reset:** delete any stage rows the three users created so the next run is clean.

---

## Cut-down order if you're behind

Running long is the default failure mode. Trim **within** demos rather than dropping one — the arc needs all five. In order:
1. D2 Part A — talk over Portal screenshots instead of clicking through blades.
2. D4 Part A — one screenshot of the registration instead of a live walk.
3. D3 — skip the `npm test` coda; the three Postman calls + the spoof are enough.
4. D5 — show only **Outsider** (the strongest account) instead of all three.

**Never cut a whole demo.** D1 without D5 is a problem with no solution; D5 without D1 is a solution with no problem.

---

## Appendix — slide/code alignment notes (still true)

These gotchas survive the restructure; keep them in mind when you re-map the deck.

1. **App-Only Auth is a *certificate*, not a client secret.** The code authenticates with a **certificate** because SharePoint REST rejects secret-based app-only tokens (`401 Unsupported app only token`). Any slide that says "generate a client secret" should say **"generate & upload a certificate"** (`az ad app credential reset --id <clientId> --create-cert --append`), stored in `SHAREPOINT_CERT_PEM`. The shape in the repo:
   ```ts
   import { spfi } from "@pnp/sp";
   import { SPDefault } from "@pnp/nodejs";
   import { AzureIdentity } from "@pnp/azidjsclient";
   import { ClientCertificateCredential } from "@azure/identity";

   const credential = new ClientCertificateCredential(tenantId, clientId, { certificate: pem });
   export const getSp = (siteUrl: string) =>
     spfi(siteUrl).using(SPDefault(), AzureIdentity(credential, [sharePointResource]));
   ```
   This is the D2 reveal.

2. **One combined handler per tier — no action words in the URL.** The repo uses `/api/{tier}/{itemId?}` with `methods: ['GET','POST','PATCH','DELETE']`, and the **HTTP method maps to the CRUD action** inside `dispatchCrudOperation` ([helpers.ts:103](../src/functions/helpers.ts#L103)). Both tiers are wired by one shared `registerCrudRoutes(prefix, authorize)` ([helpers.ts:169](../src/functions/helpers.ts#L169)). If the deck shows a separate-function-per-verb diagram, drop it in favor of the combined-handler one.

3. **`@azure/identity` credential via the PnPjs `AzureIdentity` behavior** — if a slide mentions "AzureIdentityFetchClient or MSAL," simplify it to this, which is what the code uses.
