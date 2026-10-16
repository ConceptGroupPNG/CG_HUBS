/* =====================================================================
   CG HUBS - shared core library, used by every app in this repository.

   Load order in each app's page:
     1. MSAL Browser 3.x   <script src="https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.28.1/lib/msal-browser.min.js"></script>
     2. shared/cg-config.js
     3. shared/cg-core.js
     4. the app's own script

   Everything lives under ONE global name, `CG`, so this file can be added
   to an existing Hub without clashing with that Hub's own functions.

   THIS REPOSITORY IS PUBLIC - no secrets, staff data or customer data here.

   Brings together what was proven in production:
     - CG-Time (Oct 2026): redirect-only sign-in (the popup loop fix),
       inline Microsoft endpoints (so browser blockers can't silently stop
       sign-in), token keep-alive, one silent retry after a 401, and a
       "session expired" banner instead of raw errors.
     - CG Training HUB / PSP HUB: saveWithFieldFallback (skip columns the
       list doesn't have, and retry).
     - New: one shared file for every Hub, on / off / preview app switches,
       local-date helpers, and automatic wait-and-retry when SharePoint is
       busy (429 / 503).
   ===================================================================== */
(function (global) {
  "use strict";

  const LOGIN_HOST = "https://login.microsoftonline.com";
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const DEFAULT_SCOPES = ["Sites.ReadWrite.All", "User.Read"];
  const KEEPALIVE_MS = 10 * 60 * 1000;          // refresh the token every 10 minutes
  const REDIRECT_GUARD_MS = 2 * 60 * 1000;      // at most one automatic sign-in redirect per 2 minutes
  const MAX_BUSY_RETRIES = 3;                   // 429 / 503 retries

  const CG = {};

  /* -----------------------------------------------------------------
     Small helpers
     ----------------------------------------------------------------- */
  CG.esc = v => String(v ?? "").replace(/[&<>"']/g,
    c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  const pad = n => String(n).padStart(2, "0");
  // Dates are built from LOCAL calendar parts - never toISOString(),
  // which shifts the date by a day in PNG (UTC+10).
  CG.date = {
    ymd:  d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`,
    hhmm: d => `${pad(d.getHours())}:${pad(d.getMinutes())}`,
    today: () => CG.date.ymd(new Date()),
    addDays:   (s, n) => { const [y, m, d] = s.split("-").map(Number); return CG.date.ymd(new Date(y, m - 1, d + n)); },
    addMonths: (s, n) => { const [y, m, d] = s.split("-").map(Number); return CG.date.ymd(new Date(y, m - 1 + n, d)); },
    // "2026-10-16" -> "16/10/2026"
    display: s => { if (!s) return ""; const [y, m, d] = String(s).slice(0, 10).split("-"); return `${d}/${m}/${y}`; }
  };

  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const store = {
    get(k) { try { return global.sessionStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { global.sessionStorage.setItem(k, v); } catch (e) { /* storage blocked */ } },
    del(k) { try { global.sessionStorage.removeItem(k); } catch (e) { /* storage blocked */ } }
  };

  /* -----------------------------------------------------------------
     Apps and switches (from cg-config.js)
     ----------------------------------------------------------------- */
  const SHARED_URL = (global.document && document.currentScript)
    ? new URL(".", document.currentScript.src).href : "";
  CG.SHARED_URL = SHARED_URL;                                   // .../CG_HUBS/shared/
  CG.ROOT_URL = SHARED_URL ? new URL("..", SHARED_URL).href : ""; // .../CG_HUBS/

  CG.apps = () => (typeof CG_APPS !== "undefined" ? CG_APPS : []);
  CG.app = key => CG.apps().find(a => a.key === key) || null;
  CG.appUrl = key => { const a = CG.app(key); return a ? CG.ROOT_URL + a.folder : CG.ROOT_URL; };

  CG.appState = key => {
    const v = String((CG.app(key) || {}).status || "off").trim().toLowerCase();
    return v === "on" || v === "preview" ? v : "off";
  };
  // Can this email see / use the app?  (Hides an app - it is NOT security:
  // SharePoint permissions still decide who can read or change data.)
  CG.appAllowed = (key, email) => {
    const st = CG.appState(key);
    if (st === "on") return true;
    if (st === "off") return false;
    const me = String(email || "").trim().toLowerCase();
    const list = typeof CG_PREVIEW !== "undefined" ? CG_PREVIEW : [];
    return !!me && list.some(e => String(e).trim().toLowerCase() === me);
  };

  /* -----------------------------------------------------------------
     Sign-in (MSAL Browser 3.x, PKCE, full-page redirects only)
     ----------------------------------------------------------------- */
  let msalApp = null, scopes = DEFAULT_SCOPES, appKey = "", keepAliveTimer = null;
  let cacheLocation = "sessionStorage";
  const markerKey = () => `cg.authRedirectAt.${appKey}`;
  const markRedirect = () => store.set(markerKey(), String(Date.now()));
  const clearRedirect = () => store.del(markerKey());
  const redirectRecentlyTried = () => {
    const t = parseInt(store.get(markerKey()) || "0", 10);
    return !!t && (Date.now() - t) < REDIRECT_GUARD_MS;
  };

  // True when this page has only been loaded as Microsoft's return page inside
  // a popup or hidden iframe. The app must then do nothing - this is what
  // caused the Oct 2026 "sign-in window keeps opening" loop in CG Time.
  CG.isAuthReturnInChildWindow = () => {
    const frag = "&" + String(global.location.hash || "").replace(/^#/, "");
    if (!/&(code|error)=/.test(frag)) return false;
    let inFrame = false;
    try { inFrame = global.self !== global.top; } catch (e) { inFrame = true; }
    const inPopup = !!global.opener && global.opener !== global;
    return inFrame || inPopup;
  };

  // The return address is always the app's folder (…/psp/), even if the page
  // was opened as …/psp/index.html - it must match what is registered in Entra.
  CG.redirectUri = () => global.location.origin + global.location.pathname.replace(/index\.html$/i, "");

  function msalConfig(clientId, tenantId) {
    return {
      auth: {
        clientId,
        authority: `${LOGIN_HOST}/${tenantId}`,
        redirectUri: CG.redirectUri(),
        navigateToLoginRequestUrl: true,
        // Microsoft's fixed endpoints for this tenant, supplied inline so that
        // browser blockers can't stop sign-in by blocking the discovery lookups.
        knownAuthorities: ["login.microsoftonline.com"],
        cloudDiscoveryMetadata: JSON.stringify({
          tenant_discovery_endpoint: `${LOGIN_HOST}/${tenantId}/v2.0/.well-known/openid-configuration`,
          "api-version": "1.1",
          metadata: [{
            preferred_network: "login.microsoftonline.com",
            preferred_cache: "login.windows.net",
            aliases: ["login.microsoftonline.com", "login.windows.net", "login.microsoft.com", "sts.windows.net"]
          }]
        }),
        authorityMetadata: JSON.stringify({
          authorization_endpoint: `${LOGIN_HOST}/${tenantId}/oauth2/v2.0/authorize`,
          token_endpoint: `${LOGIN_HOST}/${tenantId}/oauth2/v2.0/token`,
          end_session_endpoint: `${LOGIN_HOST}/${tenantId}/oauth2/v2.0/logout`,
          issuer: `${LOGIN_HOST}/${tenantId}/v2.0`,
          jwks_uri: `${LOGIN_HOST}/${tenantId}/discovery/v2.0/keys`
        })
      },
      // sessionStorage (default): signed out when the browser closes.
      // Phone apps pass cacheLocation: "localStorage" to stay signed in.
      cache: { cacheLocation, storeAuthStateInCookie: false }
    };
  }

  // Turns a sign-in failure into something a person can act on.
  CG.describeAuthError = e => {
    const raw = e ? `${e.errorCode || ""} ${e.message || e}`.trim() : "unknown error";
    if (/endpoints_resolution_error|untrusted_authority|network_error|no_network_connectivity|post_request_failed|get_request_failed|Failed to fetch|ERR_BLOCKED/i.test(raw)) {
      return "Couldn't sign in — your browser blocked this page from reaching Microsoft sign-in " +
        "(login.microsoftonline.com). In Microsoft Edge this is usually Tracking prevention set to Strict: " +
        "Settings → Privacy, search, and services → Tracking prevention → choose Balanced, or add " +
        global.location.hostname + " under Exceptions. Otherwise an ad/tracker-blocking extension is " +
        "blocking it - switch it off for this site. Then refresh and sign in again. Technical detail: " + raw;
    }
    return "Couldn't sign in — " + raw;
  };

  CG.auth = {
    /* Call once when the page loads.
         await CG.auth.init({ app: "psp" })
       Options: scopes, clientId (if not in cg-config.js), and
       cacheLocation: "localStorage" for phone apps that should stay signed in.
       Returns the signed-in account, or null if the person needs to click
       Sign in (show your sign-in screen). Throws only with a message that
       can be shown on screen as it is (use CG.describeAuthError). */
    async init(opts = {}) {
      appKey = opts.app || "";
      const app = CG.app(appKey);
      const clientId = opts.clientId || (app && app.clientId);
      if (!clientId) throw new Error(`CG.auth.init: no client ID for app "${appKey}" in cg-config.js`);
      if (opts.scopes) scopes = opts.scopes;
      if (opts.cacheLocation === "localStorage") cacheLocation = "localStorage";
      if (CG.isAuthReturnInChildWindow()) return null;   // leave the response for the parent window
      if (!global.msal || !global.msal.PublicClientApplication) {
        throw new Error("Microsoft sign-in library (MSAL) didn't load - check the internet connection or a browser blocker, then refresh.");
      }
      msalApp = new global.msal.PublicClientApplication(msalConfig(clientId, CG_TENANT.TENANT_ID));
      await msalApp.initialize();

      let result;
      try { result = await msalApp.handleRedirectPromise(); }
      catch (e) { clearRedirect(); throw new Error(CG.describeAuthError(e)); }

      const accounts = msalApp.getAllAccounts();
      if (result && result.account) msalApp.setActiveAccount(result.account);
      else if (accounts.length) msalApp.setActiveAccount(msalApp.getActiveAccount() || accounts[0]);
      else return null;

      // Make sure a token can actually be obtained before the app starts.
      const t = await CG.auth.token({ interactive: true });
      if (!t) return null;            // navigating away to Microsoft sign-in
      clearRedirect();
      CG.auth.startKeepAlive();
      return msalApp.getActiveAccount();
    },

    account: () => (msalApp ? msalApp.getActiveAccount() : null),
    email: () => String((CG.auth.account() || {}).username || "").toLowerCase(),

    /* interactive:false (default) never leaves the page - it throws.
       interactive:true (page load only) falls back to ONE same-window
       redirect at most every 2 minutes, then shows the real error. */
    async token({ interactive = false, forceRefresh = false } = {}) {
      if (!msalApp) throw new Error("Not signed in yet.");
      const account = msalApp.getActiveAccount();
      const request = { scopes, account, forceRefresh };
      try {
        return (await msalApp.acquireTokenSilent(request)).accessToken;
      } catch (e) {
        if (!interactive) throw e;
        if (redirectRecentlyTried()) {
          clearRedirect();
          const who = account && account.username ? account.username : "this account";
          const code = e && (e.errorCode || e.name) ? (e.errorCode || e.name) : "unknown_error";
          const detail = e && (e.errorMessage || e.message) ? (e.errorMessage || e.message) : String(e);
          throw new Error(`Microsoft sign-in worked for ${who}, but this app still couldn't get permission to read SharePoint (${code}). ${detail} — please send a screenshot of this message to IT.`);
        }
        markRedirect();
        await msalApp.acquireTokenRedirect(request);
        return null;
      }
    },

    /* A quiet token for an OPTIONAL extra permission (e.g. reading staff
       photos). Never redirects; returns null if the permission isn't
       granted, so the feature can simply switch itself off. */
    async tokenFor(extraScopes) {
      if (!msalApp || !msalApp.getActiveAccount()) return null;
      try { return (await msalApp.acquireTokenSilent({ scopes: extraScopes, account: msalApp.getActiveAccount() })).accessToken; }
      catch (e) { return null; }
    },

    signIn() {
      clearRedirect(); markRedirect();
      return msalApp.loginRedirect({ scopes }).catch(e => { clearRedirect(); throw new Error(CG.describeAuthError(e)); });
    },
    signOut() {
      return msalApp.logoutRedirect({ account: msalApp.getActiveAccount() });
    },

    // Keeps the token fresh so a tab left open over lunch can still save.
    startKeepAlive() {
      if (keepAliveTimer) return;
      keepAliveTimer = setInterval(() => {
        CG.auth.token().catch(err => console.warn("Background token refresh failed - will retry on next request.", err));
      }, KEEPALIVE_MS);
    }
  };

  /* -----------------------------------------------------------------
     "Session expired" banner - shown instead of raw errors
     ----------------------------------------------------------------- */
  let bannerShown = false;
  // An app with its own "session expired" banner can set
  //   CG.onSessionExpired = { show: fn, hide: fn }
  // and the core will call those instead of showing its own.
  CG.onSessionExpired = null;
  CG.showSessionExpired = () => {
    if (bannerShown) return;
    bannerShown = true;
    if (CG.onSessionExpired && CG.onSessionExpired.show) { CG.onSessionExpired.show(); return; }
    if (!global.document) return;
    let b = document.getElementById("cg-session-banner");
    if (!b) {
      b = document.createElement("div");
      b.id = "cg-session-banner";
      b.setAttribute("role", "alert");
      b.style.cssText = "position:fixed;top:0;left:0;right:0;z-index:99999;display:flex;gap:12px;align-items:center;" +
        "justify-content:center;padding:10px 16px;background:#B03030;color:#fff;font:14px 'Segoe UI',Calibri,Arial,sans-serif";
      b.innerHTML = 'Your session has expired - please sign in again to continue. ' +
        '<button type="button" style="background:#fff;color:#B03030;border:0;border-radius:4px;padding:6px 14px;font-weight:600;cursor:pointer">Sign in</button>';
      b.querySelector("button").onclick = () => CG.auth.signIn().catch(e => alert(e.message));
      document.body.appendChild(b);
    }
    b.style.display = "flex";
  };
  const hideSessionExpired = () => {
    if (!bannerShown) return;
    bannerShown = false;
    if (CG.onSessionExpired && CG.onSessionExpired.hide) { CG.onSessionExpired.hide(); return; }
    const b = global.document && document.getElementById("cg-session-banner");
    if (b) b.style.display = "none";
  };

  /* -----------------------------------------------------------------
     Microsoft Graph requests
     ----------------------------------------------------------------- */
  // Readable message from a Graph error response. The error also carries
  // err.status (e.g. 404), err.code (Graph's code, e.g. "itemNotFound") and
  // err.body (the raw response text) for apps that need the detail.
  async function graphError(res) {
    let msg = `${res.status} ${res.statusText || ""}`.trim(), code = "", body = "";
    try { body = await res.text(); } catch (e) { /* no body */ }
    try { const j = JSON.parse(body); if (j && j.error) { if (j.error.message) msg = j.error.message; code = j.error.code || ""; } } catch (e) { /* not JSON */ }
    const err = new Error(msg);
    err.status = res.status;
    err.code = code;
    err.body = body;
    return err;
  }

  /* CG.graph(method, url, body?, headers?)
     url may be a full URL or a path like "/sites/…".
     - 401: one silent token refresh and replay; if that fails, the session
       banner is shown and a plain "session expired" error is thrown.
     - 429 / 503: waits (Retry-After, else 2s, 4s, 8s) and retries.
     - Other errors throw with Graph's own message and err.status set. */
  CG.graph = async function (method, url, body, headers) {
    const full = /^https?:/i.test(url) ? url : GRAPH + url;
    // Objects are sent as JSON; a string is taken to be JSON already and sent
    // as it is; a Blob/File is sent as raw content.
    const isBlob = !!(global.Blob && body instanceof global.Blob);
    const isJson = body !== undefined && body !== null && !isBlob;
    // If the sign-in can't be renewed quietly, show the banner and a plain
    // message - never Microsoft's technical error text.
    const tokenOrExpired = async forceRefresh => {
      try { return await CG.auth.token({ forceRefresh }); }
      catch (e) {
        console.warn("Silent token renewal failed.", e);
        CG.showSessionExpired();
        const err = new Error("Your session has expired. Please sign in again to continue.");
        err.status = 401;
        throw err;
      }
    };
    const send = async (forceRefresh = false) => fetch(full, {
      method,
      headers: Object.assign({ Authorization: "Bearer " + await tokenOrExpired(forceRefresh) },
        isJson ? { "Content-Type": "application/json" } : {}, headers || {}),
      body: body === undefined || body === null ? undefined : (isJson && typeof body !== "string" ? JSON.stringify(body) : body)
    });

    let res, busy = 0;
    for (;;) {
      res = await send();
      if ((res.status === 429 || res.status === 503) && busy < MAX_BUSY_RETRIES) {
        const after = parseInt(res.headers && res.headers.get ? res.headers.get("Retry-After") || "" : "", 10);
        await sleep((after > 0 ? Math.min(after, 30) : 2 ** (busy + 1)) * 1000);
        busy++;
        continue;
      }
      break;
    }

    if (res.status === 401) {
      try { res = await send(true); } catch (e) { console.warn("Token refresh after 401 failed.", e); }
      if (res.status === 401) {
        CG.showSessionExpired();
        const err = new Error("Your session has expired. Please sign in again to continue.");
        err.status = 401;
        throw err;
      }
    }
    if (!res.ok) throw await graphError(res);
    hideSessionExpired();
    if (res.status === 204) return null;
    return res.json().catch(() => null);
  };

  /* -----------------------------------------------------------------
     SharePoint lists
     ----------------------------------------------------------------- */
  const siteIds = {};
  // "/sites/PSP_Test" -> Graph site id (looked up once, then remembered)
  async function siteId(sitePath) {
    if (!siteIds[sitePath]) siteIds[sitePath] = (await CG.graph("GET", `/sites/${CG_TENANT.SP_HOST}:${sitePath}`)).id;
    return siteIds[sitePath];
  }
  const listUrl = async (sitePath, list) => `/sites/${await siteId(sitePath)}/lists/${encodeURIComponent(list)}/items`;

  CG.sp = {
    siteId,

    // Every item of a list, page by page. Returns [{ id, ...fields }].
    // query: extra OData, e.g. "&$filter=fields/IsActive eq 1"
    async items(sitePath, list, query = "") {
      let url = `${await listUrl(sitePath, list)}?expand=fields&$top=999${query}`;
      const out = [];
      while (url) {
        const page = await CG.graph("GET", url);
        (page.value || []).forEach(i => out.push(Object.assign({ id: i.id }, i.fields)));
        url = page["@odata.nextLink"] || null;
      }
      return out;
    },

    async create(sitePath, list, fields) {
      const r = await CG.graph("POST", await listUrl(sitePath, list), { fields });
      return Object.assign({ id: r.id }, r.fields);
    },
    async update(sitePath, list, id, fields) {
      return CG.graph("PATCH", `${await listUrl(sitePath, list)}/${id}/fields`, fields);
    },
    async remove(sitePath, list, id) {
      return CG.graph("DELETE", `${await listUrl(sitePath, list)}/${id}`);
    },

    /* Create (id empty) or update (id given). Any column SharePoint says it
       doesn't recognise is dropped and the save retried - the proven
       saveWithFieldFallback behaviour. Returns { id, ...fields, skipped: [] }. */
    async save(sitePath, list, id, fields) {
      const f = Object.assign({}, fields), skipped = [];
      for (let attempt = 0; attempt < 10; attempt++) {
        try {
          const r = id ? await CG.sp.update(sitePath, list, id, f) : await CG.sp.create(sitePath, list, f);
          if (skipped.length) console.warn(`${list}: columns not on the list, skipped:`, skipped);
          return Object.assign({ id: id || r.id }, r, { skipped });
        } catch (e) {
          const m = /Field '([^']+)' is not recognized/i.exec(e.message || "");
          if (m && Object.prototype.hasOwnProperty.call(f, m[1])) { skipped.push(m[1]); delete f[m[1]]; continue; }
          throw e;
        }
      }
      throw new Error(`${list}: too many unrecognised columns - save abandoned.`);
    },

    isMissingList: e => e && (e.status === 404 || /not (be )?found|does not exist/i.test(e.message || "")),
    isForbidden:   e => e && (e.status === 403 || /access denied|forbidden|not authorized/i.test(e.message || ""))
  };

  CG.VERSION = "1.3.0";
  global.CG = CG;
})(typeof window !== "undefined" ? window : globalThis);
