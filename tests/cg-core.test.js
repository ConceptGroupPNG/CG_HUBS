// Tests for shared/cg-core.js - run with:  node --test tests/
// Microsoft sign-in (MSAL) and Microsoft Graph are simulated; nothing real is contacted.
const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const CONFIG_SRC = fs.readFileSync(path.join(__dirname, "../shared/cg-config.js"), "utf8");
const CORE_SRC = fs.readFileSync(path.join(__dirname, "../shared/cg-core.js"), "utf8");

// Builds a fresh simulated browser each time, loads cg-config.js + cg-core.js into it.
function load({ hash = "", inFrame = false, opener = null, silent, fetchQueue = [], pathname = "/CG_HUBS/psp/index.html" } = {}) {
  const session = new Map();
  const calls = { redirects: 0, fetches: [], silent: 0, forceRefresh: 0 };
  const accounts = [{ username: "Test.User@concept.com.pg" }];
  let active = null;

  class PCA {
    constructor(cfg) { calls.config = cfg; }
    async initialize() {}
    async handleRedirectPromise() { return null; }
    getAllAccounts() { return accounts; }
    getActiveAccount() { return active; }
    setActiveAccount(a) { active = a; }
    async acquireTokenSilent(req) {
      calls.silent++; if (req.forceRefresh) calls.forceRefresh++;
      return silent ? silent(req, calls) : { accessToken: "tok" + calls.silent };
    }
    async acquireTokenRedirect() { calls.redirects++; }
    async loginRedirect() { calls.redirects++; }
    async logoutRedirect() {}
  }

  const win = {
    location: { hash, origin: "https://conceptgrouppng.github.io", pathname, hostname: "conceptgrouppng.github.io" },
    opener,
    sessionStorage: { getItem: k => (session.has(k) ? session.get(k) : null), setItem: (k, v) => session.set(k, String(v)), removeItem: k => session.delete(k) },
    msal: { PublicClientApplication: PCA },
    console: { warn() {}, log() {}, error() {} },
    setInterval: () => 1, setTimeout: (fn) => { fn(); return 1; },
    URL, JSON, Promise, Date, Math, Object, String, Error, RegExp, parseInt,
    fetch: async (url, opts) => {
      calls.fetches.push({ url, opts });
      const next = fetchQueue.shift();
      if (!next) throw new Error("unexpected fetch " + url);
      return typeof next === "function" ? next(url, opts) : next;
    }
  };
  win.self = win; win.top = inFrame ? {} : win; win.window = win; win.globalThis = win;
  const ctx = vm.createContext(win);
  vm.runInContext(CONFIG_SRC + "\nthis.CG_APPS = CG_APPS; this.CG_TENANT = CG_TENANT; this.CG_PREVIEW = CG_PREVIEW;", ctx);
  vm.runInContext(CORE_SRC, ctx);
  return { CG: win.CG, win, calls, session };
}

const res = (status, body, headers = {}) => ({
  status, ok: status >= 200 && status < 300, statusText: "",
  headers: { get: k => headers[k] ?? null },
  json: async () => body,
  text: async () => (body === null || body === undefined ? "" : JSON.stringify(body))
});

/* ---------------- dates ---------------- */
test("dates use local calendar parts, never UTC", () => {
  const { CG } = load();
  assert.equal(CG.date.ymd(new Date(2026, 9, 16, 23, 30)), "2026-10-16");
  assert.equal(CG.date.addDays("2026-12-31", 1), "2027-01-01");
  assert.equal(CG.date.addMonths("2026-01-31", 1), "2026-03-03"); // JavaScript rolls 31 Feb over into March
  assert.equal(CG.date.display("2026-10-16"), "16/10/2026");
  assert.equal(CG.date.display(""), "");
});

test("esc escapes HTML", () => {
  const { CG } = load();
  assert.equal(CG.esc(`<a href="x">'&'</a>`), "&lt;a href=&quot;x&quot;&gt;&#39;&amp;&#39;&lt;/a&gt;");
  assert.equal(CG.esc(null), "");
});

/* ---------------- app switches ---------------- */
test("on / off / preview switches", () => {
  const { CG } = load();
  CG.app("psp").status = "on"; CG.app("csa").status = "preview"; CG.app("sla").status = "off";
  assert.equal(CG.appAllowed("psp", "anyone@concept.com.pg"), true);
  assert.equal(CG.appAllowed("csa", "MSCULLY@concept.com.pg"), true);   // case-insensitive
  assert.equal(CG.appAllowed("csa", "someone.else@concept.com.pg"), false);
  assert.equal(CG.appAllowed("sla", "mscully@concept.com.pg"), false);
  assert.equal(CG.appState("nonexistent"), "off");
  assert.equal(CG.apps().length, 7);
});

/* ---------------- sign-in ---------------- */
test("redirect URI is always the folder, never index.html", async () => {
  const { CG, calls } = load();
  await CG.auth.init({ app: "psp" });
  assert.equal(calls.config.auth.redirectUri, "https://conceptgrouppng.github.io/CG_HUBS/psp/");
  assert.equal(calls.config.auth.clientId, CG.app("psp").clientId);   // taken from cg-config.js
  assert.ok(calls.config.auth.authorityMetadata.includes("ea9e105c-2df6-47ef-a8e3-b4aa1d788423"));
});

test("does nothing when loaded as the sign-in return page inside a popup or iframe (loop fix)", async () => {
  for (const env of [{ inFrame: true }, { opener: {} }]) {
    const { CG, calls } = load({ hash: "#code=abc&state=x", ...env });
    assert.equal(await CG.auth.init({ app: "psp" }), null);
    assert.equal(calls.config, undefined, "MSAL must not even be created");
  }
});

test("a normal page with a sign-in response is NOT treated as a child window", () => {
  const { CG } = load({ hash: "#code=abc" });
  assert.equal(CG.isAuthReturnInChildWindow(), false);
});

test("signed-in account is returned when a token can be obtained", async () => {
  const { CG } = load();
  const acct = await CG.auth.init({ app: "psp" });
  assert.equal(acct.username, "Test.User@concept.com.pg");
  assert.equal(CG.auth.email(), "test.user@concept.com.pg");
});

test("silent token failure on page load redirects ONCE, then shows a real error instead of looping", async () => {
  const fail = () => { const e = new Error("consent needed"); e.errorCode = "consent_required"; throw e; };
  const first = load({ silent: fail });
  assert.equal(await first.CG.auth.init({ app: "psp" }), null);
  assert.equal(first.calls.redirects, 1);
  // Same browser tab comes back within 2 minutes and it still fails:
  const second = load({ silent: fail });
  for (const [k, v] of first.session) second.session.set(k, v);
  await assert.rejects(second.CG.auth.init({ app: "psp" }), /couldn't get permission.*consent_required/);
  assert.equal(second.calls.redirects, 0);
});

test("blocked Microsoft sign-in gives the plain Edge/blocker explanation", () => {
  const { CG } = load();
  assert.match(CG.describeAuthError({ errorCode: "endpoints_resolution_error", message: "x" }), /Tracking prevention/);
  assert.match(CG.describeAuthError(new Error("user_cancelled")), /^Couldn't sign in — user_cancelled/);
});

test("init refuses an app with no client ID", async () => {
  const { CG } = load();
  await assert.rejects(CG.auth.init({ app: "nope" }), /no client ID/);
});

/* ---------------- Graph ---------------- */
async function signedIn(opts) { const t = load(opts); await t.CG.auth.init({ app: "psp" }); return t; }

test("401 -> one forced token refresh and replay", async () => {
  const { CG, calls } = await signedIn({ fetchQueue: [res(401, {}), res(200, { ok: 1 })] });
  assert.deepEqual(await CG.graph("GET", "/me"), { ok: 1 });
  assert.equal(calls.forceRefresh, 1);
  assert.equal(calls.fetches[0].url, "https://graph.microsoft.com/v1.0/me");
});

test("401 twice -> plain 'session expired' error, not raw Graph JSON", async () => {
  const { CG } = await signedIn({ fetchQueue: [res(401, {}), res(401, {})] });
  await assert.rejects(CG.graph("GET", "/me"), e => e.status === 401 && /session has expired/.test(e.message));
});

test("429 / 503 busy -> waits and retries", async () => {
  const { CG, calls } = await signedIn({ fetchQueue: [res(429, {}, { "Retry-After": "1" }), res(503, {}), res(200, { v: 1 })] });
  assert.deepEqual(await CG.graph("GET", "/x"), { v: 1 });
  assert.equal(calls.fetches.length, 3);
});

test("errors carry Graph's own message and status", async () => {
  const { CG } = await signedIn({ fetchQueue: [res(404, { error: { message: "List 'X' does not exist" } })] });
  await assert.rejects(CG.graph("GET", "/x"), e => e.status === 404 && CG.sp.isMissingList(e));
});

test("JSON body is sent with Content-Type; 204 returns null", async () => {
  const { CG, calls } = await signedIn({ fetchQueue: [res(204, null)] });
  assert.equal(await CG.graph("PATCH", "/x", { A: 1 }), null);
  assert.equal(calls.fetches[0].opts.headers["Content-Type"], "application/json");
  assert.equal(calls.fetches[0].opts.body, '{"A":1}');
});

/* ---------------- SharePoint ---------------- */
const SITE = res(200, { id: "site-1" });

test("items() follows every page and flattens fields", async () => {
  const { CG, calls } = await signedIn({ fetchQueue: [SITE,
    res(200, { value: [{ id: "1", fields: { Title: "A" } }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next" }),
    res(200, { value: [{ id: "2", fields: { Title: "B" } }] })] });
  const items = await CG.sp.items("/sites/PSP_Test", "PSP_Billing");
  assert.deepEqual(JSON.parse(JSON.stringify(items.map(i => [i.id, i.Title]))), [["1", "A"], ["2", "B"]]);
  assert.match(calls.fetches[0].url, /sites\/conceptcompg\.sharepoint\.com:\/sites\/PSP_Test$/);
  assert.match(calls.fetches[1].url, /sites\/site-1\/lists\/PSP_Billing\/items\?expand=fields&\$top=999$/);
});

test("site id is looked up once and remembered", async () => {
  const { CG, calls } = await signedIn({ fetchQueue: [SITE, res(200, { value: [] }), res(200, { value: [] })] });
  await CG.sp.items("/sites/PSP_Test", "A"); await CG.sp.items("/sites/PSP_Test", "B");
  assert.equal(calls.fetches.length, 3);
});

test("save() skips unrecognised columns and retries (field fallback)", async () => {
  const bad = f => res(400, { error: { message: `Field '${f}' is not recognized` } });
  const { CG, calls } = await signedIn({ fetchQueue: [SITE, bad("Old1"), bad("Old2"), res(201, { id: "9", fields: { Title: "T" } })] });
  const r = await CG.sp.save("/sites/PSP_Test", "L", null, { Title: "T", Old1: 1, Old2: 2 });
  assert.equal(r.id, "9");
  assert.deepEqual([...r.skipped], ["Old1", "Old2"]);
  assert.deepEqual(JSON.parse(calls.fetches[3].opts.body), { fields: { Title: "T" } });
});

test("save() with an id updates; other errors are not swallowed", async () => {
  const { CG, calls } = await signedIn({ fetchQueue: [SITE, res(200, { Title: "U" }), res(403, { error: { message: "Access denied" } })] });
  const r = await CG.sp.save("/sites/PSP_Test", "L", "5", { Title: "U" });
  assert.equal(r.id, "5");
  assert.equal(calls.fetches[1].opts.method, "PATCH");
  await assert.rejects(CG.sp.save("/sites/PSP_Test", "L", "5", { Title: "U" }), e => CG.sp.isForbidden(e));
});

/* ---------------- added in 1.1.0 ---------------- */
test("sign-in is forgotten when the browser closes, unless the app asks to stay signed in", async () => {
  const a = load(); await a.CG.auth.init({ app: "time" });
  assert.equal(a.calls.config.cache.cacheLocation, "sessionStorage");
  const b = load(); await b.CG.auth.init({ app: "mobile", cacheLocation: "localStorage" });
  assert.equal(b.calls.config.cache.cacheLocation, "localStorage");
  assert.equal(b.calls.config.auth.clientId, "ec5454e0-c794-419b-9ec9-464e9917ef3c");
});

test("if the token can't be renewed quietly, Graph calls give the plain 'session expired' message", async () => {
  let n = 0;
  const { CG, calls } = load({ silent: () => { if (++n > 1) { const e = new Error("AADSTS700082: refresh token expired"); e.errorCode = "interaction_required"; throw e; } return { accessToken: "t" }; } });
  await CG.auth.init({ app: "psp" });
  await assert.rejects(CG.graph("GET", "/me"), e => e.status === 401 && /session has expired/.test(e.message) && !/AADSTS/.test(e.message));
  assert.equal(calls.fetches.length, 0);
});

/* ---------------- added in 1.2.0 ---------------- */
test("errors also carry Graph's code and the raw response body", async () => {
  const { CG } = await signedIn({ fetchQueue: [res(403, { error: { code: "ErrorAccessDenied", message: "Access is denied." } })] });
  await assert.rejects(CG.graph("POST", "/users/x/sendMail", {}), e =>
    e.status === 403 && e.code === "ErrorAccessDenied" && e.message === "Access is denied." && /"code":"ErrorAccessDenied"/.test(e.body));
});

test("a string body is sent as it is (not JSON-encoded twice)", async () => {
  const { CG, calls } = await signedIn({ fetchQueue: [res(201, { id: "1" })] });
  await CG.graph("POST", "/x", JSON.stringify({ a: 1 }));
  assert.equal(calls.fetches[0].opts.body, '{"a":1}');
  assert.equal(calls.fetches[0].opts.headers["Content-Type"], "application/json");
});

test("tokenFor() quietly returns a token, or null when the extra permission isn't granted", async () => {
  const ok = await signedIn();
  assert.match(await ok.CG.auth.tokenFor(["ProfilePhoto.Read.All"]), /^tok/);
  const no = await signedIn({ silent: (req) => { if (req.scopes[0] === "ProfilePhoto.Read.All") throw new Error("consent_required"); return { accessToken: "t" }; } });
  assert.equal(await no.CG.auth.tokenFor(["ProfilePhoto.Read.All"]), null);
  assert.equal(no.calls.redirects, 0);
});

/* ---------------- added in 1.3.0 ---------------- */
test("an app can use its own session-expired banner instead of the core's", async () => {
  const { CG } = await signedIn({ fetchQueue: [res(401, {}), res(401, {}), res(200, { ok: 1 })] });
  const seen = [];
  CG.onSessionExpired = { show: () => seen.push("show"), hide: () => seen.push("hide") };
  await assert.rejects(CG.graph("GET", "/me"), e => e.status === 401);
  await CG.graph("GET", "/me");               // a later success clears it
  assert.deepEqual([...seen], ["show", "hide"]);
});
