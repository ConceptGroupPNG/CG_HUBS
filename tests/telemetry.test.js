// Tests for CG.telemetry in shared/cg-core.js - run with:  node --test tests/
// A simulated browser and SharePoint; nothing real is contacted.
const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const CONFIG_SRC = fs.readFileSync(path.join(__dirname, "../shared/cg-config.js"), "utf8");
const CORE_SRC = fs.readFileSync(path.join(__dirname, "../shared/cg-core.js"), "utf8");

const res = (status, body) => ({
  status, ok: status >= 200 && status < 300, statusText: "",
  headers: { get: () => null },
  json: async () => body, text: async () => JSON.stringify(body)
});

// route(method, url, body) -> response; default: everything succeeds
function load({ route, scopes, telemetryConfig, withDocument = true, protocol } = {}) {
  const calls = [];
  const listeners = {};
  const storage = new Map();
  let active = null;
  class PCA {
    async initialize() {}
    async handleRedirectPromise() { return null; }
    getAllAccounts() { return [{ username: "Carlos.Atalou@concept.com.pg", name: "Carlos Atalou" }]; }
    getActiveAccount() { return active; }
    setActiveAccount(a) { active = a; }
    async acquireTokenSilent() { return { accessToken: "tok" }; }
  }
  const doc = {
    visibilityState: "visible", lastModified: "11/04/2026 15:20:11", body: {},
    addEventListener: (t, fn) => { (listeners["doc:" + t] = listeners["doc:" + t] || []).push(fn); },
    getElementById: () => null
  };
  const win = {
    location: { protocol, reload() { win.reloaded = true; }, hash: "#invoices", origin: "https://conceptgrouppng.github.io", pathname: "/CG_HUBS/psp/", hostname: "conceptgrouppng.github.io", href: "https://conceptgrouppng.github.io/CG_HUBS/psp/" },
    sessionStorage: { getItem: () => null, setItem() {}, removeItem() {} },
    localStorage: { getItem: k => storage.get(k) ?? null, setItem: (k, v) => storage.set(k, String(v)) },
    navigator: { userAgent: "Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 Chrome/131.0 Safari/537.36 Edg/131.0" },
    performance: { getEntriesByType: () => [{ loadEventEnd: 1834.6 }] },
    screen: { width: 1920, height: 1080 },
    matchMedia: () => ({ matches: false }),
    msal: { PublicClientApplication: PCA },
    console: { warn() {}, log() {}, error() {} },
    setInterval: () => 7, clearInterval() {}, setTimeout: fn => { fn(); return 1; },
    addEventListener: (t, fn) => { (listeners[t] = listeners[t] || []).push(fn); },
    URL, JSON, Promise, Date, Math, Object, String, Error, RegExp, parseInt, Number, Array,
    fetch: async (url, opts) => {
      const body = opts.body ? JSON.parse(opts.body) : undefined;
      calls.push({ method: opts.method, url: decodeURIComponent(url), body, headers: opts.headers });
      const r = route && route(opts.method, decodeURIComponent(url), body);
      if (r) return r;
      if (/\/sites\/[^/]+:\/sites\//.test(url)) return res(200, { id: "site-" + url.split(":").pop() });
      if (opts.method === "POST") return res(201, { id: "41", fields: {} });
      if (opts.method === "GET") return res(200, { value: [] });
      return res(200, {});
    }
  };
  if (withDocument) win.document = doc;
  win.self = win; win.top = win; win.window = win; win.globalThis = win;
  const ctx = vm.createContext(win);
  const cfg = telemetryConfig === undefined ? "" : `\nCG_TELEMETRY.enabled = ${telemetryConfig};`;
  vm.runInContext(CONFIG_SRC + cfg, ctx);
  vm.runInContext(CORE_SRC, ctx);
  const fire = (t, ev) => (listeners[t] || []).forEach(fn => fn(ev));
  return { CG: win.CG, calls, fire, doc, storage, start: () => win.CG.auth.init({ app: "psp", scopes }) };
}

const posts = (calls, list) => calls.filter(c => c.method === "POST" && c.url.includes(`/lists/${list}/items`));

test("a visit creates one session row with who, app, build, load time and device", async () => {
  const { CG, calls, start } = load();
  await start();
  await CG.telemetry._flush();
  const p = posts(calls, "CG_HubSessions");
  assert.equal(p.length, 1);
  const f = p[0].body.fields;
  assert.equal(f.App, "psp");
  assert.equal(f.UserName, "Carlos Atalou");
  assert.equal(f.UserEmail, "carlos.atalou@concept.com.pg");
  assert.equal(f.Screens, "invoices");
  assert.equal(f.LoadMs, 1835);
  assert.equal(f.Device, "Desktop");
  assert.equal(f.Browser, "Edge 131");
  assert.match(f.Build, /^core \d+\.\d+\.\d+ · page 2026-11-04 15:20$/);
  assert.ok(calls.some(c => c.url.includes("conceptcompg.sharepoint.com:/sites/CG_Platform")), "lists are on CG_Platform");
});

test("heartbeat updates the same row; nothing is sent while the page is hidden", async () => {
  const { CG, calls, start, doc, fire } = load();
  await start();
  await CG.telemetry._flush();
  CG.telemetry.screen("Billing");
  const realNow = Date.now;
  try {
    Date.now = () => realNow() + 10 * 60000;          // ten minutes later
    doc.visibilityState = "hidden";
    fire("doc:visibilitychange");
    await CG.telemetry._flush();
    assert.equal(calls.filter(c => c.method === "PATCH").length, 0);
    doc.visibilityState = "visible";                  // back on screen
    fire("doc:visibilitychange");
    await CG.telemetry._flush();
  } finally { Date.now = realNow; }
  const p = calls.filter(c => c.method === "PATCH");
  assert.equal(p.length, 1);
  assert.match(p[0].url, /lists\/CG_HubSessions\/items\/41\/fields$/);
  assert.equal(p[0].body.PageViews, 2);
  assert.equal(p[0].body.Screens, "invoices, Billing");
  assert.ok(p[0].body.LastSeen);
});

test("errors: recorded once per minute per message, capped per visit, counted on the session", async () => {
  const { CG, calls, start } = load();
  await start();
  await CG.telemetry._flush();
  CG.telemetry.error("save", "Couldn't save invoice", "PSP_Invoices");
  CG.telemetry.error("save", "Couldn't save invoice", "PSP_Invoices");   // duplicate within a minute
  for (let i = 0; i < 30; i++) CG.telemetry.error("script", "boom " + i, "");
  await CG.telemetry._flush();
  const e = posts(calls, "CG_HubErrors");
  assert.equal(e.length, 20);                      // maxErrorsPerSession
  assert.equal(e[0].body.fields.Title, "save");
  assert.equal(e[0].body.fields.App, "psp");
  assert.equal(e[0].body.fields.Page, "/CG_HUBS/psp/#invoices");
  assert.equal(e[0].body.fields.SessionId, CG.telemetry.sessionId());
  assert.equal(CG.telemetry._state().errors, 32);  // all counted, even those not sent
});

test("script errors and failed promises in the page are caught; other sites' scripts are ignored", async () => {
  const { CG, calls, start, fire } = load();
  await start();
  await CG.telemetry._flush();
  fire("error", { message: "Script error.", filename: "" });
  fire("error", { message: "x is not defined", filename: "https://conceptgrouppng.github.io/CG_HUBS/psp/", lineno: 120, error: { stack: "at f" } });
  fire("unhandledrejection", { reason: { status: 401, message: "expired" } });
  fire("unhandledrejection", { reason: { message: "Save failed" } });
  await CG.telemetry._flush();
  const e = posts(calls, "CG_HubErrors").map(c => c.body.fields.Message);
  assert.deepEqual(e, ["x is not defined", "Save failed"]);
});

test("SharePoint throttling and server errors in the app are reported; expected 404s are not", async () => {
  let n = 0;
  const { CG, calls, start } = load({
    route: (m, url) => {
      if (url.includes("/lists/PSP_Billing/")) { n++; return n <= 4 ? res(429, { error: { message: "busy" } }) : null; }
      if (url.includes("/lists/PSP_Missing/")) return res(404, { error: { code: "itemNotFound", message: "not found" } });
      if (url.includes("/lists/PSP_Broken/")) return res(500, { error: { message: "Server fell over" } });
      return null;
    }
  });
  await start();
  await CG.telemetry._flush();
  await assert.rejects(CG.sp.items("/sites/PSP_Test", "PSP_Billing"));   // 429 x4 -> gives up
  await assert.rejects(CG.sp.items("/sites/PSP_Test", "PSP_Missing"));
  await assert.rejects(CG.sp.items("/sites/PSP_Test", "PSP_Broken"));
  await CG.telemetry._flush();
  assert.equal(CG.telemetry._state().throttles, 3);
  const e = posts(calls, "CG_HubErrors").map(c => c.body.fields.Message);
  assert.deepEqual(e, ["429 busy", "500 Server fell over"]);
});

test("switches itself off when the lists don't exist yet - the app carries on", async () => {
  const { CG, calls, start } = load({ route: (m, url) => url.includes("CG_HubSessions") ? res(404, { error: { message: "List not found" } }) : null });
  const account = await start();
  assert.ok(account, "sign-in still succeeds");
  await CG.telemetry._flush();
  assert.equal(CG.telemetry.enabled(), false);
  const before = calls.length;
  CG.telemetry.error("save", "after switch-off", "");
  await CG.telemetry._flush();
  assert.equal(calls.length, before);
});

test("read-only apps (Mission Control, launcher) and telemetry off: nothing is written", async () => {
  const a = load({ scopes: ["Sites.Read.All", "User.Read"] });
  await a.start(); await a.CG.telemetry._flush();
  assert.equal(a.CG.telemetry.enabled(), false);
  assert.equal(posts(a.calls, "CG_HubSessions").length, 0);

  const b = load({ telemetryConfig: "false" });
  await b.start(); await b.CG.telemetry._flush();
  assert.equal(b.CG.telemetry.enabled(), false);
});

test("old rows are cleaned up at most once a day, using the indexed columns", async () => {
  const { CG, calls, start, storage } = load({
    route: (m, url) => (m === "GET" && url.includes("CG_HubSessions") && url.includes("lt '")) ? res(200, { value: [{ id: "7" }, { id: "8" }] }) : null
  });
  await start();
  await CG.telemetry._flush();
  const gets = calls.filter(c => c.method === "GET" && /lt '/.test(c.url));
  assert.equal(gets.length, 2);
  assert.ok(gets[0].url.includes("fields/LastSeen lt '"));
  assert.ok(gets[1].url.includes("fields/Created lt '"));
  assert.deepEqual(calls.filter(c => c.method === "DELETE").map(c => c.url.split("/").pop()), ["7", "8"]);
  assert.equal(storage.get("cg.telemetry.cleanup").length, 10);   // remembered as today's date
});

/* ---------------- "A new version is available" bar ---------------- */
test("new version bar: shown when the page or shared files change, never on a local copy", async () => {
  let version = "v1";
  const head = (m, url) => m === "HEAD" ? { status: 200, ok: true, headers: { get: k => (k === "ETag" ? `"${version}-${url.split("/").pop()}"` : null) }, json: async () => ({}) } : null;
  const live = load({ protocol: "https:", route: head });
  await live.start();
  await new Promise(r => setImmediate(r));                  // baseline taken
  assert.equal(await live.CG.updates.check(), false);      // nothing new yet
  version = "v2";                                           // an update is published
  assert.equal(await live.CG.updates.check(), true);
  assert.equal(live.CG.updates._state().shown, true);
  assert.equal(await live.CG.updates.check(), false);      // shown once, not stacked
  assert.equal(live.calls.filter(c => c.method === "HEAD").length, 9);   // baseline + 2 checks, 3 files each; none once shown

  const local = load({ route: head });                      // no https: tests / file copies
  await local.start();
  assert.equal(local.CG.updates._state().on, false);
  assert.equal(local.calls.filter(c => c.method === "HEAD").length, 0);
});
