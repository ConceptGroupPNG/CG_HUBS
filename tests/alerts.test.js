// Tests for the alerts job (alerts/rules.js, alerts/run.js) - run with:  node --test tests/
// SharePoint, Microsoft sign-in and Teams are simulated; nothing real is contacted.
const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const os = require("os");
const path = require("path");
const R = require("../alerts/rules.js");
const { run, loadConfig } = require("../alerts/run.js");

const NOW = new Date(2026, 10, 4, 14, 17);
const MIN = 60000, HOUR = 60 * MIN;
const ago = ms => new Date(NOW - ms);
const CFG = { remindHours: 6, hubProblems: { errorsPerHour: 3 }, bridge: { key: "zkteco-bridge", downMinutes: 15, deviceMinutes: 15 } };

const hub = (app, name, extra = {}) => Object.assign({ app, name, status: "ACTIVE", errors: [], lists: [], userErrors: [] }, extra);
// A heartbeat row as P.readHeartbeats returns it
const beat = (extra = {}) => ({ ok: true, rows: { "zkteco-bridge": Object.assign({ key: "zkteco-bridge", lastSeen: ago(3 * MIN),
  deviceSerial: "NES1", deviceLastSeen: ago(20 * 1000), lastSyncAt: ago(4 * MIN), lastSyncError: "", version: "1.1.0", host: "BioDevAirvos2" }, extra) } });

/* ---------------- rules ---------------- */
test("Hub problems: unreadable, partly unreadable, and errors per hour", () => {
  const s = { hubs: [
    hub("csa", "CSA HUB", { status: "OFFLINE", errors: [{ label: "Clients", list: "CSA_Clients", message: "No permission" }] }),
    hub("sla", "SLA HUB", { status: "PARTIAL", errors: [{ label: "Tickets", list: "SLA_Tickets", message: "List not found" }] }),
    hub("psp", "PSP HUB", { status: "ERRORS", userErrors: [
      { at: ago(5 * MIN), user: "Miriam Here", message: "500 Server error" },
      { at: ago(20 * MIN), user: "Carlos Atalou", message: "429 busy" },
      { at: ago(50 * MIN), user: "Miriam Here", message: "x is not defined" } ] }),
    hub("training", "CG Training HUB", { userErrors: [{ at: ago(5 * MIN), user: "A", message: "one-off" }, { at: ago(2 * HOUR), user: "A", message: "old" }] })
  ] };
  const p = R.evaluate(s, CFG, NOW);
  assert.deepEqual(p.map(x => x.key), ["hub:csa:offline", "hub:sla:lists", "hub:psp:errors"]);
  assert.match(p[2].title, /3 errors in the last hour/);
  assert.match(p[2].text, /Carlos Atalou, Miriam Here/);
});

test("ZKTeco bridge from its heartbeat: running, down, device lost, sync failing, not installed yet", () => {
  const withBeat = hb => ({ hubs: [], heartbeats: hb });
  const keys = hb => R.evaluate(withBeat(hb), CFG, NOW).map(x => x.key);
  assert.deepEqual(keys(beat()), []);                                                   // all well
  assert.deepEqual(keys(beat({ lastSeen: ago(40 * MIN) })), ["bridge:down"]);
  assert.match(R.evaluate(withBeat(beat({ lastSeen: ago(40 * MIN) })), CFG, NOW)[0].text, /on BioDevAirvos2/);
  assert.deepEqual(keys(beat({ deviceLastSeen: ago(30 * MIN) })), ["bridge:device"]);
  assert.deepEqual(keys(beat({ deviceLastSeen: null })), ["bridge:device"]);
  assert.deepEqual(keys(beat({ lastSyncError: "Graph GET -> 401" })), ["bridge:sync"]);
  assert.deepEqual(keys({ ok: true, rows: {} }), []);                                   // heartbeat version not installed yet
  assert.deepEqual(keys({ ok: false, reason: "The CG_Heartbeats list isn't set up yet." }), []);
  assert.deepEqual(R.evaluate(withBeat(beat({ lastSeen: ago(2 * HOUR) })), Object.assign({}, CFG, { bridge: { enabled: false } }), NOW), []);
  // ...but it always says why it can't check, rather than staying silent
  assert.match(R.bridgeStatus(withBeat({ ok: true, rows: {} }), CFG, NOW), /not reporting yet/);
  assert.match(R.bridgeStatus(withBeat({ ok: false, reason: "The CG_Heartbeats list isn't set up yet." }), CFG, NOW), /can't check/);
  assert.match(R.bridgeStatus(withBeat(beat()), CFG, NOW), /^bridge running; device checked in just now$/);
});

test("lists read only partly are always reported", () => {
  const s = { hubs: [{ lists: [{ list: "CG_EmployeeTimeRecords", ok: true, note: "List is over 5,000 items" }, { list: "SLA_Tickets", ok: true }] }] };
  assert.deepEqual(R.partial(s), ["CG_EmployeeTimeRecords"]);
});

test("announced once, reminded every 6 hours, resolved when it clears", () => {
  const p = [{ key: "bridge:silent", title: "ZKTeco bridge silent for 45m", text: "…" }];
  let r = R.diff(p, {}, CFG, NOW);
  assert.deepEqual(r.messages.map(m => m.type), ["new"]);

  let later = new Date(+NOW + 10 * MIN);
  r = R.diff(p, r.memory, CFG, later);
  assert.equal(r.messages.length, 0);                       // no repeat 10 minutes later

  later = new Date(+NOW + 6 * HOUR + MIN);
  r = R.diff(p, r.memory, CFG, later);
  assert.deepEqual(r.messages.map(m => m.type), ["reminder"]);
  assert.equal(r.memory["bridge:silent"].since, NOW.toISOString());   // still counted from the start

  r = R.diff([], r.memory, CFG, new Date(+later + 10 * MIN));
  assert.deepEqual(r.messages.map(m => [m.type, m.title]), [["resolved", "ZKTeco bridge silent for 45m"]]);
  assert.deepEqual(r.memory, {});
});

test("Teams card and the public log line", () => {
  const c = R.card([{ type: "new", title: "PSP HUB: 3 errors in the last hour", text: "People affected: Miriam Here" },
                    { type: "resolved", title: "ZKTeco bridge silent", text: "", since: ago(HOUR) }], NOW, "https://x/control/");
  const card = c.attachments[0];
  assert.equal(card.contentType, "application/vnd.microsoft.card.adaptive");
  const texts = card.content.body.map(b => b.text).join("\n");
  assert.match(texts, /🔴 New: PSP HUB: 3 errors/);
  assert.match(texts, /✅ Resolved: ZKTeco bridge silent · since/);
  assert.equal(card.content.actions[0].url, "https://x/control/");
  const line = R.logLine([{ key: "hub:psp:errors", text: "People affected: Miriam Here" }], [{ type: "new" }]);
  assert.doesNotMatch(line, /Miriam/);
  assert.match(line, /hub:psp:errors/);
});

/* ---------------- the whole job, against a simulated SharePoint + Teams ---------------- */
const json = (status, body) => ({ status, ok: status >= 200 && status < 300, headers: { get: () => null }, json: async () => body });
const iso = d => d.toISOString();
const item = (at, who, app) => ({ id: "1", createdDateTime: iso(at), lastModifiedDateTime: iso(at),
  lastModifiedBy: app ? { application: { displayName: who } } : { user: { displayName: who } } });

function fakeFetch({ bridgeAt = ago(5 * MIN), heartbeat = { lastSeen: ago(2 * MIN), deviceLastSeen: ago(10 * 1000) }, signInFails = false, errors = [] } = {}) {
  const posts = [];
  const f = async (url, opts = {}) => {
    url = decodeURIComponent(url);
    if (url.startsWith("https://login.microsoftonline.com/")) return signInFails ? json(401, { error: "invalid_client" }) : json(200, { access_token: "app-token" });
    if (url.startsWith("https://teams.example/")) { posts.push(JSON.parse(opts.body)); return json(202, {}); }
    if (/\/sites\/[^/]+:\/sites\//.test(url)) return json(200, { id: "site-" + url.split(":").pop() });
    if (url.includes("/lists/CG_HubSessions/")) return json(200, { value: [] });
    if (url.includes("/lists/CG_HubErrors/")) return json(200, { value: errors.map(e => ({ createdDateTime: iso(e.at), fields: { App: e.app, UserName: e.user, Title: "graph", Message: e.message } })) });
    if (url.includes("/lists/CG_EmployeeTimeRecords/")) return json(200, { value: [item(bridgeAt, "SharePoint App", true)] });
    if (url.includes("/lists/CG_Heartbeats/")) return json(200, { value: heartbeat ? [{ lastModifiedDateTime: iso(heartbeat.lastSeen), fields: {
      Title: "zkteco-bridge", LastSeen: iso(heartbeat.lastSeen), DeviceSerial: "NES1", DeviceLastSeen: heartbeat.deviceLastSeen ? iso(heartbeat.deviceLastSeen) : undefined,
      LastSyncAt: iso(ago(5 * MIN)), LastSyncError: heartbeat.error || "", Version: "1.1.0", Host: "BioDevAirvos2" } }] : [] });
    return json(200, { value: [] });
  };
  return { f, posts };
}

function tmpState() { return path.join(fs.mkdtempSync(path.join(os.tmpdir(), "alerts-")), "state.json"); }
const config = () => { const c = loadConfig(); c.CG_ALERTS.clientId = "11111111-2222-3333-4444-555555555555"; return c; };
const ENV = file => ({ ALERTS_CLIENT_SECRET: "s", TEAMS_WEBHOOK_URL: "https://teams.example/hook", STATE_FILE: file });

test("not set up yet: does nothing, stays green", async () => {
  const { f, posts } = fakeFetch();
  const line = await run({ STATE_FILE: tmpState() }, { fetch: f, now: NOW, config: loadConfig() });
  assert.match(line, /aren't set up yet/);
  assert.equal(posts.length, 0);
});

test("all well: nothing sent; a test run sends 'connected' with what it can see", async () => {
  const file = tmpState();
  const a = fakeFetch();
  assert.match(await run(ENV(file), { fetch: a.f, now: NOW, config: config() }), /; 0 problem/);
  assert.equal(a.posts.length, 0);

  const b = fakeFetch();
  await run(Object.assign(ENV(file), { SEND_TEST: "true" }), { fetch: b.f, now: NOW, config: config() });
  assert.equal(b.posts.length, 1);
  const text = b.posts[0].attachments[0].content.body.map(x => x.text).join("\n");
  assert.match(text, /Alerts are connected/);
  assert.match(text, /telemetry on \(sees 0 visits by 0 people and 0 errors in the last day\)/);
  assert.match(text, /ZKTeco bridge running; device checked in/);
  assert.match(text, /Problems right now: none/);
});

test("bridge stops, then comes back: one alert, one resolved, nothing in between", async () => {
  const file = tmpState();
  const down = fakeFetch({ heartbeat: { lastSeen: ago(50 * MIN), deviceLastSeen: ago(50 * MIN) } });
  const line = await run(ENV(file), { fetch: down.f, now: NOW, config: config() });
  assert.match(line, /bridge:down/);
  assert.doesNotMatch(line, /BioDevAirvos2/);                 // no PC names in the public log
  assert.equal(down.posts.length, 1);
  assert.match(JSON.stringify(down.posts[0]), /ZKTeco bridge not running \(no heartbeat for 50m\)/);

  const still = fakeFetch({ heartbeat: { lastSeen: ago(60 * MIN), deviceLastSeen: ago(60 * MIN) } });
  await run(ENV(file), { fetch: still.f, now: new Date(+NOW + 10 * MIN), config: config() });
  assert.equal(still.posts.length, 0);

  const back = fakeFetch({ heartbeat: { lastSeen: new Date(+NOW + 19 * MIN), deviceLastSeen: new Date(+NOW + 19 * MIN) } });
  await run(ENV(file), { fetch: back.f, now: new Date(+NOW + 20 * MIN), config: config() });
  assert.equal(back.posts.length, 1);
  assert.match(JSON.stringify(back.posts[0]), /Resolved/);
});

test("nobody scanning overnight is not an alert (the old false alarm)", async () => {
  const x = fakeFetch({ bridgeAt: ago(10 * HOUR) });            // no attendance writes for 10 hours, heartbeat fine
  assert.match(await run(ENV(tmpState()), { fetch: x.f, now: NOW, config: config() }), /0 problem/);
  assert.equal(x.posts.length, 0);
});

test("before the heartbeat version is installed: no alert, and the test message says so", async () => {
  const x = fakeFetch({ heartbeat: null });
  const line = await run(Object.assign(ENV(tmpState()), { SEND_TEST: "true" }), { fetch: x.f, now: NOW, config: config() });
  assert.match(line, /not reporting yet/);
  assert.match(JSON.stringify(x.posts[0]), /ZKTeco bridge not reporting yet/);
});

test("errors people hit reach Teams by name, but the log line has no names", async () => {
  const errs = [1, 2, 3].map(n => ({ app: "psp", at: ago(n * 5 * MIN), user: "Miriam Here", message: "500 Server error " + n }));
  const x = fakeFetch({ errors: errs });
  const line = await run(ENV(tmpState()), { fetch: x.f, now: NOW, config: config() });
  assert.match(line, /hub:psp:errors/);
  assert.doesNotMatch(line, /Miriam|Server error/);
  assert.match(JSON.stringify(x.posts[0]), /Miriam Here/);
});

test("sign-in fails (e.g. expired secret): Teams is told once and the run shows red", async () => {
  const file = tmpState();
  const x = fakeFetch({ signInFails: true });
  await assert.rejects(run(ENV(file), { fetch: x.f, now: NOW, config: config() }), /couldn't read SharePoint/);
  assert.equal(x.posts.length, 1);
  assert.match(JSON.stringify(x.posts[0]), /can't read SharePoint/);
  const y = fakeFetch({ signInFails: true });
  await assert.rejects(run(ENV(file), { fetch: y.f, now: new Date(+NOW + 10 * MIN), config: config() }));
  assert.equal(y.posts.length, 0);                          // not repeated every 10 minutes
});

test("cg-config.js: alert settings are complete", () => {
  const c = loadConfig();
  assert.ok(c.CG_ALERTS.remindHours > 0);
  assert.ok(c.CG_ALERTS.hubProblems.errorsPerHour > 0);
  const b = c.CG_ALERTS.bridge;
  assert.ok(b.key && b.downMinutes > 0 && b.deviceMinutes > 0);
  assert.equal(c.CG_TELEMETRY.heartbeats, "CG_Heartbeats");
});
