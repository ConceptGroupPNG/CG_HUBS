// Tests for control/pulse.js (Mission Control's Data Pulse logic) - run with:  node --test tests/
// No browser, no Microsoft Graph: only the calculations behind the screen.
const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const P = require("../control/pulse.js");

// Local times, so the tests behave the same in Port Moresby and on GitHub's (UTC) test machine.
const NOW = new Date(2026, 10, 4, 14, 17);          // Wed 04-Nov-2026 14:17
const at = (d, h, m) => new Date(2026, 10, d, h, m); // a time in November 2026
const MIN = 60000, DAY = 86400000;

test("ageText", () => {
  assert.equal(P.ageText(20 * 1000), "now");
  assert.equal(P.ageText(65 * MIN), "1h 05m");
  assert.equal(P.ageText(230 * MIN), "3h 50m");
  assert.equal(P.ageText(45 * MIN), "45m");
  assert.equal(P.ageText(97 * DAY + 5 * MIN), "97d");
  assert.equal(P.ageText(NaN), "—");
});

test("freshness uses the local calendar day, not 24 hours", () => {
  assert.equal(P.freshness(at(4, 0, 5), NOW), "today");
  assert.equal(P.freshness(at(3, 23, 59), NOW), "week");          // 14 hours ago, but yesterday
  assert.equal(P.freshness(new Date(NOW - 7 * DAY), NOW), "week");
  assert.equal(P.freshness(new Date(NOW - 22 * DAY), NOW), "month");
  assert.equal(P.freshness(new Date(NOW - 80 * DAY), NOW), "quiet");
  assert.equal(P.freshness(null, NOW), "unknown");
});

test("stamp and who", () => {
  assert.equal(P.stamp(at(4, 13, 12)), "04-Nov 13:12");
  assert.equal(P.who({ user: { displayName: "Carlos Atalou" } }), "Carlos Atalou");
  assert.equal(P.who({ application: { displayName: "Power Automate" } }), "Power Automate (app)");
  assert.equal(P.who(null), "Unknown");
});

test("fromItem tells added from updated", () => {
  const added = P.fromItem({ createdDateTime: "2026-11-04T03:12:00Z", lastModifiedDateTime: "2026-11-04T03:12:02Z", lastModifiedBy: { user: { displayName: "A" } } });
  const updated = P.fromItem({ createdDateTime: "2026-10-01T00:00:00Z", lastModifiedDateTime: "2026-11-04T03:12:00Z", lastModifiedBy: { user: { displayName: "B" } } });
  assert.equal(added.kind, "added");
  assert.equal(updated.kind, "updated");
  assert.equal(updated.by, "B");
});

test("summarise: hub status, people, hourly writes, events and users", () => {
  const w = (d, h, m, by, kind = "updated") => ({ at: at(d, h, m), by, kind });
  const reads = [
    { app: "psp", hubName: "PSP HUB", label: "Calendar", list: "PSP_Act_Calendar", ok: true,
      items: [w(4, 13, 12, "Carlos Atalou"), w(4, 10, 40, "Carlos Atalou"), w(2, 9, 0, "Miriam Here")] },
    { app: "psp", hubName: "PSP HUB", label: "Billing", list: "PSP_Billing", ok: true,
      items: [w(4, 10, 19, "Brendan Kodana", "added")] },
    { app: "csa", hubName: "CSA HUB", label: "Clients", list: "CSA_Clients", ok: true,
      items: [], last: { at: new Date(NOW - 80 * DAY), by: "Marcus Scully" } },
    { app: "sla", hubName: "SLA HUB", label: "Tickets", list: "SLA_Tickets", ok: false, error: "No permission to read this list" },
    { app: "sla", hubName: "SLA HUB", label: "Agreements", list: "SLA_Agreements", ok: true, items: [], last: null, note: "List is over 5,000 items" }
  ];
  const s = P.summarise(reads, NOW);

  const psp = s.hubs.find(h => h.app === "psp");
  assert.equal(psp.status, "ACTIVE");
  assert.equal(psp.writesToday, 3);
  assert.equal(psp.people, 2);
  assert.equal(psp.age, "1h 05m");
  assert.equal(psp.lists[0].freshness, "today");
  assert.equal(psp.userList[0].name, "Carlos Atalou");     // most recent first
  assert.equal(psp.userList[0].today, 2);
  assert.equal(psp.userList[1].name, "Brendan Kodana");
  assert.deepEqual(psp.userList.find(u => u.name === "Miriam Here").lists, ["Calendar"]);

  const csa = s.hubs.find(h => h.app === "csa");
  assert.equal(csa.status, "QUIET");
  assert.equal(csa.lists[0].freshness, "quiet");
  assert.equal(csa.lists[0].last.by, "Marcus Scully");

  const sla = s.hubs.find(h => h.app === "sla");
  assert.equal(sla.status, "PARTIAL");
  assert.equal(sla.errors.length, 1);
  assert.equal(sla.lists[1].note, "List is over 5,000 items");
  assert.equal(sla.lists[1].freshness, "empty");          // nothing in it yet: grey, not red
  assert.equal(sla.lists[0].freshness, "unknown");        // couldn't be read: red

  assert.equal(s.totals.writesToday, 3);
  assert.equal(s.totals.peopleToday, 2);
  assert.equal(s.totals.failing, 1);
  assert.equal(s.hourly[10], 2);
  assert.equal(s.hourly[13], 1);
  assert.equal(s.totals.busiestHour, 10);
  assert.equal(s.events[0].by, "Carlos Atalou");          // newest first
  assert.equal(s.events.length, 4);
});

test("a Hub changed today only in a large list (time known, items not) is ACTIVE", () => {
  const s = P.summarise([{ app: "time", hubName: "CG Time", label: "Attendance", list: "CG_EmployeeTimeRecords", ok: true,
    items: [], last: { at: at(4, 14, 12), by: "Not known (large list)" }, note: "List is over 5,000 items" }], NOW);
  assert.equal(s.hubs[0].status, "ACTIVE");
  assert.equal(s.hubs[0].people, 0);
});

test("a Hub whose lists all fail is OFFLINE", () => {
  const s = P.summarise([{ app: "time", hubName: "CG Time", label: "Time Entries", list: "CSA_TimeEntries", ok: false, error: "x" }], NOW);
  assert.equal(s.hubs[0].status, "OFFLINE");
});

test("Graph addresses: filter is a UTC moment, names are encoded", () => {
  const u = P.recentUrl("site1", "CSA's List", new Date(Date.UTC(2026, 9, 28, 4, 17)));
  assert.match(u, /lists\/CSA's%20List\/items/);
  assert.match(u, /fields\/Modified ge '2026-10-28T04:17:00.000Z'/);
  assert.match(u, /\$orderby=fields\/Modified desc/);
  assert.match(P.latestUrl("site1", "SLA_Tickets"), /\$top=1$/);
});

test("explain gives plain-English reasons", () => {
  assert.match(P.explain({ status: 404, message: "itemNotFound" }), /not found/);
  assert.match(P.explain({ status: 403, message: "Access denied" }), /permission/);
  assert.match(P.explain({ status: 400, message: "The attempted operation is prohibited because it exceeds the list view threshold." }), /5,000/);
});

test("cg-config.js: every Mission Control Hub and site exists", () => {
  const ctx = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname, "../shared/cg-config.js"), "utf8") +
    "\nthis.CG_APPS = CG_APPS; this.CG_MONITOR = CG_MONITOR;", ctx);
  const control = ctx.CG_APPS.find(a => a.key === "control");
  assert.ok(control, "Mission Control is in CG_APPS");
  assert.deepEqual(Array.from(control.scopes), ["Sites.Read.All", "User.Read"]);   // read-only
  ctx.CG_MONITOR.hubs.forEach(h => {
    const app = ctx.CG_APPS.find(a => a.key === h.app);
    assert.ok(app, `${h.app} is a Hub in CG_APPS`);
    assert.ok(h.site || app.site, `${h.app} has a SharePoint site`);
    h.lists.forEach(([label, list]) => { assert.ok(label && list, `${h.app}: label and list name`); });
  });
});

test("programs (SharePoint App, flows) are shown but never counted as people; their runs are folded", () => {
  const app = { user: { displayName: "SharePoint App" } };
  const flow = { application: { displayName: "Power Automate" } };
  assert.equal(P.isApp(app), true);
  assert.equal(P.isApp(flow), true);
  assert.equal(P.isApp({ user: { displayName: "Miriam Here" } }), false);

  const items = [];
  for (let m = 20; m >= 6; m -= 2) items.push({ at: at(4, 15, m), by: "SharePoint App", kind: "updated", isApp: true });
  items.splice(3, 0, { at: at(4, 15, 13), by: "Peter Kent", kind: "added", isApp: false });
  const s = P.summarise([{ app: "time", hubName: "CG Time", label: "Attendance Records", list: "CG_EmployeeTimeRecords", ok: true, items }], NOW);
  const h = s.hubs[0];
  assert.equal(h.people, 1);                                      // Peter only
  assert.equal(s.totals.peopleToday, 1);
  assert.equal(h.writesToday, 9);                                 // the chart still counts every write
  assert.deepEqual(s.events.map(e => [e.by, e.count]), [["SharePoint App", 4], ["Peter Kent", 1], ["SharePoint App", 4]]);
  assert.equal(h.userList[h.userList.length - 1].name, "SharePoint App");   // programs listed after people
});

test("Hubs with no watched lists (phone apps) still appear and aren't 'unreachable'", () => {
  const s = P.summarise([], NOW, { hubs: [{ app: "mobile", name: "Concept Time Mobile" }] });
  assert.equal(s.hubs.length, 1);
  assert.equal(s.hubs[0].status, "QUIET");
});

test("withTelemetry: online now, people today, errors, load time, builds and who is behind", () => {
  const reads = [{ app: "psp", hubName: "PSP HUB", label: "Billing", list: "PSP_Billing", ok: true,
                   items: [{ at: at(4, 10, 19), by: "Brendan Kodana", kind: "updated" }] }];
  const s = P.summarise(reads, NOW, { hubs: [{ app: "psp", name: "PSP HUB" }, { app: "mobile", name: "Concept Time Mobile" }] });
  const sess = (app, user, lastSeen, build, loadMs, device = "Desktop") =>
    ({ app, user, started: new Date(lastSeen - 30 * MIN), lastSeen, build, loadMs, device, views: 1, throttles: 0 });
  const NEW = "core 1.7.0 · page 2026-11-04 15:20", OLD = "core 1.6.0 · page 2026-10-30 09:00";
  P.withTelemetry(s, { ok: true,
    sessions: [
      sess("psp", "Carlos Atalou", at(4, 14, 15), NEW, 1800),     // 2 min ago -> online
      sess("psp", "Miriam Here", at(4, 13, 0), OLD, 2600),        // today, not online, old build
      sess("psp", "Brendan Kodana", at(3, 16, 0), NEW, 1500),     // yesterday
      sess("mobile", "Peter Kent", at(4, 14, 14), NEW, 900, "Phone")
    ],
    errors: [
      { app: "psp", at: at(4, 13, 50), user: "Miriam Here", kind: "graph", message: "500 Server error" },
      { app: "psp", at: new Date(NOW - 2 * DAY), user: "Miriam Here", kind: "script", message: "old" }
    ] }, NOW);

  const psp = s.hubs.find(h => h.app === "psp");
  assert.equal(psp.online, 1);
  assert.deepEqual(psp.onlineNames, ["Carlos Atalou"]);
  assert.equal(psp.people, 3);                    // Brendan (wrote today) + Carlos + Miriam
  assert.equal(psp.errors24h, 1);
  assert.equal(psp.status, "ERRORS");             // an error in the last hour
  assert.equal(psp.loadMs, 2200);                 // median of today's visits
  assert.equal(psp.builds[0].build, NEW);
  assert.deepEqual(psp.behind, ["Brendan Kodana", "Miriam Here"]);   // Brendan: wrote today, last visit yesterday
  assert.deepEqual(psp.oldCopy, ["Brendan Kodana"]);
  assert.equal(psp.userList[0].name, "Carlos Atalou");   // online first
  assert.equal(psp.userList[0].online, true);

  const mob = s.hubs.find(h => h.app === "mobile");
  assert.equal(mob.status, "ACTIVE");
  assert.equal(mob.online, 1);

  assert.equal(s.totals.online, 2);
  assert.equal(s.totals.peopleToday, 4);
  assert.equal(s.totals.errors24h, 1);
  assert.equal(s.totals.failing, 1);
  assert.ok(s.events.some(e => e.kind === "opened" && e.by === "Peter Kent" && e.label === "Phone"));
  assert.ok(s.events.some(e => e.kind === "error" && e.label === "500 Server error"));
});

test("withTelemetry when the lists aren't there yet: marked off, nothing else changes", () => {
  const s = P.summarise([], NOW, { hubs: [{ app: "psp", name: "PSP HUB" }] });
  P.withTelemetry(s, { ok: false, reason: "List not found" }, NOW);
  assert.equal(s.telemetry.ok, false);
  assert.equal(s.hubs[0].online, undefined);
});

test("fromSession / fromError read the list rows", () => {
  const x = P.fromSession({ fields: { App: "sla", UserName: "Mathias Sikari", LastSeen: "2026-11-04T04:10:00Z", StartedAt: "2026-11-04T03:00:00Z", LoadMs: 1200, Build: "b", Device: "Phone" } });
  assert.equal(x.app, "sla"); assert.equal(x.loadMs, 1200); assert.equal(x.device, "Phone");
  assert.equal(x.lastSeen.toISOString(), "2026-11-04T04:10:00.000Z");
  const e = P.fromError({ createdDateTime: "2026-11-04T04:11:00Z", fields: { App: "sla", UserName: "Mathias Sikari", Title: "graph", Message: "429 busy" } });
  assert.equal(e.kind, "graph"); assert.equal(e.message, "429 busy");
});

test("old copies: wrote today with no visit today is flagged; recent writers count as online", () => {
  const w = (h, m, by) => ({ at: at(4, h, m), by, kind: "updated" });
  const reads = [{ app: "psp", hubName: "PSP HUB", label: "Calendar", list: "PSP_Act_Calendar", ok: true,
                   items: [w(14, 15, "Gloria Boku"), w(7, 31, "Jacinta Kodana"), { at: at(4, 14, 16), by: "SharePoint App", kind: "updated", isApp: true }] }];
  const s = P.summarise(reads, NOW, { hubs: [{ app: "psp", name: "PSP HUB" }] });
  P.withTelemetry(s, { ok: true, errors: [], sessions: [
    { app: "psp", user: "Marcus Scully", started: at(4, 13, 0), lastSeen: at(4, 14, 16), build: "b2", device: "Desktop" } ] }, NOW);
  const psp = s.hubs[0];
  assert.deepEqual(psp.oldCopy, ["Gloria Boku", "Jacinta Kodana"]);          // never the program
  assert.deepEqual(psp.onlineNames, ["Gloria Boku", "Marcus Scully"]);       // Gloria wrote 2 minutes ago
  assert.equal(psp.userList.find(u => u.name === "Jacinta Kodana").online, false);
  assert.deepEqual(psp.behind, ["Gloria Boku", "Jacinta Kodana"]);
  assert.equal(s.totals.online, 2);
});

test("open vs active: open in the background counts as open, not active", () => {
  const s = P.summarise([], NOW, { hubs: [{ app: "psp", name: "PSP HUB" }] });
  const sess = (user, lastSeen, lastActive) => ({ app: "psp", user, started: at(4, 8, 0), lastSeen, lastActive, build: "b", device: "Desktop" });
  P.withTelemetry(s, { ok: true, errors: [], sessions: [
    sess("Carlos Atalou", at(4, 14, 16), at(4, 14, 15)),     // open and using it
    sess("Gloria Boku", at(4, 14, 16), at(4, 13, 40)),       // open, idle 37 min (on Excel)
    sess("Owen Boku", at(4, 13, 0), at(4, 12, 59)),          // closed an hour ago
    sess("Miriam Here", at(4, 14, 14), null)                 // older core: only reported while on screen
  ] }, NOW);
  const h = s.hubs[0];
  assert.deepEqual(h.onlineNames, ["Carlos Atalou", "Gloria Boku", "Miriam Here"]);
  assert.deepEqual(h.activeNames, ["Carlos Atalou", "Miriam Here"]);
  assert.equal(h.online, 3); assert.equal(h.active, 2);
  assert.equal(s.totals.online, 3); assert.equal(s.totals.active, 2);
  const g = h.userList.find(u => u.name === "Gloria Boku");
  assert.equal(g.online, true); assert.equal(!!g.active, false);
  assert.equal(g.lastActive.getTime(), at(4, 13, 40).getTime());
  assert.deepEqual(h.userList.map(u => u.name).slice(0, 3), ["Carlos Atalou", "Miriam Here", "Gloria Boku"]);   // active first
});
