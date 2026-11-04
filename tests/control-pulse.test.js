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
  assert.equal(sla.lists[1].freshness, "unknown");

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
    assert.ok(h.lists.length > 0);
    h.lists.forEach(([label, list]) => { assert.ok(label && list, `${h.app}: label and list name`); });
  });
});
