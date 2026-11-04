/* =====================================================================
   Mission Control - the Data Pulse logic (no screen, no Graph calls).

   Kept separate from control/index.html so it can be tested on its own:
   tests/control-pulse.test.js. Exposes one global, `CGPulse`.

   Times: every "today" / hour test uses the browser's local calendar
   (Port Moresby, UTC+10), never UTC.
   ===================================================================== */
(function (global) {
  "use strict";

  const MIN = 60 * 1000, HOUR = 60 * MIN, DAY = 24 * HOUR;

  const P = {};

  // Same local calendar day?
  P.sameDay = (a, b) =>
    a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();

  // "45m", "3h 50m", "8d"  (under a minute: "now")
  P.ageText = (ms) => {
    if (!(ms >= 0)) return "—";
    if (ms < MIN) return "now";
    if (ms < HOUR) return Math.floor(ms / MIN) + "m";
    if (ms < DAY) return Math.floor(ms / HOUR) + "h " + String(Math.floor((ms % HOUR) / MIN)).padStart(2, "0") + "m";
    return Math.floor(ms / DAY) + "d";
  };

  // How fresh a list's last write is: "today" | "week" | "month" | "quiet" | "unknown"
  P.freshness = (at, now) => {
    if (!at) return "unknown";
    if (P.sameDay(at, now)) return "today";
    const age = now - at;
    if (age <= 7 * DAY) return "week";
    if (age <= 30 * DAY) return "month";
    return "quiet";
  };

  // "04-Nov 13:12"
  const MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const two = n => String(n).padStart(2, "0");
  P.stamp = (d) => d ? `${two(d.getDate())}-${MON[d.getMonth()]} ${two(d.getHours())}:${two(d.getMinutes())}` : "—";
  P.clock = (d) => `${two(d.getHours())}:${two(d.getMinutes())}`;

  // Who made a change, from a Graph identitySet (a person, or a flow / app).
  P.who = (identitySet) => {
    const s = identitySet || {};
    return (s.user && (s.user.displayName || s.user.email)) ||
           (s.application && (s.application.displayName ? s.application.displayName + " (app)" : "")) || "Unknown";
  };

  /* One Graph list item -> { at, by, kind }.
     kind is "added" when it was created at (about) the same moment it was
     last changed, otherwise "updated". */
  P.fromItem = (item) => {
    const at = new Date(item.lastModifiedDateTime);
    const created = item.createdDateTime ? new Date(item.createdDateTime) : null;
    const kind = created && Math.abs(at - created) < 5000 ? "added" : "updated";
    return { at, by: P.who(item.lastModifiedBy), kind };
  };

  /* Turns what was read from SharePoint into everything the screen shows.

     reads: one entry per watched list:
       { app, hubName, label, list,
         ok: true,  items: [{ at, by, kind }...] (newest first, within the window),
                    last: { at, by } | null   (the newest write ever, when known),
                    capped: bool              (more writes than were read)
                    note: "text"              (optional: read only partly, and why)
         ok: false, error: "message" }
     now: Date

     Returns { hubs: [...], events: [...], hourly: [24 numbers], totals: {...} } */
  P.summarise = (reads, now, opts = {}) => {
    const maxEvents = opts.maxEvents || 15;
    const hubs = [], byApp = {};
    const events = [];
    const hourly = new Array(24).fill(0);
    const peopleToday = new Set();
    let writesToday = 0;

    reads.forEach(r => {
      let hub = byApp[r.app];
      if (!hub) {
        hub = byApp[r.app] = { app: r.app, name: r.hubName, lists: [], users: {}, errors: [],
                               writesToday: 0, peopleToday: new Set(), last: null };
        hubs.push(hub);
      }
      if (!r.ok) {
        hub.errors.push({ list: r.list, label: r.label, message: r.error || "Couldn't be read" });
        hub.lists.push({ label: r.label, list: r.list, ok: false, error: r.error, freshness: "unknown", last: null, age: "—" });
        return;
      }
      const items = r.items || [];
      const last = r.last || (items[0] ? { at: items[0].at, by: items[0].by } : null);
      hub.lists.push({
        label: r.label, list: r.list, ok: true, capped: !!r.capped, note: r.note || "",
        last, freshness: P.freshness(last && last.at, now),
        age: last ? P.ageText(now - last.at) : "—",
        stamp: last ? P.stamp(last.at) : "—",
        writes: items.length, writesToday: items.filter(i => P.sameDay(i.at, now)).length
      });
      if (last && (!hub.last || last.at > hub.last.at)) hub.last = last;

      items.forEach(i => {
        const u = hub.users[i.by] || (hub.users[i.by] = { name: i.by, writes: 0, today: 0, last: null, lists: new Set() });
        u.writes++;
        u.lists.add(r.label);
        if (!u.last || i.at > u.last) u.last = i.at;
        if (P.sameDay(i.at, now)) {
          u.today++;
          hub.writesToday++; writesToday++;
          hub.peopleToday.add(i.by); peopleToday.add(i.by);
          hourly[i.at.getHours()]++;
        }
        events.push({ at: i.at, hub: r.hubName, label: r.label, by: i.by, kind: i.kind });
      });
    });

    hubs.forEach(h => {
      const okLists = h.lists.filter(l => l.ok).length;
      const today = h.writesToday || h.lists.some(l => l.freshness === "today");
      h.status = okLists === 0 ? "OFFLINE" : h.errors.length ? "PARTIAL" : today ? "ACTIVE" : "QUIET";
      h.people = h.peopleToday.size;
      h.age = h.last ? P.ageText(now - h.last.at) : "—";
      h.userList = Object.values(h.users)
        .map(u => ({ name: u.name, writes: u.writes, today: u.today, last: u.last, lists: Array.from(u.lists).sort() }))
        .sort((a, b) => b.last - a.last);
      delete h.users; delete h.peopleToday;
    });

    events.sort((a, b) => b.at - a.at);
    const failing = hubs.filter(h => h.status === "OFFLINE" || h.status === "PARTIAL").length;
    return {
      hubs,
      events: events.slice(0, maxEvents),
      hourly,
      totals: { writesToday, peopleToday: peopleToday.size, failing, hubs: hubs.length,
                busiestHour: hourly.reduce((best, v, h) => (v > hourly[best] ? h : best), 0) }
    };
  };

  /* The Graph address that reads one list's recent writes, newest first.
     `since` is a moment in time, so sending it as UTC (toISOString) is
     correct here - this is not a calendar date. */
  P.recentUrl = (siteId, list, since, top = 500) =>
    `/sites/${siteId}/lists/${encodeURIComponent(list)}/items` +
    `?$select=id,createdDateTime,lastModifiedDateTime,lastModifiedBy` +
    `&$expand=fields($select=Modified)` +
    `&$filter=fields/Modified ge '${since.toISOString()}'` +
    `&$orderby=fields/Modified desc&$top=${top}`;

  P.latestUrl = (siteId, list) =>
    `/sites/${siteId}/lists/${encodeURIComponent(list)}/items` +
    `?$select=id,createdDateTime,lastModifiedDateTime,lastModifiedBy` +
    `&$expand=fields($select=Modified)&$orderby=fields/Modified desc&$top=1`;

  // A plain-English reason for a list that couldn't be read.
  P.explain = (e) => {
    const m = String((e && e.message) || e || "");
    if (e && (e.status === 404 || /not (be )?found|does not exist/i.test(m))) return "List not found on this site - check its name in cg-config.js";
    if (e && (e.status === 403 || /access denied|forbidden/i.test(m))) return "No permission to read this list";
    if (/threshold|too many|prohibited|not indexed/i.test(m)) return "List is over 5,000 items - add an index on the Modified column";
    return m || "Couldn't be read";
  };

  global.CGPulse = P;
  if (typeof module !== "undefined" && module.exports) module.exports = P;
})(typeof window !== "undefined" ? window : globalThis);
