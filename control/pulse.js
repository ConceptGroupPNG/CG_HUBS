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

  // A change made by a program (a flow, the ZKTeco bridge, "SharePoint App")
  // rather than a person. These are shown, but never counted as people.
  P.isApp = (identitySet) => {
    const s = identitySet || {};
    if (s.application && !s.user) return true;
    return /^SharePoint App$/i.test(P.who(s)) || / \(app\)$/.test(P.who(s));
  };

  /* One Graph list item -> { at, by, kind, isApp }.
     kind is "added" when it was created at (about) the same moment it was
     last changed, otherwise "updated". */
  P.fromItem = (item) => {
    const at = new Date(item.lastModifiedDateTime);
    const created = item.createdDateTime ? new Date(item.createdDateTime) : null;
    const kind = created && Math.abs(at - created) < 5000 ? "added" : "updated";
    return { at, by: P.who(item.lastModifiedBy), kind, isApp: P.isApp(item.lastModifiedBy) };
  };

  /* The event stream: newest first, with runs of the same program writing to
     the same list folded into one line ("SharePoint App · Attendance ×14"). */
  P.feed = (events, max = 15) => {
    const sorted = events.slice().sort((a, b) => b.at - a.at);
    const out = [];
    sorted.forEach(e => {
      const prev = out[out.length - 1];
      if (prev && e.isApp && prev.isApp && prev.hub === e.hub && prev.label === e.label && prev.by === e.by) {
        prev.count = (prev.count || 1) + 1;
        prev.from = e.at;
        return;
      }
      if (out.length < max) out.push(Object.assign({ count: 1 }, e));
    });
    return out;
  };

  P.median = (nums) => {
    const v = nums.filter(n => n > 0).sort((a, b) => a - b);
    if (!v.length) return null;
    const m = Math.floor(v.length / 2);
    return v.length % 2 ? v[m] : Math.round((v[m - 1] + v[m]) / 2);
  };
  P.seconds = (ms) => (ms ? (ms / 1000).toFixed(1) + "s" : "—");

  /* Turns what was read from SharePoint into everything the screen shows.

     reads: one entry per watched list:
       { app, hubName, label, list,
         ok: true,  items: [{ at, by, kind }...] (newest first, within the window),
                    last: { at, by } | null   (the newest write ever, when known),
                    capped: bool              (more writes than were read)
                    note: "text"              (optional: read only partly, and why)
         ok: false, error: "message" }
     now: Date
     opts.hubs: [{ app, name }] - Hubs to show even when they have no lists
                (e.g. the phone apps, seen only through telemetry)

     Returns { hubs: [...], events: [...], rawEvents, hourly: [24 numbers], totals: {...} } */
  P.summarise = (reads, now, opts = {}) => {
    const maxEvents = opts.maxEvents || 15;
    const hubs = [], byApp = {};
    const events = [];
    const hourly = new Array(24).fill(0);
    const peopleToday = new Set();
    let writesToday = 0;
    const hubFor = (app, name) => {
      if (!byApp[app]) {
        byApp[app] = { app, name, lists: [], users: {}, errors: [], writesToday: 0, peopleToday: new Set(), last: null };
        hubs.push(byApp[app]);
      }
      return byApp[app];
    };
    (opts.hubs || []).forEach(h => hubFor(h.app, h.name));

    reads.forEach(r => {
      const hub = hubFor(r.app, r.hubName);
      if (!r.ok) {
        hub.errors.push({ list: r.list, label: r.label, message: r.error || "Couldn't be read" });
        hub.lists.push({ label: r.label, list: r.list, ok: false, error: r.error, freshness: "unknown", last: null, age: "—" });
        return;
      }
      const items = r.items || [];
      const last = r.last || (items[0] ? { at: items[0].at, by: items[0].by } : null);
      hub.lists.push({
        label: r.label, list: r.list, ok: true, capped: !!r.capped, note: r.note || "",
        last, freshness: last ? P.freshness(last.at, now) : "empty",
        age: last ? P.ageText(now - last.at) : "—",
        stamp: last ? P.stamp(last.at) : "—",
        writes: items.length, writesToday: items.filter(i => P.sameDay(i.at, now)).length,
        lastApp: (items.find(i => i.isApp) || {}).at || null     // newest write by a program (e.g. the ZKTeco bridge)
      });
      if (last && (!hub.last || last.at > hub.last.at)) hub.last = last;

      items.forEach(i => {
        const u = hub.users[i.by] || (hub.users[i.by] = { name: i.by, isApp: !!i.isApp, writes: 0, today: 0, last: null, lists: new Set() });
        u.writes++;
        u.lists.add(r.label);
        if (!u.last || i.at > u.last) u.last = i.at;
        if (P.sameDay(i.at, now)) {
          u.today++;
          hub.writesToday++; writesToday++;
          if (!i.isApp) { hub.peopleToday.add(i.by); peopleToday.add(i.by); }
          hourly[i.at.getHours()]++;
        }
        events.push({ at: i.at, hub: r.hubName, app: r.app, label: r.label, by: i.by, kind: i.kind, isApp: !!i.isApp });
      });
    });

    hubs.forEach(h => {
      const okLists = h.lists.filter(l => l.ok).length;
      const today = h.writesToday || h.lists.some(l => l.freshness === "today");
      h.status = (h.lists.length && okLists === 0) ? "OFFLINE" : h.errors.length ? "PARTIAL" : today ? "ACTIVE" : "QUIET";
      h.people = h.peopleToday.size;
      h.peopleNames = Array.from(h.peopleToday);
      h.age = h.last ? P.ageText(now - h.last.at) : "—";
      h.userList = Object.values(h.users)
        .map(u => ({ name: u.name, isApp: u.isApp, writes: u.writes, today: u.today, last: u.last, lists: Array.from(u.lists).sort(),
                     sessions: 0, lastSeen: null, online: false, device: "" }))
        .sort((a, b) => (a.isApp - b.isApp) || (b.last - a.last));
      delete h.users; delete h.peopleToday;
    });

    const failing = hubs.filter(h => h.status === "OFFLINE" || h.status === "PARTIAL").length;
    return {
      hubs,
      events: P.feed(events, maxEvents),
      rawEvents: events,
      hourly,
      peopleNames: Array.from(peopleToday),
      totals: { writesToday, peopleToday: peopleToday.size, failing, hubs: hubs.length,
                busiestHour: hourly.reduce((best, v, h) => (v > hourly[best] ? h : best), 0) }
    };
  };

  /* Adds what the Hubs reported themselves (telemetry) to a summary.

     tele: { ok: true, sessions: [...], errors: [...] }  or  { ok: false, reason }
       session: { app, user, started, lastSeen, views, screens, errors, throttles,
                  build, loadMs, device }
       error:   { app, at, user, kind, message, detail, page }
     Changes the summary in place and returns it. */
  P.withTelemetry = (summary, tele, now, opts = {}) => {
    const onlineMs = (opts.onlineMinutes || 6) * 60000;
    summary.telemetry = tele && tele.ok
      ? { ok: true, visits: tele.sessions.length, people: new Set(tele.sessions.map(x => x.user)).size, errors: tele.errors.length }
      : { ok: false, reason: (tele && tele.reason) || "" };
    if (!tele || !tele.ok) return summary;

    const allOnline = new Set(), allToday = new Set(summary.peopleNames || []);
    const extra = [];
    let errors24h = 0;
    summary.hubs.forEach(h => {
      const ss = tele.sessions.filter(x => x.app === h.app);
      const es = tele.errors.filter(x => x.app === h.app).sort((a, b) => b.at - a.at);
      const online = new Set(), today = new Set(h.peopleNames || []);
      const builds = {};
      ss.forEach(x => {
        if (now - x.lastSeen <= onlineMs) { online.add(x.user); allOnline.add(x.user); }
        if (P.sameDay(x.lastSeen, now)) { today.add(x.user); allToday.add(x.user); }
        let u = h.userList.find(v => v.name === x.user);
        if (!u) { u = { name: x.user, isApp: false, writes: 0, today: 0, last: null, lists: [], sessions: 0, lastSeen: null, online: false, device: "" }; h.userList.push(u); }
        u.sessions++;
        if (!u.lastSeen || x.lastSeen > u.lastSeen) { u.lastSeen = x.lastSeen; u.device = x.device || ""; u.build = x.build || ""; }
        if (now - x.lastSeen <= onlineMs) u.online = true;
        const b = builds[x.build || "?"] || (builds[x.build || "?"] = { build: x.build || "?", people: new Set(), lastSeen: x.lastSeen });
        b.people.add(x.user);
        if (x.lastSeen > b.lastSeen) b.lastSeen = x.lastSeen;
        if (x.started) extra.push({ at: x.started, hub: h.name, app: h.app, label: x.device || "Desktop", by: x.user, kind: "opened", isApp: false });
      });
      es.forEach(e => extra.push({ at: e.at, hub: h.name, app: h.app, label: e.message, by: e.user, kind: "error", isApp: false }));

      h.online = online.size;
      h.onlineNames = Array.from(online).sort();
      h.people = today.size;
      h.userErrors = es;
      h.errors24h = es.filter(e => now - e.at <= 86400000).length;
      errors24h += h.errors24h;
      h.loadMs = P.median(ss.filter(x => P.sameDay(x.lastSeen, now)).map(x => x.loadMs));
      h.throttles = ss.reduce((t, x) => t + (x.throttles || 0), 0);
      // Builds, newest first (the build text ends with the page's publish time).
      h.builds = Object.values(builds).map(b => ({ build: b.build, people: Array.from(b.people).sort(), lastSeen: b.lastSeen }))
        .sort((a, b) => b.build.localeCompare(a.build));
      const newest = h.builds[0] && h.builds[0].build;
      h.behind = h.userList.filter(u => u.build && newest && u.build !== newest).map(u => u.name).sort();
      h.userList.sort((a, b) => (b.online - a.online) || (a.isApp - b.isApp) ||
        ((b.lastSeen || b.last || 0) - (a.lastSeen || a.last || 0)));

      const recentErr = es.some(e => now - e.at <= 3600000);
      if (h.status === "ACTIVE" || h.status === "QUIET") {
        if (recentErr) h.status = "ERRORS";
        else if (h.online || h.people) h.status = "ACTIVE";
      }
    });
    summary.events = P.feed(summary.rawEvents.concat(extra), opts.maxEvents || 15);
    summary.totals.online = allOnline.size;
    summary.totals.peopleToday = allToday.size;
    summary.totals.errors24h = errors24h;
    summary.totals.failing = summary.hubs.filter(h => ["OFFLINE", "PARTIAL", "ERRORS"].includes(h.status)).length;
    return summary;
  };

  // A row from CG_HubSessions / CG_HubErrors (Graph list item) -> plain object
  P.fromSession = (item) => {
    const f = item.fields || {};
    return { app: f.App || "", user: f.UserName || P.who(item.createdBy), started: f.StartedAt ? new Date(f.StartedAt) : null,
             lastSeen: new Date(f.LastSeen || item.lastModifiedDateTime), views: f.PageViews || 0, screens: f.Screens || "",
             errors: f.ErrorCount || 0, throttles: f.Throttles || 0, build: f.Build || "", loadMs: f.LoadMs || null, device: f.Device || "" };
  };
  P.fromError = (item) => {
    const f = item.fields || {};
    return { app: f.App || "", at: new Date(item.createdDateTime), user: f.UserName || P.who(item.createdBy),
             kind: f.Title || "error", message: f.Message || "", detail: f.Detail || "", page: f.Page || "" };
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

  /* -----------------------------------------------------------------
     Reading SharePoint - shared by Mission Control (in the browser, with
     CG.graph) and the alerts job (alerts/run.js, in GitHub Actions).
       graph(method, url, body, headers) -> parsed JSON, or throws an error
                                            carrying .status
       siteId("/sites/X")                -> the Graph site id
     ----------------------------------------------------------------- */
  const PREFER = { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" };
  P.isMissing = e => !!e && (e.status === 404 || /not (be )?found|does not exist/i.test(e.message || ""));
  P.isForbidden = e => !!e && (e.status === 401 || e.status === 403 || /access denied|forbidden|not authorized/i.test(e.message || ""));

  // Runs jobs a few at a time, so SharePoint isn't asked for 20 lists at once.
  P.pool = async (jobs, size) => {
    const out = new Array(jobs.length); let next = 0;
    await Promise.all(Array.from({ length: Math.min(size, jobs.length) }, async () => {
      while (next < jobs.length) { const i = next++; out[i] = await jobs[i](); }
    }));
    return out;
  };

  P.readList = async (graph, siteId, base, list, since) => {
    base = Object.assign({}, base, { list });
    try {
      const page = await graph("GET", P.recentUrl(siteId, list, since), undefined, PREFER);
      const items = ((page && page.value) || []).map(P.fromItem);
      let last = null;
      if (!items.length) {
        const one = await graph("GET", P.latestUrl(siteId, list), undefined, PREFER);
        if (one && one.value && one.value[0]) { const i = P.fromItem(one.value[0]); last = { at: i.at, by: i.by }; }
      }
      return Object.assign(base, { ok: true, items, last, capped: !!(page && page["@odata.nextLink"]) });
    } catch (e) {
      if (P.isMissing(e) || P.isForbidden(e)) return Object.assign(base, { ok: false, error: P.explain(e) });
      // Big lists can't be sorted/filtered without an index: fall back to the
      // list's own "last changed" time so the pulse still shows something.
      try {
        const meta = await graph("GET", `/sites/${siteId}/lists/${encodeURIComponent(list)}?$select=lastModifiedDateTime`);
        return Object.assign(base, { ok: true, items: [], last: { at: new Date(meta.lastModifiedDateTime), by: "Not known (large list)" }, note: P.explain(e) });
      } catch (e2) {
        return Object.assign(base, { ok: false, error: P.explain(e) });
      }
    }
  };

  /* What the Hubs reported themselves (CG_TELEMETRY in cg-config.js).
     Returns { ok: true, sessions, errors } or { ok: false, reason }. */
  P.readTelemetry = async (graph, siteIdOf, T, since) => {
    if (!T || T.enabled === false) return { ok: false, reason: "Telemetry is switched off in cg-config.js" };
    const all = async (url) => {
      const out = [];
      for (let n = 0; url && n < 10; n++) {
        const page = await graph("GET", url, undefined, PREFER);
        ((page && page.value) || []).forEach(i => out.push(i));
        url = (page && page["@odata.nextLink"]) || null;
      }
      return out;
    };
    try {
      const sid = await siteIdOf(T.site);
      const base = list => `/sites/${sid}/lists/${encodeURIComponent(list)}/items?$expand=fields&$top=500`;
      const [ss, es] = await Promise.all([
        all(`${base(T.sessions)}&$filter=fields/LastSeen ge '${since.toISOString()}'`),
        all(`${base(T.errors)}&$filter=fields/Created ge '${since.toISOString()}'`)
      ]);
      return { ok: true, sessions: ss.map(P.fromSession), errors: es.map(P.fromError) };
    } catch (e) {
      return { ok: false, reason: P.isMissing(e) ? "The CG_Platform site or its lists aren't set up yet." : P.explain(e) };
    }
  };

  /* Reads every watched list and the telemetry, and returns the summary
     the screen (or the alert rules) work from.
       opts: { graph, siteId, apps: [CG_APPS entries], monitor: CG_MONITOR,
               telemetry: CG_TELEMETRY, windowDays, now: Date } */
  P.collect = async (opts) => {
    const now = opts.now || new Date();
    const since = new Date(now - (opts.windowDays || 7) * 86400000);
    const appOf = key => (opts.apps || []).find(a => a.key === key);
    const jobs = [];
    for (const h of opts.monitor.hubs) {
      const app = appOf(h.app);
      if (!app || !h.lists.length) continue;
      const sitePath = h.site || app.site;
      let sid = null, siteErr = null;
      try { sid = await opts.siteId(sitePath); } catch (e) { siteErr = P.explain(e); }
      for (const [label, list] of h.lists) {
        const base = { app: h.app, hubName: app.name, label };
        jobs.push(sid ? () => P.readList(opts.graph, sid, base, list, since)
                      : async () => Object.assign({}, base, { list, ok: false, error: `Site ${sitePath}: ${siteErr}` }));
      }
    }
    const [reads, tele] = await Promise.all([P.pool(jobs, 4), P.readTelemetry(opts.graph, opts.siteId, opts.telemetry, since)]);
    const hubs = opts.monitor.hubs.map(h => appOf(h.app)).filter(Boolean).map(a => ({ app: a.key, name: a.name }));
    return P.withTelemetry(P.summarise(reads, now, { hubs }), tele, now);
  };

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
