/* =====================================================================
   CG HUBS alerts - the rules (no SharePoint, no Teams: just decisions).

   Used by alerts/run.js in GitHub Actions; tested by tests/alerts.test.js.
   Settings: CG_ALERTS in shared/cg-config.js.

   Flow each run:
     evaluate(summary)  -> the problems that exist right now
     diff(problems, memory) -> what to tell Teams (new / reminder / resolved)
                               and the memory to keep for next run
     card(messages)     -> the Teams message
   ===================================================================== */
"use strict";
const P = require("../control/pulse.js");

const MIN = 60000, HOUR = 60 * MIN;
const R = {};

/* The problems that exist right now.
   Returns [{ key, hub, title, text }] - key identifies a problem across runs. */
R.evaluate = (summary, cfg, now) => {
  const out = [];
  const hp = cfg.hubProblems || {};
  if (hp.enabled !== false) {
    summary.hubs.forEach(h => {
      if (h.status === "OFFLINE") {
        out.push({ key: `hub:${h.app}:offline`, hub: h.name, title: `${h.name} can't be read`,
                   text: `None of its watched lists could be read: ${h.errors.map(e => `${e.label} (${e.message})`).join("; ")}` });
      } else if (h.errors.length) {
        out.push({ key: `hub:${h.app}:lists`, hub: h.name, title: `${h.name}: ${h.errors.length} list${h.errors.length > 1 ? "s" : ""} can't be read`,
                   text: h.errors.map(e => `${e.label} (${e.list}): ${e.message}`).join("; ") });
      }
      const limit = Number(hp.errorsPerHour) || 3;
      const lastHour = (h.userErrors || []).filter(e => now - e.at <= HOUR);
      if (lastHour.length >= limit) {
        const people = Array.from(new Set(lastHour.map(e => e.user))).sort();
        const latest = lastHour[0];
        out.push({ key: `hub:${h.app}:errors`, hub: h.name, title: `${h.name}: ${lastHour.length} errors in the last hour`,
                   text: `People affected: ${people.join(", ")}. Latest (${P.clock(latest.at)}): ${latest.message}` });
      }
    });
  }

  // The ZKTeco bridge, from its own heartbeat (it only writes attendance
  // rows when someone scans, so silence in that list proves nothing).
  // Nothing is raised before the bridge has reported at least once, or
  // while the heartbeat list can't be read - R.bridgeStatus says which.
  const b = cfg.bridge || {};
  if (b.enabled !== false && b.key) {
    const st = P.heartbeatState(summary.heartbeats, b.key, now, b);
    const r = st.row;
    const where = r && r.host ? ` on ${r.host}` : "";
    if (st.state === "down") {
      out.push({ key: "bridge:down", hub: "CG Time", title: `ZKTeco bridge not running (${st.text})`,
                 text: `Last heartbeat ${P.stamp(r.lastSeen)}${where}. Clock-ins aren't reaching CG Time. Check the bridge PC is on and the "CG Time ZKTeco Bridge" task is running.` });
    } else if (st.state === "device") {
      out.push({ key: "bridge:device", hub: "CG Time", title: "ZKTeco device not connected to the bridge",
                 text: `The bridge${where} is running, but ${st.text}. Check the device is powered on and on the network.` });
    } else if (st.state === "sync") {
      out.push({ key: "bridge:sync", hub: "CG Time", title: "ZKTeco bridge can't sync to SharePoint",
                 text: `${r.lastSyncError}${r.lastSyncAt ? ` (last good sync ${P.stamp(r.lastSyncAt)})` : ""}` });
    }
  }
  return out;
};

/* Compares today's problems with what was already said.
   memory: { [key]: { since, lastSent, title } }  (ISO strings, from the last run)
   Returns { messages: [{ type: "new"|"reminder"|"resolved", key, title, text, since }], memory } */
R.diff = (problems, memory, cfg, now) => {
  const remind = (Number(cfg.remindHours) || 6) * HOUR;
  const next = {}, messages = [];
  problems.forEach(p => {
    const was = memory[p.key];
    if (!was) {
      messages.push(Object.assign({ type: "new", since: now }, p));
      next[p.key] = { since: now.toISOString(), lastSent: now.toISOString(), title: p.title };
    } else {
      const since = new Date(was.since);
      if (now - new Date(was.lastSent) >= remind) {
        messages.push(Object.assign({ type: "reminder", since }, p));
        next[p.key] = { since: was.since, lastSent: now.toISOString(), title: p.title };
      } else {
        next[p.key] = { since: was.since, lastSent: was.lastSent, title: p.title };
      }
    }
  });
  Object.keys(memory).forEach(k => {
    if (!next[k]) messages.push({ type: "resolved", key: k, title: memory[k].title, text: "", since: new Date(memory[k].since) });
  });
  return { messages, memory: next };
};

/* The bridge in one short, name-free line (test message and job log). */
R.bridgeStatus = (summary, cfg, now) => {
  const b = cfg.bridge || {};
  if (b.enabled === false) return "bridge check off";
  return "bridge " + P.heartbeatState(summary.heartbeats, b.key, now, b).text;
};

/* Lists read only partly (too big to filter) - the checks that depend on
   them can't see everything, so this is always reported, never silent. */
R.partial = (summary) => {
  const out = [];
  summary.hubs.forEach(h => h.lists.forEach(l => { if (l.ok && l.note) out.push(l.list); }));
  return out;
};

/* The Teams message (an Adaptive Card for a Teams "Workflows" webhook). */
R.card = (messages, now, link, heading) => {
  const icon = { new: "🔴", reminder: "🟠", resolved: "✅", info: "ℹ️" };
  const label = { new: "New", reminder: "Still happening", resolved: "Resolved", info: "" };
  const body = [
    { type: "TextBlock", size: "Medium", weight: "Bolder", text: heading || "CG HUBS Mission Control" },
    { type: "TextBlock", isSubtle: true, spacing: "None", text: `${P.stamp(now)} Port Moresby` }
  ];
  messages.forEach(m => {
    const age = m.since && m.type !== "new" ? ` · since ${P.stamp(m.since)} (${P.ageText(now - m.since)})` : "";
    body.push({ type: "TextBlock", wrap: true, weight: "Bolder", spacing: "Medium",
                text: `${icon[m.type] || ""} ${label[m.type] ? label[m.type] + ": " : ""}${m.title}${age}` });
    if (m.text) body.push({ type: "TextBlock", wrap: true, spacing: "Small", text: m.text });
  });
  return {
    type: "message",
    attachments: [{
      contentType: "application/vnd.microsoft.card.adaptive",
      content: {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json", type: "AdaptiveCard", version: "1.4",
        body,
        actions: link ? [{ type: "Action.OpenUrl", title: "Open Mission Control", url: link }] : []
      }
    }]
  };
};

/* A short, name-free line for the job's log (the repository is public). */
R.logLine = (problems, messages, extra) =>
  `${extra ? extra + "; " : ""}${problems.length} problem(s) now [${problems.map(p => p.key).join(", ")}]; ` +
  `sent ${messages.filter(m => m.type === "new").length} new, ${messages.filter(m => m.type === "reminder").length} reminder, ` +
  `${messages.filter(m => m.type === "resolved").length} resolved`;

module.exports = R;
