/* =====================================================================
   CG HUBS alerts - the scheduled job (run by .github/workflows/alerts.yml)

     node alerts/run.js

   Environment (GitHub secrets / workflow):
     ALERTS_CLIENT_SECRET  secret of the "CG HUBS Alerts" Entra app
     TEAMS_WEBHOOK_URL     the Teams channel's Workflows webhook address
     SEND_TEST=true        also send a "connected" message with what it can see
     STATE_FILE            where the alert memory is kept (default .alerts-state.json)
     TZ                    must be Pacific/Port_Moresby (set by the workflow)

   THIS REPOSITORY IS PUBLIC and so are the job's logs: nothing printed here
   may contain a name, an error message from a Hub, or a secret. Details
   only ever go to Teams.
   ===================================================================== */
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const P = require("../control/pulse.js");
const R = require("./rules.js");

const GRAPH = "https://graph.microsoft.com/v1.0";

function loadConfig() {
  const ctx = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname, "../shared/cg-config.js"), "utf8") +
    "\nthis.c = { CG_TENANT, CG_APPS, CG_MONITOR, CG_TELEMETRY, CG_ALERTS };", ctx);
  return ctx.c;
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

// App-only sign-in (client credentials). Returns an access token for Graph.
async function getToken(tenantId, clientId, secret, fetchFn = fetch) {
  const res = await fetchFn(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: clientId, client_secret: secret, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" }).toString()
  });
  const j = await res.json().catch(() => ({}));
  if (!res.ok || !j.access_token) {
    const err = new Error(`Sign-in failed (${res.status} ${j.error || ""})`);   // error code only - no detail text
    err.status = res.status;
    throw err;
  }
  return j.access_token;
}

function makeGraph(token, fetchFn = fetch) {
  return async function graph(method, url, body, headers) {
    const full = /^https?:/i.test(url) ? url : GRAPH + url;
    for (let attempt = 0; ; attempt++) {
      const res = await fetchFn(full, {
        method,
        headers: Object.assign({ Authorization: "Bearer " + token }, body ? { "Content-Type": "application/json" } : {}, headers || {}),
        body: body ? JSON.stringify(body) : undefined
      });
      if ((res.status === 429 || res.status === 503) && attempt < 3) {
        const after = parseInt((res.headers && res.headers.get && res.headers.get("Retry-After")) || "", 10);
        await sleep((after > 0 ? Math.min(after, 30) : 2 ** (attempt + 1)) * 1000);
        continue;
      }
      if (!res.ok) {
        let msg = `${res.status}`, code = "";
        try { const j = await res.json(); if (j && j.error) { msg = j.error.message || msg; code = j.error.code || ""; } } catch (e) { /* not JSON */ }
        const err = new Error(msg);
        err.status = res.status; err.code = code;
        throw err;
      }
      return res.status === 204 ? null : res.json();
    }
  };
}

async function postToTeams(url, payload, fetchFn = fetch) {
  const res = await fetchFn(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
  if (!res.ok) throw new Error(`Teams didn't accept the message (${res.status})`);
}

function readMemory(file) {
  try { return JSON.parse(fs.readFileSync(file, "utf8")); } catch (e) { return {}; }
}

/* The whole run. Returns a short name-free status line; throws only on a
   problem with the job itself (so the GitHub run shows red). */
async function run(env = process.env, deps = {}) {
  const fetchFn = deps.fetch || fetch;
  const now = deps.now || new Date();
  const c = deps.config || loadConfig();
  const A = c.CG_ALERTS || {};
  const stateFile = env.STATE_FILE || ".alerts-state.json";
  const memory = readMemory(stateFile);
  const save = m => fs.writeFileSync(stateFile, JSON.stringify(m, null, 1));
  const link = (c.CG_APPS.find(a => a.key === "control") ? "https://conceptgrouppng.github.io/CG_HUBS/control/" : "");

  if (A.enabled === false) { save(memory); return "Alerts are switched off in cg-config.js - nothing checked."; }
  if (!A.clientId || String(A.clientId).startsWith("<<") || !env.ALERTS_CLIENT_SECRET || !env.TEAMS_WEBHOOK_URL) {
    save(memory);
    return "Alerts aren't set up yet (client ID, ALERTS_CLIENT_SECRET or TEAMS_WEBHOOK_URL missing) - nothing checked.";
  }

  // If the job can't sign in or read anything, that is itself an alert.
  let problems, summary = null;
  try {
    const token = await getToken(c.CG_TENANT.TENANT_ID, A.clientId, env.ALERTS_CLIENT_SECRET, fetchFn);
    const graph = makeGraph(token, fetchFn);
    const ids = {};
    const siteId = async p => (ids[p] = ids[p] || (await graph("GET", `/sites/${c.CG_TENANT.SP_HOST}:${p}`)).id);
    summary = await P.collect({ graph, siteId, apps: c.CG_APPS, monitor: c.CG_MONITOR, telemetry: c.CG_TELEMETRY, windowDays: 1, now });
    problems = R.evaluate(summary, A, now);
  } catch (e) {
    problems = [{ key: "alerts:job", hub: "", title: "The alerts job can't read SharePoint",
                  text: `${e.message}. Check the CG HUBS Alerts app's secret (it may have expired) and its site access.` }];
  }

  const { messages, memory: next } = R.diff(problems, memory, A, now);
  const toSend = messages.slice();
  if (String(env.SEND_TEST).toLowerCase() === "true") {
    const t = summary ? summary.totals : null;
    const seen = summary
      ? `Can read ${summary.hubs.filter(h => h.lists.some(l => l.ok)).length} Hubs; telemetry ${summary.telemetry && summary.telemetry.ok ? "on" : "NOT readable"}; ` +
        `today ${t.peopleToday} people, ${t.writesToday} writes, ${t.errors24h || 0} errors in 24h. ` +
        `Problems right now: ${problems.length ? problems.map(p => p.title).join("; ") : "none"}.`
      : "Couldn't read SharePoint - see the message below.";
    toSend.unshift({ type: "info", title: "Alerts are connected", text: seen });
  }
  if (toSend.length) await postToTeams(env.TEAMS_WEBHOOK_URL, R.card(toSend, now, link), fetchFn);
  save(next);
  const line = R.logLine(problems, messages);
  if (problems.some(p => p.key === "alerts:job")) { const err = new Error("Job couldn't read SharePoint - " + line); err.line = line; throw err; }
  return line;
}

module.exports = { run, getToken, makeGraph, loadConfig };

if (require.main === module) {
  run().then(line => console.log(line), e => { console.error(e.line ? e.message : "Alerts job failed: " + (e.status || "") + " " + (e.code || "")); process.exit(1); });
}
