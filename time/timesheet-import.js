/* =====================================================================
   CG TIME - Upload Timesheets (Excel)
   ---------------------------------------------------------------------
   Turns the rows of an Excel timesheet (one row per Employee per day,
   e.g. the G4S "times PPE" sheets) into CG_EmployeeTimeRecords rows,
   and works out what an import would do before anything is written.
   Nothing here talks to SharePoint: the page reads the file (SheetJS),
   hands the rows here, shows the plan, and saves only on Import.

   Works in the browser (window.TimesheetImport) and in Node
   (module.exports) so it can be tested without a browser.

   Expected columns (matched by header text, any order, case and
   spacing ignored; a few common alternatives are accepted):
     Employee#  Job#  Name  Date  Time In  Time Out  Total Hours
     Break Start  Break End  Break Hours  Hours to pay
   Only Employee# and Date are required; Time In/Out are what make a
   day count. Total Hours / Hours to pay are read only to compare with
   CG Time's own figure - CG Time always works hours out itself.
   ===================================================================== */
(function (global) {
  "use strict";

  // Header text -> field. Headers are normalised by normHeader first.
  const COLUMNS = {
    employeeNumber: ["employee#", "employee no", "employee number", "emp#", "emp no", "employee id", "staff#", "staff no"],
    job: ["job#", "job no", "job", "job number"],
    name: ["name", "employee name", "employee"],
    date: ["date", "work date", "day"],
    timeIn: ["time in", "clock in", "in"],
    timeOut: ["time out", "clock out", "out"],
    totalHours: ["total hours", "total hrs", "total"],
    breakStart: ["break start", "break out"],
    breakEnd: ["break end", "break in"],
    breakHours: ["break hours", "break hrs", "break"],
    hoursToPay: ["hours to pay", "hours to be paid", "paid hours", "pay hours"],
  };
  const REQUIRED = ["employeeNumber", "date"];
  const LABELS = {
    employeeNumber: "Employee#", job: "Job#", name: "Name", date: "Date", timeIn: "Time In",
    timeOut: "Time Out", totalHours: "Total Hours", breakStart: "Break Start", breakEnd: "Break End",
    breakHours: "Break Hours", hoursToPay: "Hours to pay",
  };

  // "Employee #" / "employee no." / "EMPLOYEE#" -> "employeeno"
  function normHeader(s) {
    return String(s == null ? "" : s).toLowerCase().replace(/#/g, "no").replace(/number/g, "no").replace(/[^a-z0-9]/g, "");
  }
  const ALIAS = {};
  Object.keys(COLUMNS).forEach(k => COLUMNS[k].forEach(a => { ALIAS[normHeader(a)] = k; }));

  // Maps a header row to { field: columnIndex }.
  function mapHeaders(headerRow) {
    const map = {};
    (headerRow || []).forEach((h, i) => {
      const key = ALIAS[normHeader(h)];
      if (key && map[key] == null) map[key] = i;
    });
    return map;
  }

  const pad2 = n => String(n).padStart(2, "0");

  /* ---------------- dates -> "YYYY-MM-DD" ---------------- */
  // Excel day serial (days since 30-12-1899). Worked out in UTC on purpose -
  // the serial is a calendar day, so no time zone may shift it.
  function serialToDateStr(serial) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(serial + 1e-7) * 86400000);
    return d.getUTCFullYear() + "-" + pad2(d.getUTCMonth() + 1) + "-" + pad2(d.getUTCDate());
  }
  const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  function validYmd(y, m, d) {
    if (y < 100) y += 2000;
    if (m < 1 || m > 12 || d < 1 || d > 31) return "";
    const dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCMonth() !== m - 1) return ""; // e.g. 31-02
    return y + "-" + pad2(m) + "-" + pad2(d);
  }
  // Accepts an Excel serial, a Date, "2026-11-09", or day-first text
  // ("09/11/2026", "9-11-26", "9 Nov 2026", "09-Nov-2026") - PNG writes
  // dates day first, so "09/11/2026" is 9 November, never 11 September.
  function toDateStr(v) {
    if (v == null || v === "") return "";
    if (typeof v === "number" && isFinite(v)) return v > 0 ? serialToDateStr(v) : "";
    if (v instanceof Date && !isNaN(v)) return v.getFullYear() + "-" + pad2(v.getMonth() + 1) + "-" + pad2(v.getDate());
    const s = String(v).trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return validYmd(+m[1], +m[2], +m[3]);
    m = s.match(/^(\d{1,2})[\/\-. ](\d{1,2})[\/\-. ](\d{2,4})$/);
    if (m) return validYmd(+m[3], +m[2], +m[1]);
    m = s.match(/^(?:[a-z]+,?\s+)?(\d{1,2})[\s\-\/]+([a-z]{3})[a-z]*[\s\-\/,]+(\d{2,4})$/i);
    if (m) {
      const mi = MONTHS.indexOf(m[2].toLowerCase());
      if (mi >= 0) return validYmd(+m[3], mi + 1, +m[1]);
    }
    if (/^\d+(\.\d+)?$/.test(s)) return serialToDateStr(parseFloat(s));
    return "";
  }

  /* ---------------- times -> "HH:MM" (24-hour) ---------------- */
  // Accepts an Excel time (fraction of a day, or a date-time serial - only
  // the time part is kept), a Date, or text "7:30", "07:30:00", "7:30 PM",
  // "0730". Returns "" for blank, null for something unreadable.
  function toTimeStr(v) {
    if (v == null || v === "") return "";
    if (v instanceof Date && !isNaN(v)) return pad2(v.getHours()) + ":" + pad2(v.getMinutes());
    if (typeof v === "number" && isFinite(v)) {
      const frac = v - Math.floor(v);
      let mins = Math.round(frac * 24 * 60);
      if (mins >= 24 * 60) mins = 0;
      return pad2(Math.floor(mins / 60)) + ":" + pad2(mins % 60);
    }
    const s = String(v).trim().toLowerCase();
    if (!s) return "";
    let m = s.match(/^(\d{1,2})(?::|\.)(\d{2})(?::\d{2})?\s*(am|pm|a\.m\.|p\.m\.)?$/) || s.match(/^(\d{1,2})(\d{2})\s*(am|pm)?$/);
    if (m) {
      let h = +m[1]; const mi = +m[2]; const ap = (m[3] || "").replace(/\./g, "");
      if (ap === "pm" && h < 12) h += 12;
      if (ap === "am" && h === 12) h = 0;
      if (h > 23 || mi > 59) return null;
      return pad2(h) + ":" + pad2(mi);
    }
    if (/^0?\.\d+$/.test(s)) return toTimeStr(parseFloat(s));
    return null;
  }

  function toNumber(v) {
    if (v == null || v === "") return null;
    const n = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, ""));
    return isFinite(n) ? n : null;
  }

  // Excel shows "03" for a Job# typed as text but SheetJS may hand back 3.
  const cleanText = v => (v == null ? "" : String(v).trim());

  /* ---------------- sheet rows -> timesheet rows ---------------- */
  // aoa: the sheet as an array of rows (SheetJS sheet_to_json header:1, raw).
  // Returns { columns, missing, rows, problems }.
  //   rows:     [{ line, employeeNumber, name, job, date, timeIn, timeOut,
  //                breakStart, breakEnd, breakHours, sheetTotalHours,
  //                sheetHoursToPay, hasTimes, warnings[], errors[] }]
  //   problems: sheet-level messages (no header row, etc.)
  function parseSheet(aoa) {
    const out = { columns: {}, missing: [], rows: [], problems: [] };
    const rows = Array.isArray(aoa) ? aoa : [];
    // The header is the first row (of the first 15) that names both an
    // Employee# and a Date column - allows a title row or two above it.
    let headerIdx = -1;
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
      const map = mapHeaders(rows[i]);
      if (map.employeeNumber != null && map.date != null) { headerIdx = i; out.columns = map; break; }
    }
    if (headerIdx < 0) {
      out.missing = REQUIRED.map(k => LABELS[k]);
      out.problems.push("Couldn't find a header row with both an Employee# and a Date column.");
      return out;
    }
    const c = out.columns;
    const cell = (r, k) => (c[k] == null ? "" : r[c[k]]);
    for (let i = headerIdx + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const empRaw = cleanText(cell(r, "employeeNumber"));
      const dateRaw = cell(r, "date");
      if (!empRaw && (dateRaw == null || dateRaw === "")) continue; // blank / spacer row
      const row = {
        line: i + 1, // spreadsheet row number, as Excel shows it
        employeeNumber: empRaw.toUpperCase(),
        name: cleanText(cell(r, "name")),
        job: cleanText(cell(r, "job")),
        date: toDateStr(dateRaw),
        timeIn: "", timeOut: "", breakStart: "", breakEnd: "",
        breakHours: toNumber(cell(r, "breakHours")),
        sheetTotalHours: toNumber(cell(r, "totalHours")),
        sheetHoursToPay: toNumber(cell(r, "hoursToPay")),
        hasTimes: false, warnings: [], errors: [],
      };
      if (!empRaw) row.errors.push("No Employee#.");
      if (!row.date) row.errors.push(`Date "${cleanText(dateRaw)}" isn't a date.`);
      ["timeIn", "timeOut", "breakStart", "breakEnd"].forEach(k => {
        const t = toTimeStr(cell(r, k));
        if (t === null) { row.errors.push(`${LABELS[k]} "${cleanText(cell(r, k))}" isn't a time.`); row[k] = ""; }
        else row[k] = t;
      });
      row.hasTimes = !!(row.timeIn || row.timeOut);
      if (row.hasTimes && !(row.timeIn && row.timeOut)) {
        row.warnings.push(row.timeIn ? "No Time Out." : "No Time In.");
      }
      // A day nobody worked carries no break (sheets often leave a default
      // break against weekends) - ignore it rather than storing it.
      if (!row.hasTimes) { row.breakStart = ""; row.breakEnd = ""; }
      else if (!!row.breakStart !== !!row.breakEnd) {
        row.warnings.push(row.breakStart ? "Break Start with no Break End - the sheet's Break Hours are used." : "Break End with no Break Start - the sheet's Break Hours are used.");
        row.breakStart = ""; row.breakEnd = "";
      }
      out.rows.push(row);
    }
    if (!out.rows.length) out.problems.push("No timesheet rows found under the header.");
    return out;
  }

  // Minutes of break to store as an override when the sheet gives no
  // Break Start/End but does give Break Hours, so CG Time pays what the
  // sheet pays instead of falling back to the 20-minute default.
  // "" = no override (break punches, or nothing given at all).
  function breakOverrideMinutes(row) {
    if (!row.hasTimes) return "";
    if (row.breakStart && row.breakEnd) return "";
    if (row.breakHours == null || row.breakHours < 0) return "";
    return String(Math.round(row.breakHours * 60));
  }

  /* ---------------- what Import would do ---------------- */
  // ctx: {
  //   employees:  CG_Employees items ({ id, fields })
  //   findRecord: (employeeNumber, dateStr) -> existing time record item or null
  //   findPeriod: dateStr -> CG_PayPeriods item or null
  //   newEmployees: Set of upper-case Employee#s that will be created on Import
  //   overwrite:  true to replace days already on file
  // }
  // Each row gets .employee, .period, .existing, .override and an .action:
  //   "create" | "update" | "unchanged" | "skip-existing" | "blank" | "error"
  // plus any extra warnings/errors. Returns { rows, unknownEmployees, counts }.
  function plan(parsed, ctx) {
    const byNum = new Map();
    (ctx.employees || []).forEach(e => {
      const n = cleanText(e.fields && e.fields.EmployeeNumber).toUpperCase();
      if (n && !byNum.has(n)) byNum.set(n, e);
    });
    const willCreate = ctx.newEmployees || new Set();
    const unknown = new Map(); // Employee# -> { employeeNumber, name, job, days }
    const seen = new Set();
    const counts = { create: 0, update: 0, unchanged: 0, "skip-existing": 0, blank: 0, error: 0 };

    const rows = parsed.rows.map(src => {
      const row = Object.assign({}, src, { warnings: src.warnings.slice(), errors: src.errors.slice() });
      row.employee = byNum.get(row.employeeNumber) || null;
      row.period = row.date ? (ctx.findPeriod(row.date) || null) : null;
      row.existing = (row.employeeNumber && row.date) ? (ctx.findRecord(row.employeeNumber, row.date) || null) : null;
      row.override = breakOverrideMinutes(row);

      if (row.employeeNumber && !row.employee) {
        const u = unknown.get(row.employeeNumber) || { employeeNumber: row.employeeNumber, name: row.name, job: row.job, days: 0 };
        if (!u.name && row.name) u.name = row.name;
        if (!u.job && row.job) u.job = row.job;
        if (row.hasTimes) u.days++;
        unknown.set(row.employeeNumber, u);
      }

      if (row.errors.length) { row.action = "error"; counts.error++; return row; }

      const key = row.employeeNumber + "|" + row.date;
      if (seen.has(key)) {
        row.errors.push("This Employee and date appear more than once in the sheet - only the first is imported.");
        row.action = "error"; counts.error++; return row;
      }
      seen.add(key);

      if (!row.hasTimes) { row.action = "blank"; counts.blank++; return row; }

      if (!row.employee && !willCreate.has(row.employeeNumber)) {
        row.errors.push("Not in CG Time - tick Add for this Employee above to import their days.");
      } else if (row.employee && row.employee.fields.IsActive === false) {
        row.errors.push("This Employee is inactive in CG Time - reactivate them on the Employees tab first.");
      }
      if (!row.period) row.errors.push("No Pay Period covers this date - add it on the Calendar tab first.");
      if (row.errors.length) { row.action = "error"; counts.error++; return row; }

      if (row.employee && row.name && row.employee.fields.Title &&
          normName(row.name) !== normName(row.employee.fields.Title)) {
        row.warnings.push(`Name on the sheet is "${row.name}" but CG Time has "${row.employee.fields.Title}".`);
      }

      if (row.existing) {
        const f = row.existing.fields || {};
        if (sameAsRecord(row, f)) { row.action = "unchanged"; counts.unchanged++; return row; }
        if (!ctx.overwrite) { row.action = "skip-existing"; counts["skip-existing"]++; return row; }
        const src = f.EntrySource || "Manual";
        if (src !== "Manual") row.warnings.push(`Replaces the ${src} record on file (${f.TimeIn || "—"} to ${f.TimeOut || "—"}).`);
        row.action = "update"; counts.update++; return row;
      }
      row.action = "create"; counts.create++; return row;
    });

    return { rows, unknownEmployees: Array.from(unknown.values()), counts };
  }

  function normName(s) { return String(s || "").toLowerCase().replace(/[^a-z]/g, ""); }

  function sameAsRecord(row, f) {
    const ov = f.BreakMinutesOverride == null || f.BreakMinutesOverride === "" ? "" : String(Math.round(parseFloat(f.BreakMinutesOverride)));
    return (f.TimeIn || "") === row.timeIn && (f.TimeOut || "") === row.timeOut &&
      (f.BreakStart || "") === row.breakStart && (f.BreakEnd || "") === row.breakEnd &&
      ov === row.override && !f.ShiftName;
  }

  // Columns for the blank template the page offers for download.
  const TEMPLATE_HEADERS = ["Employee#", "Job#", "Name", "Date", "Time In", "Time Out", "Total Hours", "Break Start", "Break End", "Break Hours", "Hours to pay"];

  const api = { parseSheet, plan, mapHeaders, normHeader, toDateStr, toTimeStr, breakOverrideMinutes, TEMPLATE_HEADERS, LABELS };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else global.TimesheetImport = api;
})(typeof window !== "undefined" ? window : globalThis);
