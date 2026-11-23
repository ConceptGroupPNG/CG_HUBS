// Tests for CG Time "Upload Timesheets" (time/timesheet-import.js) - run with:  node --test tests/
// The sheet below copies the layout of the G4S "times PPE" Excel timesheets (raw cell values as
// SheetJS reads them: dates as Excel serials, times as fractions of a day), with made-up people
// (this repo is public).
const { test } = require("node:test");
const assert = require("node:assert/strict");
const I = require("../time/timesheet-import.js");

const HEAD = ["Employee#", "Job#", "Name", "Date", "Time In", "Time Out", "Total Hours", "Break Start", "Break End", "Break Hours", "Hours to pay"];
const T = (h, m) => (h * 60 + m) / 1440;      // Excel time
const D = 46335;                               // Excel serial for 09-11-2026 (a Monday)
const SHEET = [
  HEAD,
  ["CS0001", "03", "Test Person", D,     T(7, 30), T(17, 0),  9.5,  T(12, 0), T(13, 0), 0.9999999999999991, 8.5],
  ["CS0001", "03", "Test Person", D + 3, "",       "",        0,    T(12, 0), T(12, 0), 0, 0],          // didn't work
  ["CS0001", "03", "Test Person", D + 5, "",       "",        0,    "",       "",       0.33, 0],       // weekend
  ["CS0001", "03", "Test Person", D + 8, T(7, 15), T(17, 0),  9.75, T(12, 0), T(12, 0), 0, 9.75],       // no break
  ["CSW0002", "01", "Other Person", D,   T(8, 45), T(17, 0),  8.25, "",       "",       1, 7.25],       // Break Hours only
];

const emp = (num, name, extra) => ({ id: "e-" + num, fields: Object.assign({ EmployeeNumber: num, Title: name, IsActive: true }, extra) });
const PERIOD = { id: "p1", fields: { PayNumber: 24, PPEStart: "2026-11-09", PPEEnd: "2026-11-22" } };
const ctx = (over) => Object.assign({
  employees: [emp("CSW0002", "Other Person")],
  findRecord: () => null,
  findPeriod: d => (d >= "2026-11-09" && d <= "2026-11-22" ? PERIOD : null),
  newEmployees: new Set(),
  overwrite: true,
}, over);

test("dates and times read from Excel values and text", () => {
  assert.equal(I.toDateStr(D), "2026-11-09");
  assert.equal(I.toDateStr("09/11/2026"), "2026-11-09");  // day first
  assert.equal(I.toDateStr("9-11-26"), "2026-11-09");
  assert.equal(I.toDateStr("9 Nov 2026"), "2026-11-09");
  assert.equal(I.toDateStr("Mon 09-Nov-2026"), "2026-11-09");
  assert.equal(I.toDateStr("2026-11-09"), "2026-11-09");
  assert.equal(I.toDateStr("31/02/2026"), "");
  assert.equal(I.toDateStr("soon"), "");
  assert.equal(I.toTimeStr(T(7, 30)), "07:30");
  assert.equal(I.toTimeStr(D + T(17, 15)), "17:15");     // date-time serial keeps the time
  assert.equal(I.toTimeStr(0.34375), "08:15");
  assert.equal(I.toTimeStr("7:30"), "07:30");
  assert.equal(I.toTimeStr("07:30:00"), "07:30");
  assert.equal(I.toTimeStr("5:15 pm"), "17:15");
  assert.equal(I.toTimeStr("12:00 AM"), "00:00");
  assert.equal(I.toTimeStr("0730"), "07:30");
  assert.equal(I.toTimeStr(""), "");
  assert.equal(I.toTimeStr("25:00"), null);
  assert.equal(I.toTimeStr("late"), null);
});

test("headers matched by name, in any order, with a title row above", () => {
  const map = I.mapHeaders(["Name", "EMPLOYEE NO.", "Work Date", "Clock In", "Clock Out"]);
  assert.deepEqual(map, { name: 0, employeeNumber: 1, date: 2, timeIn: 3, timeOut: 4 });
  const p = I.parseSheet([["G4S timesheet - PPE 22/11/26"], [], ["Date", "Employee #", "Time In", "Time Out"], [D, "cs0009", "08:00", "16:00"]]);
  assert.equal(p.rows.length, 1);
  assert.equal(p.rows[0].employeeNumber, "CS0009");
  assert.equal(p.rows[0].timeOut, "16:00");
  const bad = I.parseSheet([["Who", "When"], ["x", "y"]]);
  assert.equal(bad.rows.length, 0);
  assert.match(bad.problems[0], /Employee# and a Date/);
});

test("parses the G4S layout", () => {
  const p = I.parseSheet(SHEET);
  assert.equal(p.problems.length, 0);
  assert.equal(p.rows.length, 5);
  const [r1, off, wkend, nobreak, bh] = p.rows;
  assert.deepEqual([r1.line, r1.employeeNumber, r1.job, r1.date, r1.timeIn, r1.timeOut, r1.breakStart, r1.breakEnd],
    [2, "CS0001", "03", "2026-11-09", "07:30", "17:00", "12:00", "13:00"]);
  assert.equal(r1.sheetHoursToPay, 8.5);
  assert.equal(I.breakOverrideMinutes(r1), "");            // break punches used as-is
  assert.equal(off.hasTimes, false);
  assert.equal(off.breakStart, "");                        // no break stored on a day not worked
  assert.equal(wkend.hasTimes, false);
  assert.equal(I.breakOverrideMinutes(wkend), "");
  assert.equal(nobreak.breakStart, "12:00");
  assert.equal(nobreak.breakEnd, "12:00");                 // 12:00-12:00 = no break
  assert.equal(bh.breakStart, "");
  assert.equal(I.breakOverrideMinutes(bh), "60");          // Break Hours 1 -> 60 min override
});

test("flags unreadable cells and half-filled rows", () => {
  const p = I.parseSheet([HEAD, ["CS0001", "", "", D, "7:30", "", "", "12:00", "", "", ""], ["CS0001", "", "", "someday", "late", "17:00"]]);
  assert.match(p.rows[0].warnings.join(" "), /No Time Out/);
  assert.match(p.rows[0].warnings.join(" "), /Break Start with no Break End/);
  assert.equal(p.rows[0].breakStart, "");
  assert.match(p.rows[1].errors.join(" "), /isn't a date/);
  assert.match(p.rows[1].errors.join(" "), /Time In "late" isn't a time/);
});

test("plan: unknown Employees hold their days back until ticked to add", () => {
  const parsed = I.parseSheet(SHEET);
  let p = I.plan(parsed, ctx());
  assert.deepEqual(p.unknownEmployees, [{ employeeNumber: "CS0001", name: "Test Person", job: "03", days: 2 }]);
  assert.deepEqual(p.rows.map(r => r.action), ["error", "blank", "blank", "error", "create"]);
  assert.match(p.rows[0].errors[0], /Not in CG Time/);
  p = I.plan(parsed, ctx({ newEmployees: new Set(["CS0001"]) }));
  assert.deepEqual(p.rows.map(r => r.action), ["create", "blank", "blank", "create", "create"]);
  assert.deepEqual(p.counts, { create: 3, update: 0, unchanged: 0, "skip-existing": 0, blank: 2, error: 0 });
});

test("plan: days on file are replaced, skipped or left alone", () => {
  const parsed = I.parseSheet(SHEET);
  const onFile = {
    "CSW0002|2026-11-09": { id: "r1", fields: { TimeIn: "08:45", TimeOut: "17:00", BreakStart: "", BreakEnd: "", BreakMinutesOverride: 60, EntrySource: "Manual" } },
  };
  const employees = [emp("CS0001", "Test Person"), emp("CSW0002", "Other Person")];
  const findRecord = (n, d) => onFile[n + "|" + d] || null;
  let p = I.plan(parsed, ctx({ employees, findRecord }));
  assert.equal(p.rows[4].action, "unchanged");             // identical -> nothing written
  onFile["CS0001|2026-11-09"] = { id: "r2", fields: { TimeIn: "07:31", TimeOut: "16:58", EntrySource: "Biometric" } };
  p = I.plan(parsed, ctx({ employees, findRecord }));
  assert.equal(p.rows[0].action, "update");
  assert.match(p.rows[0].warnings.join(" "), /Replaces the Biometric record/);
  p = I.plan(parsed, ctx({ employees, findRecord, overwrite: false }));
  assert.equal(p.rows[0].action, "skip-existing");
});

test("plan: no Pay Period, inactive Employee, duplicates and name differences", () => {
  const parsed = I.parseSheet([HEAD,
    ["CSW0002", "", "Other Persn", D, T(8, 0), T(16, 0)],
    ["CSW0002", "", "Other Person", D, T(9, 0), T(16, 0)],
    ["CSW0002", "", "Other Person", D + 30, T(8, 0), T(16, 0)],
    ["CS0003", "", "Gone Person", D, T(8, 0), T(16, 0)],
  ]);
  const p = I.plan(parsed, ctx({ employees: [emp("CSW0002", "Other Person"), emp("CS0003", "Gone Person", { IsActive: false })] }));
  assert.equal(p.rows[0].action, "create");
  assert.match(p.rows[0].warnings[0], /Name on the sheet is "Other Persn"/);
  assert.equal(p.rows[1].action, "error");
  assert.match(p.rows[1].errors[0], /more than once/);
  assert.match(p.rows[2].errors[0], /No Pay Period/);
  assert.match(p.rows[3].errors[0], /inactive/);
});
