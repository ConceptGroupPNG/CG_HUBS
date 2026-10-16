# CG Time Attendance

A one-page installable web app (PWA) for employees to clock themselves
in and out from their phone — gated by GPS so it only works at a Clock
In/Out location they've been assigned to in CG Time. Separate from
**Concept Time Mobile** (that one logs billable/consulting hours against
CSA/SLA work); this one writes straight into the same
**CG_EmployeeTimeRecords** list the CG Time dashboard already uses for
staff Time & Attendance.

## What's in this folder

```
cg-time-attendance/
  index.html          the entire app — one page, inline styles & script
  manifest.json        PWA metadata (name, icon, colors)
  sw.js                 tiny service worker so the app shell opens instantly
  logo.png              the CG Time logo
  background.png        the watch/gears background photo
  icons/                home-screen icons
README.md               this file
```

## One required change before this works: add Email to CG_Employees

CG_EmployeeTimeRecords today is filled in by a supervisor or the ZKTeco
device — nothing currently links a signed-in Microsoft account to a row
in **CG_Employees**. This app needs that link to know who's clocking in.

1. Open the **CG_Employees** list on the CG_Time SharePoint site.
2. Add a column named **Email** (Single line of text).
3. Fill in each active employee's Concept Group email address (the same
   one they sign in to Microsoft 365 with).

Anyone without an Email value on file will see "Your sign-in isn't
linked to an Employee record yet" instead of the clock in/out screen —
a clear signal for who's still missing, rather than a silent failure.

## What gets written

- **Clock In:** creates a new CG_EmployeeTimeRecords row for that
  Employee + today, with `TimeIn` set to the current time (24hr
  `HH:MM`, matching the dashboard's own format exactly),
  `EntrySource: "Mobile"`, and `PayNumber` looked up from whichever
  CG_PayPeriods period covers today (left blank if none is on file for
  today — doesn't block clocking in).
- **Clock Out:** updates that same day's row with `TimeOut`. If the
  day's `TimeIn` already came from a Manual or Biometric entry (e.g.
  they tapped the office device this morning), Clock Out still works —
  it only ever touches `TimeOut`, leaving `EntrySource` as whatever the
  clock-in already was.
- If a day already has both `TimeIn` and `TimeOut`, the app shows
  "Today's attendance complete" and won't let them clock in again.
- Nothing here overwrites Break Start/End — those still fall back to
  the Employee's/Shift's configured break length exactly like a
  Biometric-only day does today.

This shows up in the CG Time dashboard's Daily Report, Attendance, and
Exceptions tabs exactly like any other row — "Mobile" just becomes a
third value alongside "Manual" and "Biometric" in EntrySource.

## Clock locations (managed entirely from CG Time)

Geofencing is no longer a single hardcoded office point in this app —
it's per-Employee, driven by two lists on the CG_Time SharePoint site
that CG Time's own **Employees → Locations** tab manages:

- **CG_ClockLocations** — one row per named location (e.g. "Head
  Office", "Warehouse"), each with its own `Latitude`, `Longitude`,
  `RadiusMeters`, and `IsActive`.
- **CG_EmployeeLocationAssignments** — a join list saying which
  Employees may use which Locations (an Employee can have none, one,
  or several).

This app only ever **reads** these two lists — adding, editing,
retiring a Location, or changing who's assigned to it all happens in
CG Time, never here. At sign-in, this app looks up the signed-in
Employee's assigned + Active Locations and checks their live GPS
against all of them, accepting Clock In/Out if they're within range of
**any one** of them. If GPS accuracy has someone borderline, adjusting
a Location's `RadiusMeters` in CG Time takes effect immediately next
time they open the app — nothing to change here.

If an Employee has no Locations assigned (or the two lists above
haven't been set up yet on the CG_Time site), the app shows "You're
not assigned to a Clock In/Out location yet" and disables Clock In —
same clear-diagnosis approach as the "not linked to an Employee
record" banner, rather than silently failing.

Location is re-checked at the moment someone taps Clock In/Out, not
just on page load, so a stale "in range" reading from earlier can't be
used to clock in from somewhere else later.

**Worth knowing:** phone GPS can be spoofed with fake-location apps.
This gate is a reasonable deterrent and an audit trail, not a hard
security guarantee — the same trade-off as most mobile geofencing.

## Azure AD setup

This app needs its own dedicated App Registration, same pattern as
Concept Time Mobile:

1. **Entra ID → App registrations → New registration.** Name it
   something like "CG Time Attendance".
2. Platform: **Single-page application**. Redirect URI: the exact URL
   this page will be hosted at (see "Where to host it" below), e.g.
   `https://marcusappdev.github.io/CG_Training/cg-time-attendance/`
   — note the trailing slash, GitHub Pages serves this folder with
   one, and the registered redirect URI has to match exactly.
3. **API permissions → Add a permission → Microsoft Graph → Delegated →**
   add `Sites.ReadWrite.All` and `User.Read`, then **Grant admin consent**.
4. Copy the new **Application (client) ID** and paste it into
   `index.html`, replacing `REPLACE_WITH_YOUR_APP_CLIENT_ID`
   (look for `CONFIG.clientId` near the top of the `<script>` block).

## Where to host it

Same as every other Hub: drop this whole folder into the
`marcusappdev/CG_Training` repo, alongside `cg-time-mobile/` and the
rest, e.g. as `cg-time-attendance/`. It'll publish at
`https://marcusappdev.github.io/CG_Training/cg-time-attendance/` —
that exact URL (with trailing slash) is what goes in the Azure AD
redirect URI above.

## Home screen install

Same as Concept Time Mobile — open the URL in the phone's browser,
then "Add to Home Screen" (Android Chrome: menu → Install app / Add to
Home screen; iOS Safari: Share → Add to Home Screen).

## Quick test checklist

- [ ] Add your own email to your CG_Employees row, assign yourself at
      least one Location in CG Time's Employees → Locations tab, then
      sign in — you should land on the clock in/out screen, not the
      "not linked" banner.
- [ ] From within range of an assigned Location: location pill shows
      green/in-range and names that Location, Clock In button is enabled.
- [ ] Tap Clock In — check CG_EmployeeTimeRecords in SharePoint (or the
      CG Time dashboard's Daily Report tab) for a new row with today's
      date, your Employee Number, `TimeIn` set, `EntrySource: "Mobile"`.
- [ ] Tap Clock Out — same row now has `TimeOut` set too; app shows
      "Today's attendance complete" and disables the button.
- [ ] From well outside every assigned Location's radius (or with
      location permission denied): button stays disabled, message names
      the nearest assigned Location and how far away you are.
- [ ] Remove all of your Location assignments in CG Time, reload: app
      shows "not assigned to a Clock In/Out location" and Clock In stays
      disabled, without ever prompting for GPS.
- [ ] Install to home screen, close the app fully, reopen — still
      signed in, still shows today's correct status.
