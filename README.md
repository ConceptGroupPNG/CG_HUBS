# CG HUBS — Concept Group Hubs

The web apps used across Concept Group, served from
**https://conceptgrouppng.github.io/CG_HUBS/**.

| Folder | App | Business unit |
|---|---|---|
| `/` | CG HUBS launcher | All |
| `training/` | CG Training HUB | Concept Training Services |
| `psp/` | PSP HUB | Payroll Services PNG |
| `csa/` | CSA HUB | Brian Hart Consulting |
| `sla/` | SLA HUB | Internal & client support |
| `time/` | CG Time dashboard | Group time, attendance & payroll |
| `mobile/` | Concept Time Mobile | Phone time logging |
| `attendance/` | CG Time Attendance | Phone clock in / out |
| `shared/` | Shared core library, central settings (`cg-config.js`), shared styles and logos | — |

Apps are brought across from the old `marcusappdev/CG_Training` repository one at a
time, then switched on for everyone on a single switch-over day.

## How changes are made

1. Every change is made on its own branch and arrives as a **pull request**.
2. The pull request is reviewed (**Files changed** tab) and merged into `main`.
3. GitHub Pages publishes `main` automatically within a few minutes.
4. To undo a change, open the merged pull request and click **Revert**.

Nobody uploads files directly to `main`.

## Rules for this repository

- **This repository is public.** Only IDs and addresses belong here — never passwords,
  client secrets, staff data or customer data.
- App settings (tenant, SharePoint sites, client IDs, on/off/preview switches) live in
  one place: `shared/cg-config.js`.
- Dates are built from local calendar parts (YYYY-MM-DD), never `toISOString()`, to avoid
  the UTC+10 shift.

## Related repositories (private)

- `ConceptGroupPNG/concept-time-widget` — Concept Time desktop widget (Electron)
- `ConceptGroupPNG/cg-time-zkteco-bridge` — ZKTeco biometric bridge for CG Time

© Concept Group Limited
