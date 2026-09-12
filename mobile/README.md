# Concept Time — Mobile

A phone version of the Concept Time desktop widget: the same sign-in, the
same "pick what you're working on, start/stop a timer or log hours
manually" flow, writing to the exact same SharePoint lists — just as one
mobile-friendly page instead of an Electron window. No "always on top"
(there's no desktop to float over on a phone), and no draggable frameless
window — otherwise, functional parity with the widget.

It's a **Progressive Web App (PWA)**, not a native iOS/Android app. That's
a deliberate choice, not a compromise: it keeps to the exact same
architecture as every other Hub (a static page on GitHub Pages, talking to
SharePoint via Microsoft Graph), so it ships today — no Apple Developer
Program enrollment, no app-store review, no separate native codebase to
maintain. A PWA still installs to the home screen with its own icon and
opens full-screen like an app; see "Adding it to a phone's home screen"
below.

## What's in this folder

```
cg-time-mobile/
  index.html          the entire app — one page, inline styles & script
  manifest.json        PWA metadata (name, icon, colors)
  sw.js                 tiny service worker so the app shell opens instantly
  logo.png              the CG Time logo (same one used on the dashboard)
  background.jpg        the watch/gears background photo (same one used on the dashboard)
  icons/                home-screen icons generated from that logo
README.md               this file
```

Everything reads/writes the same **CG_Time** SharePoint site as the
desktop widget:
- Reads active work from **CSA_ActiveEntities** (same as the widget's
  dropdown — CSAs, Proposals, PSP activities, SLA tickets, whatever is
  flagged active).
- Writes time to **CSA_TimeEntries** with the same fields the widget
  uses (`Title`, `CSA`, `WorkDate`, `TimeStarted`, `TimeStopped`, `Hours`,
  `EntryMethod`, `EntryUser`, `Source`) — entries made from a phone are
  indistinguishable from ones made on the desktop widget in any report
  built on that list. `EntryMethod` stays `"Timer"` / `"Manual"` exactly
  like the widget; if you'd ever want to tell phone and desktop entries
  apart later, that's a one-line change (e.g. `"Timer (Mobile)"`), not a
  redesign.

One small addition beyond strict parity: a **Sign out** link next to the
user's name. The desktop widget doesn't have one since it's one person's
own PC; a phone is more often shared or borrowed, so it seemed worth the
one extra tap.

## Azure AD setup

This app has its own dedicated App Registration, **"CG Time Mobile"**
(client ID `ec5454e0-c794-419b-9ec9-464e9917ef3c`) — matching the pattern
SLA HUB, the ZKTeco Bridge, and the desktop Widget each already follow
(one registration per component), which the strategic review recommended
as the long-term target for the whole platform, rather than the four Hubs
currently sharing one registration. `index.html`'s `CONFIG.clientId` is
already set to this ID — nothing to change there.

Registration details, for reference (already done):
- Platform: **Single-page application**, redirect URI
  `https://marcusappdev.github.io/CG_Training/cg-time-mobile/` (note the
  trailing slash — GitHub Pages serves this folder with one, and the
  registered redirect URI has to match exactly, slash and all).
- API permissions: Microsoft Graph → Delegated → `Sites.ReadWrite.All`
  and `User.Read`, with admin consent granted.

If this page ever gets hosted at a different URL, add that as an
additional redirect URI on the same registration (Authentication → Add a
platform → Single-page application, or just add another redirect URI
under the existing SPA platform) rather than creating a new registration.

## Where to host it

Same pattern as every other Hub: a folder inside the
`marcusappdev/CG_Training` GitHub repository, published by GitHub Pages.

1. Copy the whole `cg-time-mobile/` folder into the repo (alongside
   `CG-Time/`, `csa-hub/`, `psp-hub/`, `sla-hub/`).
2. Commit and push.
3. It'll be live at
   `https://marcusappdev.github.io/CG_Training/cg-time-mobile/` — that's
   the exact URL to use as the redirect URI above. (If you'd rather use a
   different folder name, the URL changes to match — just use whatever
   you actually deploy to as the redirect URI.)

## Adding it to a phone's home screen

- **Android (Chrome):** open the URL, tap the **⋮** menu → **Add to Home
  screen** / **Install app**.
- **iPhone (Safari):** open the URL, tap the **Share** icon → **Add to
  Home Screen**.

Either way it opens full-screen with the CG Time icon, no browser
address bar — the closest a web app gets to feeling like a native one.

## Quick test checklist

- Open the URL, sign in, confirm the active-work dropdown populates.
- Start the timer, wait a few seconds, stop it — confirm a row appears in
  **CSA_TimeEntries** with `EntryMethod = Timer` and a sensible `Hours`
  value (rounded up to the nearest quarter hour, same as the widget).
- Log a manual entry — confirm `EntryMethod = Manual`.
- Tap **Add detail** on a just-logged entry and save a note — confirm
  `ActivityDetail` is set on that row.
- Add to home screen and re-open from there — confirm it opens signed in
  (no repeat login) and full-screen.
