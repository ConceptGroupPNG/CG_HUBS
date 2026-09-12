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

## Azure AD setup — pick one

The app needs an Azure AD App Registration to sign in with. Two ways to
get there; either works, and the code doesn't care which you pick beyond
one config line.

### Option A — reuse the widget's existing registration (least setup)

The desktop widget already has its own registration ("Concept Time
Widget", client ID `f42f07db-5396-44f2-96dc-e2b64b49ee9c`) with exactly
the permissions this page needs (`Sites.ReadWrite.All`, `User.Read`).
`index.html` is already configured to use it — you only need to teach
that same registration about this new page:

1. In the Azure Portal, go to **Entra ID → App registrations** and open
   the one named **Concept Time Widget** (or search the client ID above).
2. Go to **Authentication → Add a platform → Single-page application**
   (this is a different platform type from the widget's existing
   "Mobile and desktop applications" one — you're adding a second one
   alongside it, not replacing it).
3. For the redirect URI, enter the exact URL this page will be hosted at
   once you've deployed it — see "Where to host it" below for what that
   URL will be (e.g.
   `https://marcusappdev.github.io/CG_Training/cg-time-mobile/`).
4. Save. Nothing else to change — no new permissions, no new consent.

### Option B — give this app its own dedicated registration

Matches the pattern SLA HUB, the ZKTeco Bridge, and the Widget each
already follow (one registration per component) — the direction the
strategic review recommended as the long-term target for the whole
platform, rather than the four Hubs currently sharing one registration.

1. **Entra ID → App registrations → New registration.** Name it
   something like "Concept Time Mobile".
2. Platform: **Single-page application**, redirect URI as in Option A,
   step 3.
3. **API permissions → Add a permission → Microsoft Graph → Delegated →**
   add `Sites.ReadWrite.All` and `User.Read`, then **Grant admin consent**.
4. Copy the new **Application (client) ID** and paste it into
   `index.html`, replacing the `clientId` value near the top of the
   `<script>` block (look for `CONFIG.clientId`).

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
