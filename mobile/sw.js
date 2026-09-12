// Concept Time — minimal app-shell service worker.
// Only caches this app's own files (so it opens instantly and is
// installable). Never touches Microsoft Graph, sign-in, or the MSAL CDN
// script — those always go straight to the network.
// Bump this version string any time index.html (or any other shell file)
// changes — that's what makes the browser notice this file differs, install
// the new service worker, and evict the old cached copy. Without a bump,
// phones that already visited keep serving the stale cached index.html
// indefinitely, even after the repo itself is updated.
const CACHE_NAME = "concept-time-shell-v4";
const SHELL_FILES = [
  "./",
  "./index.html",
  "./manifest.json",
  "./logo.png",
  "./background.jpg",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/icon-512-maskable.png",
  "./icons/apple-touch-icon.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return; // don't intercept POST/PATCH to Graph
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return; // never intercept Graph/login/CDN

  event.respondWith(
    caches.match(req).then((cached) => {
      const network = fetch(req)
        .then((res) => {
          if (res && res.ok) caches.open(CACHE_NAME).then((c) => c.put(req, res.clone()));
          return res;
        })
        .catch(() => cached);
      return cached || network;
    })
  );
});
