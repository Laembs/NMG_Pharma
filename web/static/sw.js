/* NMGone-Web · Service Worker (PWA-Gerüst, P0).
 *
 * Strategie bewusst schlank und sicher:
 *  - Statische Assets (CSS/JS/Icons): cache-first (schneller Start, offline-fähig).
 *  - Navigationen (HTML-Seiten): network-first mit Cache-Fallback, damit man bei
 *    Netzausfall wenigstens die zuletzt gesehene Seite sieht.
 *
 * WICHTIG: Das ist KEIN Offline-Verkauf. Echte Offline-Fähigkeit (lokale
 * Warteschlange + Sync) ist eine eigene Ausbaustufe (siehe Plan_Kasse_Web.pdf).
 * POST-Anfragen werden nie gecacht.
 */
const CACHE = "nmgone-shell-v3";
const SHELL = [
  "/static/app.css",
  "/static/htmx.min.js",
  "/static/icon-192.png",
  "/static/icon-512.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return; // POST/Logins nie cachen

  const url = new URL(req.url);
  if (url.pathname.startsWith("/static/")) {
    // Cache-first für statische Dateien.
    event.respondWith(
      caches.match(req).then((hit) => hit || fetch(req).then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(req, copy));
        return res;
      }))
    );
    return;
  }

  if (req.mode === "navigate") {
    // Network-first für Seiten, Cache als Notnagel bei Netzausfall.
    event.respondWith(
      fetch(req).then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(req, copy));
        return res;
      }).catch(() => caches.match(req).then((hit) => hit || caches.match("/kasse")))
    );
  }
});
