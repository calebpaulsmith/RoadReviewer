// Offline cache for the Cat C inspection filler. The app is one
// self-contained HTML file, so caching the page (plus the manifest and
// icons) is a complete offline copy. Stale-while-revalidate: serve from
// cache instantly, refresh the cache from the network in the background
// so a redeploy is picked up on the next visit.
const CACHE = "catc-filler-v1";
self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(["./", "manifest.webmanifest", "icon-192.png"])));
  self.skipWaiting();
});
self.addEventListener("activate", (e) => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET") return;
  e.respondWith(
    caches.match(e.request).then((cached) => {
      const refresh = fetch(e.request)
        .then((res) => {
          if (res.ok) caches.open(CACHE).then((c) => c.put(e.request, res.clone()));
          return res;
        })
        .catch(() => cached);
      return cached || refresh;
    })
  );
});
