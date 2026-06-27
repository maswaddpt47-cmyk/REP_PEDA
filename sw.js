const CACHE = "conum47-v7";
const STATIC = [
  "./index.html",
  "./test.html",
  "./assets/icon.svg",
  "./assets/slide-plans.json",
];

self.addEventListener("install", e => {
  e.waitUntil(
    caches.open(CACHE)
      .then(c => Promise.allSettled(STATIC.map(u => c.add(u))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", e => {
  if (e.request.url.includes("api.github.com")) return;
  if (e.request.url.includes("/assets/pptx/")) return; // PPTXs : réseau uniquement

  // HTML et données dynamiques : réseau d'abord, cache en fallback offline
  if (e.request.mode === "navigate"
      || e.request.url.includes("links.json")
      || e.request.url.includes("manifest.json")) {
    e.respondWith(
      caches.open(CACHE).then(async cache => {
        try {
          const res = await fetch(e.request);
          if (res.ok) cache.put(e.request, res.clone());
          return res;
        } catch {
          return cache.match(e.request);
        }
      })
    );
    return;
  }

  // Autres assets statiques : cache-first
  e.respondWith(
    caches.open(CACHE).then(async cache => {
      const cached = await cache.match(e.request);
      if (cached) return cached;
      const res = await fetch(e.request);
      if (res.ok) cache.put(e.request, res.clone());
      return res;
    })
  );
});
