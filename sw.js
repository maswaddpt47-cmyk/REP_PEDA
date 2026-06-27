const CACHE      = "conum47-v6";
const CACHE_PPTX = "conum47-pptx"; // cache dédié, jamais purgé lors des mises à jour

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
      .then(keys => Promise.all(
        // Conserver CACHE et CACHE_PPTX, purger les anciennes versions
        keys.filter(k => k !== CACHE && k !== CACHE_PPTX).map(k => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", e => {
  if (e.request.url.includes("api.github.com")) return;

  // PPTXs : cache-first dans le cache dédié (mis en cache à la première génération)
  if (e.request.url.includes("/assets/pptx/")) {
    e.respondWith(
      caches.open(CACHE_PPTX).then(async cache => {
        const cached = await cache.match(e.request);
        if (cached) return cached;
        const res = await fetch(e.request);
        if (res.ok) cache.put(e.request, res.clone());
        return res;
      })
    );
    return;
  }

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
