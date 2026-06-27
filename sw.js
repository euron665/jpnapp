const CACHE = 'koku-v64';
const ASSETS = [
  '/jpnapp/',
  '/jpnapp/index.html',
  '/jpnapp/manifest.json',
  '/jpnapp/icon-192.png',
  '/jpnapp/icon-512.png'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
  ));
  self.clients.claim();
});

// NETWORK-FIRST: always fetch the latest when online, fall back to cache offline.
// This guarantees the app updates the moment you're online — no stale cache.
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  e.respondWith(
    fetch(req).then(res => {
      if (res && (res.ok || res.type === 'opaque')) {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(req, copy));
      }
      return res;
    }).catch(() =>
      caches.match(req).then(c => c || (req.mode === 'navigate' ? caches.match('/jpnapp/index.html') : Response.error()))
    )
  );
});
