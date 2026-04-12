// Ward Manager Service Worker
// ── Bump this version string on every deploy to force cache refresh ──
const VERSION = 'v6';

const CACHE_NAME = `ward-manager-${VERSION}`;

// On install — skip waiting so the new SW activates immediately
self.addEventListener('install', () => {
  self.skipWaiting();
});

// On activate — delete any old caches from previous versions
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

// On fetch — network first, fall back to cache for navigation requests
// This means users always get the latest version when online
self.addEventListener('fetch', e => {
  if (e.request.mode === 'navigate') {
    e.respondWith(
      fetch(e.request)
        .then(res => {
          // Cache the fresh response for offline fallback
          const clone = res.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(e.request, clone));
          return res;
        })
        .catch(() => caches.match(e.request))
    );
  }
});
