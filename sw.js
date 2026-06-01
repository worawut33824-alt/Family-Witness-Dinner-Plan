// Wedding Planner – Service Worker
// Caches core assets for offline use

const CACHE = 'wedding-planner-v26';
const ASSETS = [
  './',
  './index.html',
  './invite.html',
  './logo.png',
  './Gif1.png',
  './icon-192.png',
  './icon-512.png',
  './apple-touch-icon.png',
  'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js'
];

// ── Install: cache all core assets ──────────────────────────────
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// ── Activate: remove old caches ─────────────────────────────────
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// ── Fetch: cache-first for assets, network-first for API calls ──
self.addEventListener('fetch', e => {
  const url = e.request.url;

  // Let Google Apps Script calls go straight to network
  if (url.includes('script.google.com')) {
    e.respondWith(fetch(e.request).catch(() => new Response('offline', { status: 503 })));
    return;
  }

  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(response => {
        // Cache successful GET responses for html/js/css/images
        if (
          e.request.method === 'GET' &&
          response.status === 200 &&
          (url.endsWith('.html') || url.endsWith('.js') || url.endsWith('.css') ||
           url.endsWith('.png') || url.endsWith('.jpg') || url.endsWith('.webp'))
        ) {
          const clone = response.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        return response;
      }).catch(() => caches.match('./index.html'));
    })
  );
});
