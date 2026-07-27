// service-worker.js
const CACHE_NAME = 'manifest-track-v1';
const urlsToCache = [
  '/offline.html',
  '/static/manifest.json',
  // Add any other static assets you may have (e.g., custom CSS/JS files)
  // The HTML pages are dynamic – we'll cache them on first visit via fetch handler
];

// Pages whose content depends on the current session (auth state, error
// messages, etc.) – never cache these navigation responses, so an offline
// visit never accidentally serves a stale or wrongly-authenticated shell.
const SESSION_PATHS = ['/', '/login', '/logout'];

// Install event – cache core assets
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(urlsToCache))
      .then(() => self.skipWaiting())
  );
});

// Fetch event – serve from cache, fallback to network
self.addEventListener('fetch', event => {
  const requestUrl = new URL(event.request.url);

  // For navigation requests (HTML pages), try network first, then cache
  if (event.request.mode === 'navigate') {
    const isSessionPath = SESSION_PATHS.includes(requestUrl.pathname);
    event.respondWith(
      fetch(event.request)
        .then(response => {
          if (!isSessionPath) {
            // Clone the response and store in cache for next time
            const responseClone = response.clone();
            caches.open(CACHE_NAME).then(cache => {
              cache.put(event.request, responseClone);
            });
          }
          return response;
        })
        .catch(() => {
          if (isSessionPath) {
            // Session-dependent pages never come from cache - offline means
            // the offline page, not a stale login/app shell.
            return caches.match('/offline.html');
          }
          return caches.match(event.request).then(cached => {
            if (cached) return cached;
            // Offline fallback – return a simple offline page
            return caches.match('/offline.html');
          });
        })
    );
    return;
  }

  // For other requests (CSS, JS, images, API), try cache first, then network
  event.respondWith(
    caches.match(event.request)
      .then(response => {
        if (response) return response;
        return fetch(event.request).then(networkResponse => {
          // Never cache dynamic/data endpoints - always hit the network for these,
          // so downloads, generated files, and session data are never served stale.
          const noCachePaths = [
            '/tracking/', '/upload_and_prepare', '/preview_bag_parts',
            '/generate_manifest', '/save_wtbox', '/suggest_bag_markings',
            '/last_manifest_number', '/last_box_limit', '/download/', '/chamber_certificate/',
            '/download_all/'
          ];
          const shouldCache = !noCachePaths.some(p => event.request.url.includes(p));
          if (shouldCache) {
            const responseClone = networkResponse.clone();
            caches.open(CACHE_NAME).then(cache => {
              cache.put(event.request, responseClone);
            });
          }
          return networkResponse;
        });
      })
  );
});

// Activate – clean up old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cache => {
          if (cache !== CACHE_NAME) {
            return caches.delete(cache);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});