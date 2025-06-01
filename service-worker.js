/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

const CACHE_NAME = 'hitech-po-cache-v1';
const urlsToCache = [
  './', // Alias for index.html when served from root
  './index.html',
  './index.css',
  './index.tsx', // The browser fetches this as a module
  './manifest.json'
  // Add any other static assets like images or fonts if they exist
  // For the PWA icons, they are data URIs and defined in manifest.json,
  // so they don't need to be explicitly listed here for caching by the service worker itself,
  // though the browser will cache them as part of its normal PWA asset handling.
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        console.log('Opened cache');
        // Use a try-catch for addAll as it rejects if any single resource fails
        return Promise.all(
          urlsToCache.map(urlToCache => {
            return cache.add(urlToCache).catch(reason => {
              console.warn(`Failed to cache ${urlToCache}: ${reason}`);
            });
          })
        );
      })
      .then(() => self.skipWaiting()) // Activate the new service worker immediately
      .catch(err => console.error('Cache open/addAll failed during install:', err))
  );
});

self.addEventListener('activate', (event) => {
  const cacheWhitelist = [CACHE_NAME];
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheWhitelist.indexOf(cacheName) === -1) {
            console.log('Deleting old cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    })
    .then(() => self.clients.claim()) // Take control of all open clients
    .catch(err => console.error('Cache cleanup/claim failed during activate:', err))
  );
});

self.addEventListener('fetch', (event) => {
  // Only handle GET requests
  if (event.request.method !== 'GET') {
    return;
  }

  event.respondWith(
    caches.match(event.request)
      .then((response) => {
        // Cache hit - return response
        if (response) {
          return response;
        }

        // Not in cache - fetch from network
        return fetch(event.request).then(
          (networkResponse) => {
            // Check if we received a valid response
            // For 'basic' type (same-origin), status 200 is good.
            // For 'opaque' type (cross-origin, no CORS), we can't inspect status, so cache cautiously.
            // Here, we primarily cache same-origin 'basic' resources.
            if (!networkResponse || (networkResponse.type === 'basic' && networkResponse.status !== 200)) {
              return networkResponse;
            }

            // IMPORTANT: Clone the response. A response is a stream
            // and because we want the browser to consume the response
            // as well as the cache consuming the response, we need
            // to clone it so we have two streams.
            const responseToCache = networkResponse.clone();

            caches.open(CACHE_NAME)
              .then((cache) => {
                cache.put(event.request, responseToCache);
              });

            return networkResponse;
          }
        ).catch((error) => {
          console.warn('Fetch failed; returning offline fallback or error for:', event.request.url, error);
          // Optional: Fallback for when network fails and not in cache
          // if (event.request.mode === 'navigate' && event.request.headers.get('accept').includes('text/html')) {
          //   // return caches.match('./offline.html'); // If you had an offline.html
          // }
          // For other types of requests (JS, CSS, images), if not cached and network fails,
          // the browser will handle it as a failed resource load.
          // Rethrow the error to let the browser handle it if no specific fallback is provided.
          // This ensures that if an asset crucial for rendering is missing, it's evident.
          // Alternatively, for some assets, you might want to return a placeholder.
           throw error;
        });
      })
  );
});
