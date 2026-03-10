/* Neotrace Service Worker — Offline-First Strategy */

var CACHE_NAME = 'neotrace-v1';
var ASSETS_TO_CACHE = [
    '/',
    '/index.html',
    '/styles.css',
    '/app.js',
    '/logo.svg',
    '/manifest.json',
    'https://cdn.jsdelivr.net/npm/qrcode@1.5.3/build/qrcode.min.js',
    'https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js'
];

// Install — pre-cache all critical assets
self.addEventListener('install', function (event) {
    event.waitUntil(
        caches.open(CACHE_NAME).then(function (cache) {
            console.log('[SW] Pre-caching app shell');
            return cache.addAll(ASSETS_TO_CACHE);
        }).then(function () {
            return self.skipWaiting();
        })
    );
});

// Activate — clean old caches
self.addEventListener('activate', function (event) {
    event.waitUntil(
        caches.keys().then(function (cacheNames) {
            return Promise.all(
                cacheNames.filter(function (name) {
                    return name !== CACHE_NAME;
                }).map(function (name) {
                    console.log('[SW] Removing old cache:', name);
                    return caches.delete(name);
                })
            );
        }).then(function () {
            return self.clients.claim();
        })
    );
});

// Fetch — cache-first for app assets, network-first for API calls
self.addEventListener('fetch', function (event) {
    var url = new URL(event.request.url);

    // Network-first for Google Apps Script API calls
    if (url.hostname === 'script.google.com' || url.hostname === 'script.googleusercontent.com') {
        event.respondWith(
            fetch(event.request).catch(function () {
                // If offline, queue for later sync
                return new Response(JSON.stringify({ offline: true, message: 'Request queued for sync' }), {
                    headers: { 'Content-Type': 'application/json' }
                });
            })
        );
        return;
    }

    // Cache-first for everything else
    event.respondWith(
        caches.match(event.request).then(function (cachedResponse) {
            if (cachedResponse) {
                // Return cache, but also update cache in background
                fetch(event.request).then(function (networkResponse) {
                    if (networkResponse && networkResponse.status === 200) {
                        caches.open(CACHE_NAME).then(function (cache) {
                            cache.put(event.request, networkResponse);
                        });
                    }
                }).catch(function () { /* ignore network failures */ });
                return cachedResponse;
            }

            // Not in cache — fetch from network
            return fetch(event.request).then(function (networkResponse) {
                if (networkResponse && networkResponse.status === 200) {
                    var responseClone = networkResponse.clone();
                    caches.open(CACHE_NAME).then(function (cache) {
                        cache.put(event.request, responseClone);
                    });
                }
                return networkResponse;
            }).catch(function () {
                // Fallback for HTML pages
                if (event.request.headers.get('accept') && event.request.headers.get('accept').indexOf('text/html') !== -1) {
                    return caches.match('/index.html');
                }
            });
        })
    );
});

// Background Sync — retry failed QC submissions when back online
self.addEventListener('sync', function (event) {
    if (event.tag === 'qc-sync') {
        event.waitUntil(syncPendingQCReports());
    }
});

function syncPendingQCReports() {
    // Read pending submissions from IndexedDB and retry
    return self.clients.matchAll().then(function (clients) {
        clients.forEach(function (client) {
            client.postMessage({ type: 'SYNC_QC_REPORTS' });
        });
    });
}
