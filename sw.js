/* Service Worker — GEP Gestão de Eventos Pro */
const CACHE_VERSION = 'gep-v15-82';
const CACHE_NAME = 'gestao-eventos-' + CACHE_VERSION;

const PRECACHE = [
  '/gestao-de-eventos/',
  '/gestao-de-eventos/index.html',
  '/gestao-de-eventos/manifest.json',
  '/gestao-de-eventos/icon-192.png',
  '/gestao-de-eventos/icon-512.png',
  '/gestao-de-eventos/maskable-192.png'
];

self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(PRECACHE);
    }).then(function() {
      return self.skipWaiting();
    })
  );
});

self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) {
          return k.indexOf('gestao-eventos-') === 0 && k !== CACHE_NAME;
        }).map(function(k) { return caches.delete(k); })
      );
    }).then(function() {
      return self.clients.claim();
    }).then(function() {
      return self.clients.matchAll({ type: 'window' }).then(function(clients) {
        clients.forEach(function(client) {
          client.postMessage({ type: 'NEW_VERSION', version: CACHE_VERSION });
        });
      });
    })
  );
});

self.addEventListener('fetch', function(event) {
  if (event.request.method !== 'GET') return;
  if (event.request.url.indexOf('chrome-extension') !== -1) return;
  if (event.request.url.indexOf('googleapis.com') !== -1) return;

  event.respondWith(
    fetch(event.request).then(function(response) {
      if (response && response.status === 200 && response.type !== 'opaque') {
        var copy = response.clone();
        caches.open(CACHE_NAME).then(function(cache) {
          cache.put(event.request, copy);
        });
      }
      return response;
    }).catch(function() {
      return caches.match(event.request).then(function(cached) {
        return cached || caches.match('/gestao-de-eventos/index.html');
      });
    })
  );
});
