/* Service Worker — Gestão de Eventos Pro */
const CACHE_VERSION = 'gep-v15-72';
const CACHE_NAME = 'gestao-eventos-' + CACHE_VERSION;

/* Instala e assume imediatamente — sem esperar fechar abas */
self.addEventListener('install', function(event) {
  self.skipWaiting();
});

/* Ao ativar: limpa caches antigos e avisa clientes que há versão nova */
self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) {
          return k.indexOf('gestao-eventos-') === 0 && k !== CACHE_NAME;
        }).map(function(k) {
          return caches.delete(k);
        })
      );
    }).then(function() {
      return self.clients.claim();
    }).then(function() {
      /* Avisa todas as abas abertas que a versão mudou */
      return self.clients.matchAll({ type: 'window' }).then(function(clients) {
        clients.forEach(function(client) {
          client.postMessage({ type: 'NEW_VERSION', version: CACHE_VERSION });
        });
      });
    })
  );
});

/* Estratégia: rede primeiro, cache como reserva */
self.addEventListener('fetch', function(event) {
  if (event.request.method !== 'GET') return;

  event.respondWith(
    fetch(event.request).then(function(response) {
      const copy = response.clone();
      caches.open(CACHE_NAME).then(function(cache) {
        cache.put(event.request, copy);
      });
      return response;
    }).catch(function() {
      return caches.match(event.request);
    })
  );
});
