/* Service Worker — Gestão de Eventos Pro */
/* Versão do cache: mude este número quando atualizar o app para forçar atualização */
const CACHE_VERSION = 'gep-v15-44';
const CACHE_NAME = 'gestao-eventos-' + CACHE_VERSION;

/* Instala e ativa imediatamente */
self.addEventListener('install', function(event) {
  self.skipWaiting();
});

/* Limpa caches antigos ao ativar */
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
    })
  );
});

/* Estratégia: rede primeiro, cache como reserva (network-first) */
/* Isso garante que o app sempre tente buscar a versão mais nova online, */
/* mas funcione offline usando o que já visitou. */
self.addEventListener('fetch', function(event) {
  /* Só intercepta requisições GET do mesmo site */
  if (event.request.method !== 'GET') return;

  event.respondWith(
    fetch(event.request).then(function(response) {
      /* Guarda uma cópia no cache */
      const copy = response.clone();
      caches.open(CACHE_NAME).then(function(cache) {
        cache.put(event.request, copy);
      });
      return response;
    }).catch(function() {
      /* Sem internet: usa o cache */
      return caches.match(event.request);
    })
  );
});
