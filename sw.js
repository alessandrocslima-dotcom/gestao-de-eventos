const CACHE_VERSION = 'gep-v15-56c';
const CACHE_NAME = 'gestao-eventos-' + CACHE_VERSION;
const ASSETS = ['./', './index.html'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE_NAME).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => { if(e.request.method !== 'GET') return; e.respondWith(caches.match(e.request).then(r => r || fetch(e.request).then(res => { if(res.ok){ const c = res.clone(); caches.open(CACHE_NAME).then(ca => ca.put(e.request, c)); } return res; }))); });
