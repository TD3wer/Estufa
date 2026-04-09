// sw.js — service worker mínimo
const CACHE = 'estufa-v1';
const ASSETS = ['/', '/index.html'];

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

self.addEventListener('fetch', e => {
  // Deixa passar requisições ao servidor (salvar, gerar-excel etc.)
  if (e.request.url.includes('/salvar') ||
      e.request.url.includes('/dados') ||
      e.request.url.includes('/gerar-excel') ||
      e.request.url.includes('/apagar') ||
      e.request.url.includes('/limpar')) {
    return; // vai direto para a rede
  }
  e.respondWith(
    caches.match(e.request).then(r => r || fetch(e.request))
  );
});