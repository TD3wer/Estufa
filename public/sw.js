// sw.js — Phytogestor
// Estratégia: Cache-first para assets estáticos, Network-first para API
const CACHE = 'phytogestor-v2';

const ASSETS_ESTATICOS = [
  '/',
  '/index.html',
  '/manifest.json',
  '/icons/icon-192.png',
  '/icons/icon-512.png',
  '/icons/apple-touch-icon.png',
  '/icons/favicon.ico',
  // Fontes do Google (pré-cache via fetch no install)
];

// Fontes do Google Fonts — cacheadas separadamente pois são cross-origin
const FONTES = [
  'https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@400;500;600&display=swap',
];

// ── INSTALL: pré-carrega todos os assets estáticos
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(async cache => {
      // Assets locais
      await cache.addAll(ASSETS_ESTATICOS);
      // Fontes: busca e cacheia (cross-origin, não pode usar addAll diretamente)
      await Promise.allSettled(
        FONTES.map(url =>
          fetch(url, { mode: 'cors' })
            .then(res => { if (res.ok) cache.put(url, res); })
            .catch(() => {}) // se falhar (offline no install) ignora
        )
      );
    })
  );
  self.skipWaiting();
});

// ── ACTIVATE: remove caches antigos
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// ── FETCH: decide estratégia por tipo de requisição
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // 1. Requisições de API → sempre vai para a rede (nunca cacheia)
  const rotasAPI = ['/salvar', '/dados', '/gerar-excel', '/apagar', '/limpar'];
  if (rotasAPI.some(r => url.pathname.startsWith(r))) {
    e.respondWith(
      fetch(e.request).catch(() =>
        new Response(
          JSON.stringify({ erro: 'Sem conexão. Dados não puderam ser salvos.' }),
          { status: 503, headers: { 'Content-Type': 'application/json' } }
        )
      )
    );
    return;
  }

  // 2. Navegação (HTML) → Network-first com fallback para cache
  if (e.request.mode === 'navigate') {
    e.respondWith(
      fetch(e.request)
        .then(res => {
          // Atualiza cache com versão mais recente
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
          return res;
        })
        .catch(() => caches.match('/index.html'))
    );
    return;
  }

  // 3. Fontes e assets estáticos → Cache-first
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      // Não está no cache, busca na rede e guarda
      return fetch(e.request).then(res => {
        if (res && res.status === 200) {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        return res;
      }).catch(() => {
        // Se for uma imagem e não tiver cache, retorna resposta vazia
        if (e.request.destination === 'image') {
          return new Response('', { status: 404 });
        }
      });
    })
  );
});