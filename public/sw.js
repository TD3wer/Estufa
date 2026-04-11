// sw.js — Phytogestor v3 (offline completo)
const CACHE = 'phytogestor-v3';

const ASSETS = [
  '/',
  '/index.html',
  '/manifest.json',
  '/icons/icon-192.png',
  '/icons/icon-512.png',
  '/icons/apple-touch-icon.png',
  '/icons/favicon.ico',
  // SheetJS — necessário para geração de Excel offline
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  // Fontes Google
  'https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@400;500;600&display=swap',
];

// ── INSTALL: pré-cacheia tudo
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(cache =>
      // addAll falha se um item falhar; usamos allSettled para ser tolerante a falhas de rede
      Promise.allSettled(
        ASSETS.map(url =>
          cache.add(url).catch(err => console.warn('Cache falhou para:', url, err))
        )
      )
    )
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

// ── FETCH: estratégia por tipo de recurso
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // 1. Rotas de API → rede direta, sem cache
  const rotasAPI = ['/salvar', '/dados', '/gerar-excel', '/apagar', '/limpar'];
  if (rotasAPI.some(r => url.pathname.startsWith(r))) {
    e.respondWith(
      fetch(e.request).catch(() =>
        new Response(
          JSON.stringify({ erro: 'offline' }),
          { status: 503, headers: { 'Content-Type': 'application/json' } }
        )
      )
    );
    return;
  }

  // 2. Navegação (HTML) → Network-first, fallback para cache
  if (e.request.mode === 'navigate') {
    e.respondWith(
      fetch(e.request)
        .then(res => {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
          return res;
        })
        .catch(() => caches.match('/index.html'))
    );
    return;
  }

  // 3. Tudo mais (JS, fontes, imagens) → Cache-first, atualiza em background
  e.respondWith(
    caches.match(e.request).then(cached => {
      const fetchPromise = fetch(e.request).then(res => {
        if (res && res.status === 200) {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        return res;
      }).catch(() => null);

      // Retorna cache imediatamente se existir, senão espera a rede
      return cached || fetchPromise;
    })
  );
});