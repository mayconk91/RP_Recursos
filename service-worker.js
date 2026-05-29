/* Planejador de Recursos â€” Service Worker (PWA offline)
 * EstratÃ©gia:
 * - App Shell (arquivos estÃ¡ticos) em cache-first
 * - NavegaÃ§Ã£o (HTML) em network-first com fallback no cache
 */

// IMPORTANT: incremente VERSION sempre que houver mudanÃ§as no App Shell.
// Isso forÃ§a a criaÃ§Ã£o de um novo cache e evita que usuÃ¡rios instalados
// fiquem presos em versÃµes antigas.
const VERSION = '1.2.8.58';
const CACHE_NAME = `planner-${VERSION}`;

const APP_SHELL = [
  './',
  './index.html',
  './styles.css?v=1.2.8.58',
  './tour.css',
  './app.js?v=1.2.8.58',
  './tour.js',
  './enhancer.css',
  './enhancer.js',
  './enhancer2.js',
  './planner_enhancer.js',
  './capacidade_v4.html',
  './manifest.webmanifest',
  './version.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/maskable-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(APP_SHELL))
  );
});

// Permite que a UI solicite que o SW em espera assuma o controle.
self.addEventListener('message', (event) => {
  if (!event || !event.data) return;
  if (event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(
      keys
        .filter((k) => k.startsWith('planner-') && k !== CACHE_NAME)
        .map((k) => caches.delete(k))
    )).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // SÃ³ controla o que estiver dentro do escopo
  if (url.origin !== self.location.origin) return;

  // NavegaÃ§Ã£o (document): tenta rede primeiro
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put('./index.html', copy)).catch(() => {});
          return res;
        })
        .catch(() => caches.match('./index.html'))
    );
    return;
  }

  // EstÃ¡ticos: cache-first
  event.respondWith(
    caches.match(req).then((cached) => {
      if (cached) return cached;
      return fetch(req).then((res) => {
        const copy = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(req, copy)).catch(() => {});
        return res;
      });
    })
  );
});
