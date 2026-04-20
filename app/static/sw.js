const CACHE = 'cetem-flow-v1';

const STATIC_ASSETS = [
  '/static/css/style.css',
  '/static/js/app.js',
];

self.addEventListener('install', ev => {
  ev.waitUntil(
    caches.open(CACHE).then(c => c.addAll(STATIC_ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', ev => {
  ev.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', ev => {
  const { request } = ev;
  const url = new URL(request.url);

  // Só intercepta requisições do próprio domínio
  if (url.origin !== location.origin) return;

  // Recursos estáticos: cache first
  if (url.pathname.startsWith('/static/')) {
    ev.respondWith(
      caches.match(request).then(cached => cached || fetch(request).then(res => {
        const clone = res.clone();
        caches.open(CACHE).then(c => c.put(request, clone));
        return res;
      }))
    );
    return;
  }

  // Páginas HTML: network first (dados sempre frescos), fallback para cache
  if (request.mode === 'navigate') {
    ev.respondWith(
      fetch(request).catch(() => caches.match(request))
    );
  }
});
