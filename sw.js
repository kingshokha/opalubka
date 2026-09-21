/* Service worker: приложение открывается без интернета.
   Стратегия — stale-while-revalidate: сразу отдаём копию из кэша (быстро и работает
   офлайн), параллельно тянем свежую версию и кладём в кэш на следующий запуск.
   При смене версии страницы клиенту уходит сообщение и он показывает плашку «обновить». */

const CACHE = 'opalubka-v1';
const ASSETS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE)
      .then(c => c.addAll(ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  if (new URL(req.url).origin !== self.location.origin) return;

  e.respondWith((async () => {
    const cache = await caches.open(CACHE);
    const cached = await cache.match(req, { ignoreSearch: true });

    const network = fetch(req).then(async res => {
      if (res && res.ok) {
        const copy = res.clone();
        // сравниваем со старой копией: если страница изменилась — зовём клиента обновиться
        if (cached && req.destination === 'document') {
          const [a, b] = await Promise.all([cached.clone().text(), res.clone().text()]);
          if (a !== b) {
            const clients = await self.clients.matchAll({ type: 'window' });
            clients.forEach(c => c.postMessage({ type: 'update-ready' }));
          }
        }
        await cache.put(req, copy);
      }
      return res;
    }).catch(() => null);

    return cached || (await network) || new Response('Нет соединения', {
      status: 503,
      headers: { 'Content-Type': 'text/plain; charset=utf-8' }
    });
  })());
});
