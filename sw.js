/* Офлайн-режим отключён: приложение всегда грузится с GitHub Pages.

   Этот файл нужен только тем устройствам, где уже стоит service worker от прежней
   версии: браузер при открытии сверяет sw.js с сервером, находит изменения и ставит
   этот. Он удаляет весь кэш, снимает сам себя и перезагружает открытые окна —
   уже напрямую с сервера. Удалять файл нельзя: если sw.js пропадёт (404),
   браузер оставит старый service worker как есть. */

self.addEventListener('install', () => self.skipWaiting());

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    for (const key of await caches.keys()) await caches.delete(key);
    await self.clients.claim();
    await self.registration.unregister();
    const clients = await self.clients.matchAll({ type: 'window' });
    for (const c of clients) c.navigate(c.url).catch(() => {});
  })());
});
