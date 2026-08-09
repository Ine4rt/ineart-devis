/**
 * Service worker : mise en cache « cache-first » de la coquille applicative.
 * Le jeu tient en quelques dizaines de kilo-octets et n'a aucun asset binaire :
 * après la première visite, il démarre hors ligne et instantanément.
 */
const CACHE = 'voidrunner-v1';
const SHELL = [
  './', './index.html', './styles.css', './manifest.webmanifest', './icon.svg',
  './src/main.js',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((ks) =>
    Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request).then((hit) => hit || fetch(e.request).then((res) => {
      const copy = res.clone();
      caches.open(CACHE).then((c) => c.put(e.request, copy)).catch(() => {});
      return res;
    }).catch(() => caches.match('./index.html'))),
  );
});
