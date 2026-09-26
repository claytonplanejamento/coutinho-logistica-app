const CACHE = 'coutinho-logistica-v6-shell';
const SHELL = [
  './',
  './index.html',
  './manifest.json',
  './app-config.js',
  './offline.html',
  './logo-coutinho.png',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon-32.png',
  './icons/favicon-16.png'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);
  // Nao interfere no backend do Google Apps Script.
  if (url.hostname.includes('google.com') || url.hostname.includes('googleusercontent.com')) return;

  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req).catch(() => caches.match('./offline.html'))
    );
    return;
  }

  event.respondWith(
    caches.match(req).then(hit => hit || fetch(req).then(resp => {
      const clone = resp.clone();
      caches.open(CACHE).then(cache => cache.put(req, clone));
      return resp;
    }))
  );
});
