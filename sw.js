/* Offline support: caches the app files so it opens without a connection.
   Bump VERSION whenever you upload changed files so phones pick them up. */
const VERSION = 'aa-v3';
const FILES = ['./', 'index.html', 'styles.css', 'app.js', 'fields.js', 'xlsx.js', 'vendor/jszip.min.js',
  'manifest.webmanifest', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return; // AI + Teams calls go straight to the network
  // Network first so updates show up; fall back to the cached copy when offline.
  e.respondWith(fetch(req).then(res => {
    const copy = res.clone();
    if (res.ok) caches.open(VERSION).then(c => c.put(req, copy));
    return res;
  }).catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || caches.match('index.html'))));
});
