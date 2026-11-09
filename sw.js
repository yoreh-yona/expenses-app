const CACHE = 'expenses-v1.3.2';
const ASSETS = ['./', 'index.html', 'app.js', 'manifest.json', 'xlsx-fill.js', 'imgproc.js',
  'vendor/pizzip.min.js', 'vendor/exceljs.min.js', 'vendor/pdf-lib.min.js', 'icons/icon-192.png', 'icons/icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS.map(u => new Request(u, {cache: 'no-cache'}))))); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
// network-first for app code (fresh updates), cache fallback when offline
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return;
  e.respondWith(fetch(e.request, {cache: 'no-cache'}).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); return r; })
    .catch(() => caches.match(e.request, {ignoreSearch: true})));
});
