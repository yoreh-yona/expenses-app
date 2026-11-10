const CACHE = 'expenses-v2.1.0';
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

// report reminders (sent by the scheduled GitHub Action)
self.addEventListener('push', e => {
  let d = {}; try { d = e.data.json(); } catch { d = {body: e.data && e.data.text()}; }
  e.waitUntil(self.registration.showNotification(d.title || 'דוח הוצאות', {
    body: d.body || 'הדוח עוד לא נשלח', tag: 'report-nag', renotify: true, requireInteraction: true,
    icon: 'icons/icon-192.png', badge: 'icons/icon-192.png', dir: 'rtl', lang: 'he', data: {url: d.url || './#report'}
  }));
});
self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil(self.clients.matchAll({type: 'window', includeUncontrolled: true}).then(ws => {
    for (const w of ws) { w.postMessage({go: 'report'}); if ('focus' in w) return w.focus(); }
    return self.clients.openWindow(new URL(e.notification.data.url, self.registration.scope).href);
  }));
});
