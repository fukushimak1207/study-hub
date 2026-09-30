// 画面のファイルを端末に保存して、圏外でも開けるようにする。
// ⚠️ 画面のファイルを直したら CACHE_VERSION を必ず上げる(上げないと古い画面のまま)
const CACHE_VERSION = 'study-hub-v1';
const SHELL = ['./', 'index.html', 'style.css', 'app.js', 'manifest.webmanifest', 'icon.svg'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE_VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  // データ(Apps Script)は保存しない。画面のファイルだけ保存分を先に使う
  if (url.origin !== location.origin || e.request.method !== 'GET') return;
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then(hit => hit || fetch(e.request)));
});
