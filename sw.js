/* ArmaLife Stok Sorgulama — çevrimdışı önbellek. Önce ağ, olmazsa önbellek. */
var ONBELLEK = 'armastok-v4';
var KABUK = ['./', 'index.html', 'lib/veri.js', 'lib/okuyucu.min.js', 'lib/zxing_reader.wasm', 'lib/xlsx.full.min.js', 'manifest.webmanifest', 'icons/icon-192.png', 'icons/apple-touch-icon.png'];
self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(ONBELLEK).then(function (c) { return c.addAll(KABUK); }).catch(function () {}).then(function () { return self.skipWaiting(); }));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) { return Promise.all(ks.filter(function (k) { return k !== ONBELLEK; }).map(function (k) { return caches.delete(k); })); }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var r = e.request;
  if (r.method !== 'GET' || new URL(r.url).origin !== self.location.origin) return;
  var veri = /veri\.json/.test(r.url);
  e.respondWith(fetch(r).then(function (yanit) {
    if (yanit && yanit.ok && !veri) { var kopya = yanit.clone(); caches.open(ONBELLEK).then(function (c) { c.put(r, kopya); }); }
    return yanit;
  }).catch(function () {
    return caches.match(r, { ignoreSearch: veri }).then(function (m) { return m || (r.mode === 'navigate' ? caches.match('index.html') : Response.error()); });
  }));
});
