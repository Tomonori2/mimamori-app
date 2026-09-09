/* みまもり ボタン画面 サービスワーカー v1.4
 *
 * サービスワーカー＝アプリを裏で支える小さなプログラム。
 * ここでは「画面の見た目（HTMLや設定ファイル）だけ」を手元にためて、
 * 電波が弱いときでも画面がすぐ開くようにしています。
 *
 * ★ Googleとのやりとり（script.google.com）と天気（open-meteo）は、
 *   ぜったいにためません。古い情報を「きょうの結果」として見せてしまうと、
 *   見守りアプリとしては危険だからです。かならず毎回とりに行きます。
 */
var CACHE = 'mimamori-v1.4';
var SHELL = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE).then(function (c) { return c.addAll(SHELL); })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (e) {
  // 古い版のためこみを消す（スマホで「直したのに変わらない」を防ぐ）
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.map(function (k) {
        if (k !== CACHE) return caches.delete(k);
      }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  var url = e.request.url;

  // 通信が要るものは素通し（ためこまない）
  if (url.indexOf('script.google.com') >= 0 ||
      url.indexOf('googleusercontent.com') >= 0 ||
      url.indexOf('open-meteo.com') >= 0) {
    return;
  }
  if (e.request.method !== 'GET') return;

  // 画面の部品は「まず手元、なければ取りに行く」
  e.respondWith(
    caches.match(e.request).then(function (hit) {
      return hit || fetch(e.request);
    })
  );
});
