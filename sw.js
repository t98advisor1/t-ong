/* 자리배치 프로그램 - 서비스 워커
 * 앱을 수정한 뒤에는 아래 CACHE_VERSION 을 올려주세요.
 * (올리지 않으면 이미 설치한 사용자에게 옛 파일이 계속 보일 수 있습니다)
 */
var CACHE_VERSION = 'v3.1.0-29';
var CACHE_NAME = 'seat-app-' + CACHE_VERSION;

var PRECACHE = [
  '/',
  '/index.html',
  '/style.css?v=5.17.0-preview.72',
  '/app.js?v=5.17.0-preview.72',
  '/icon.svg',
  '/manifest.webmanifest'
];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(function (cache) { return cache.addAll(PRECACHE); })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys()
      .then(function (keys) {
        return Promise.all(keys.map(function (key) {
          if (key !== CACHE_NAME) return caches.delete(key);
        }));
      })
      .then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (event) {
  var req = event.request;
  if (req.method !== 'GET') return;

  // 실시간 통계 API는 캐시하지 않습니다.
  if (new URL(req.url).pathname.indexOf('/api/') === 0) {
    event.respondWith(fetch(req));
    return;
  }

  // 페이지 이동: 네트워크 우선 (새 버전을 바로 받도록), 실패 시 캐시
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then(function (res) {
          var copy = res.clone();
          caches.open(CACHE_NAME).then(function (c) { c.put('/index.html', copy); });
          return res;
        })
        .catch(function () {
          return caches.match('/index.html');
        })
    );
    return;
  }

  // 그 외 정적 자원(SheetJS CDN 포함): 캐시 우선 → 없으면 네트워크 후 캐시에 저장
  event.respondWith(
    caches.match(req).then(function (cached) {
      if (cached) return cached;
      return fetch(req).then(function (res) {
        // 불투명(opaque) 응답도 캐시해 두면 오프라인에서 스크립트로 사용 가능
        if (res && (res.ok || res.type === 'opaque')) {
          var copy = res.clone();
          caches.open(CACHE_NAME).then(function (c) { c.put(req, copy); });
        }
        return res;
      });
    })
  );
});
