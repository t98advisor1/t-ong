/* 자리배치 프로그램 - 서비스 워커
 * 앱을 수정한 뒤에는 아래 CACHE_VERSION 을 올려주세요.
 * (올리지 않으면 이미 설치한 사용자에게 옛 파일이 계속 보일 수 있습니다)
 */
var CACHE_VERSION = 'v3.2.0-73';
var CACHE_NAME = 'seat-app-' + CACHE_VERSION;

var PRECACHE = [
  '/',
  '/index.html',
  '/style.css?v=5.17.0-preview.90',
  '/app.js?v=5.17.0-preview.90',
  '/smart-board.html',
  '/smart-board-guide.html',
  '/smart-board.webmanifest',
  '/memo-board-guide.html',
  '/score-planner-guide.html',
  '/doc-converter-guide.html',
  '/smart-board.css?v=2.6.2',
  '/smart-board.js?v=2.6.2',
  '/qr-lib.js',
  '/geo-services.js?v=1.0.0',
  '/libs/jszip.min.js',
  '/libs/pdf-lib.min.js',
  '/libs/pdf.min.js',
  '/libs/pdf.worker.min.js',
  '/board.html',
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
  // 실시간 API 요청(/api/*)은 서비스 워커를 거치지 않고 브라우저가 직접 통신합니다.
  if (new URL(req.url).pathname.indexOf('/api/') === 0) return;
  if (req.method !== 'GET') return;

  // 매뉴얼 영상은 구간(Range) 요청으로 받아야 해서 서비스 워커를 거치지 않고 네트워크로 바로 받습니다.
  if (/\.mp4$/i.test(new URL(req.url).pathname)) return;

  // 페이지 이동: 네트워크 우선 (새 버전을 바로 받도록), 실패 시 요청 URL 또는 index.html 캐시
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req, { cache: 'no-cache' })
        .then(function (res) {
          var copy = res.clone();
          caches.open(CACHE_NAME).then(function (c) { c.put(req, copy); });
          return res;
        })
        .catch(function () {
          return caches.match(req).then(function (m) {
            return m || caches.match('/index.html');
          });
        })
    );
    return;
  }

  // 앱 핵심 자산(style.css, app.js, smart-board.*)은 네트워크 우선(no-cache)으로 받아 화면 깨짐/구버전 캐시 고착을 방지합니다 (/ 및 /tong/ 경로 모두 대응).
  var urlPath = new URL(req.url).pathname;
  if (urlPath.endsWith('/style.css') || urlPath.endsWith('/app.js') || urlPath.endsWith('/index.html') ||
      urlPath.endsWith('/geo-services.js') || urlPath.endsWith('/board.html') ||
      urlPath.endsWith('/smart-board.html') || urlPath.endsWith('/smart-board-guide.html') ||
      urlPath.endsWith('/smart-board.css') || urlPath.endsWith('/smart-board.js')) {
    event.respondWith(
      fetch(req, { cache: 'no-cache' })
        .then(function (res) {
          if (res && res.ok) {
            var copy = res.clone();
            caches.open(CACHE_NAME).then(function (c) { c.put(req, copy); });
          }
          return res;
        })
        .catch(function () {
          return caches.match(req);
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
