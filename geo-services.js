/* ============================================================================
 * T-ong Geo & Seoul Open Data Services
 *  - Vercel 환경변수 연동: KAKAO_MAP_JS_KEY (/api/config), SEOUL_OPEN_API_KEY (/api/seoul-data)
 *  - 1) 수업 메모보드(패들렛형) 카카오맵 지도 보드 컨트롤러
 *  - 2) 오늘의 알림장 서울시 자치구 실시간 미세먼지 배지
 *  - 3) 서울 공공데이터 × 카카오맵 수업 탐구 도구 (표 · 지도 시각화 · CSV 추출)
 * ============================================================================ */
(function () {
  'use strict';

  var LOCAL_KAKAO_KEY = 'tong_kakao_map_js_key_v1';
  var LOCAL_GU_KEY = 'tong_seoul_gu_v1';

  var SEOUL_GU_COORDS = {
    '종로구': [37.5730, 126.9794], '중구': [37.5641, 126.9979], '용산구': [37.5326, 126.9900],
    '성동구': [37.5633, 127.0369], '광진구': [37.5385, 127.0823], '동대문구': [37.5744, 127.0396],
    '중랑구': [37.6066, 127.0927], '성북구': [37.5894, 127.0167], '강북구': [37.6396, 127.0257],
    '도봉구': [37.6688, 127.0471], '노원구': [37.6542, 127.0568], '은평구': [37.6027, 126.9291],
    '서대문구': [37.5791, 126.9368], '마포구': [37.5663, 126.9019], '양천구': [37.5169, 126.8664],
    '강서구': [37.5509, 126.8495], '구로구': [37.4954, 126.8874], '금천구': [37.4568, 126.8955],
    '영등포구': [37.5264, 126.8962], '동작구': [37.5124, 126.9393], '관악구': [37.4784, 126.9516],
    '서초구': [37.4837, 127.0324], '강남구': [37.5172, 127.0473], '송파구': [37.5145, 127.1059],
    '강동구': [37.5301, 127.1238]
  };

  var configPromise = null;
  var kakaoSdkPromise = null;
  var leafletSdkPromise = null;

  function loadConfig() {
    if (configPromise) return configPromise;
    configPromise = fetch('/api/config', { cache: 'no-store' })
      .then(function (r) { return r.ok ? r.json() : null; })
      .catch(function () { return null; })
      .then(function (cfg) {
        var localKey = '';
        try { localKey = (localStorage.getItem(LOCAL_KAKAO_KEY) || '').trim(); } catch (e) {}
        var envKey = (cfg && cfg.kakaoMapJsKey) ? String(cfg.kakaoMapJsKey).trim() : '';
        var finalKey = envKey || localKey;
        return {
          kakaoMapJsKey: finalKey,
          hasKakaoMapKey: Boolean(finalKey),
          hasSeoulApi: Boolean(cfg && cfg.hasSeoulApi)
        };
      });
    return configPromise;
  }

  function ensureKakaoMaps(appKey) {
    if (!appKey) return Promise.resolve(null);
    if (window.kakao && window.kakao.maps && window.kakao.maps.Map) {
      return Promise.resolve(window.kakao.maps);
    }
    if (kakaoSdkPromise) return kakaoSdkPromise;
    kakaoSdkPromise = new Promise(function (resolve) {
      var s = document.createElement('script');
      s.src = 'https://dapi.kakao.com/v2/maps/sdk.js?appkey=' + encodeURIComponent(appKey) + '&libraries=services&autoload=false';
      s.async = true;
      s.onload = function () {
        if (window.kakao && window.kakao.maps && typeof window.kakao.maps.load === 'function') {
          window.kakao.maps.load(function () { resolve(window.kakao.maps); });
        } else {
          resolve(null);
        }
      };
      s.onerror = function () { resolve(null); };
      document.head.appendChild(s);
    });
    return kakaoSdkPromise;
  }

  function ensureLeaflet() {
    if (window.L && window.L.map) return Promise.resolve(window.L);
    if (leafletSdkPromise) return leafletSdkPromise;
    leafletSdkPromise = new Promise(function (resolve) {
      if (!document.querySelector('link[data-tong-leaflet]')) {
        var lk = document.createElement('link');
        lk.rel = 'stylesheet';
        lk.href = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css';
        lk.setAttribute('data-tong-leaflet', '1');
        document.head.appendChild(lk);
      }
      var s = document.createElement('script');
      s.src = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js';
      s.async = true;
      s.onload = function () { resolve(window.L || null); };
      s.onerror = function () { resolve(null); };
      document.head.appendChild(s);
    });
    return leafletSdkPromise;
  }

  // 메모 텍스트 내 지도 태그 인코딩/디코딩 (기존 Google Apps Script 무수정 호환)
  var MAP_TAG_RE = /\s*\[\[MAP:(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)(?:\|([^\]]*))?\]\]\s*$/;

  function parseMapTag(text) {
    var raw = String(text || '');
    var m = raw.match(MAP_TAG_RE);
    if (!m) return { lat: null, lng: null, place: '', cleanText: raw };
    return {
      lat: parseFloat(m[1]),
      lng: parseFloat(m[2]),
      place: (m[3] || '').trim(),
      cleanText: raw.replace(MAP_TAG_RE, '').trim()
    };
  }

  function encodeMapTag(text, loc) {
    var clean = String(text || '').replace(MAP_TAG_RE, '').trim();
    if (!loc || typeof loc.lat !== 'number' || typeof loc.lng !== 'number' || isNaN(loc.lat) || isNaN(loc.lng)) {
      return clean;
    }
    var safePlace = String(loc.place || '').replace(/[\[\]|]/g, '').trim().slice(0, 40);
    return clean + ' [[MAP:' + loc.lat.toFixed(6) + ',' + loc.lng.toFixed(6) + '|' + safePlace + ']]';
  }

  function escapeHtml(s) {
    return String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // 통합 지도 어댑터 (카카오맵 우선, 미설정 시 Leaflet/OSM 자동 폴백)
  async function createMapAdapter(containerEl, opts) {
    opts = opts || {};
    var centerLat = opts.lat || 37.5665;
    var centerLng = opts.lng || 126.9780;
    var cfg = await loadConfig();
    var kmaps = cfg.hasKakaoMapKey ? await ensureKakaoMaps(cfg.kakaoMapJsKey) : null;

    containerEl.innerHTML = '';
    var mapDiv = document.createElement('div');
    mapDiv.style.width = '100%';
    mapDiv.style.height = '100%';
    containerEl.appendChild(mapDiv);

    var providerBadge = document.createElement('div');
    providerBadge.className = 'geo-provider-badge';
    containerEl.appendChild(providerBadge);

    if (kmaps) {
      providerBadge.textContent = '카카오맵 연동됨';
      providerBadge.classList.add('is-kakao');
      var map = new kmaps.Map(mapDiv, {
        center: new kmaps.LatLng(centerLat, centerLng),
        level: opts.level || 7
      });
      var zoomControl = new kmaps.ZoomControl();
      map.addControl(zoomControl, kmaps.ControlPosition.RIGHT);
      var placesSvc = kmaps.services ? new kmaps.services.Places() : null;
      var geocoder = kmaps.services ? new kmaps.services.Geocoder() : null;
      var markers = [];
      var overlays = [];
      var activeInfo = new kmaps.InfoWindow({ zIndex: 10, removable: true });
      var tempPin = null;

      return {
        provider: 'kakao',
        relayout: function () { map.relayout(); },
        setCenter: function (lat, lng, level) {
          map.setCenter(new kmaps.LatLng(lat, lng));
          if (level) map.setLevel(level);
        },
        clearMarkers: function () {
          markers.forEach(function (m) { m.setMap(null); });
          overlays.forEach(function (o) { o.setMap(null); });
          markers = [];
          overlays = [];
          activeInfo.close();
        },
        setTempPin: function (lat, lng) {
          if (tempPin) tempPin.setMap(null);
          if (lat == null || lng == null) { tempPin = null; return; }
          tempPin = new kmaps.Marker({
            position: new kmaps.LatLng(lat, lng),
            map: map
          });
        },
        addMarker: function (item) {
          var pos = new kmaps.LatLng(item.lat, item.lng);
          var dot = document.createElement('div');
          dot.className = 'geo-pin-badge';
          dot.style.background = item.color || '#2563eb';
          dot.style.color = item.textColor || '#0f172a';
          dot.textContent = item.label || '핀';
          dot.title = item.title || '';
          var overlay = new kmaps.CustomOverlay({
            position: pos,
            content: dot,
            yAnchor: 1.0
          });
          overlay.setMap(map);
          overlays.push(overlay);

          var openInfo = function () {
            if (item.html) {
              activeInfo.setContent('<div class="geo-infowin">' + item.html + '</div>');
              activeInfo.setPosition(pos);
              activeInfo.open(map);
            }
            if (typeof item.onClick === 'function') item.onClick(item);
          };
          dot.addEventListener('click', openInfo);
          return { openInfo: openInfo, lat: item.lat, lng: item.lng };
        },
        fitPoints: function (pts) {
          if (!pts || !pts.length) return;
          if (pts.length === 1) {
            map.setCenter(new kmaps.LatLng(pts[0].lat, pts[0].lng));
            map.setLevel(5);
            return;
          }
          var bounds = new kmaps.LatLngBounds();
          pts.forEach(function (p) { bounds.extend(new kmaps.LatLng(p.lat, p.lng)); });
          map.setBounds(bounds);
        },
        onClick: function (cb) {
          kmaps.event.addListener(map, 'click', function (mouseEvent) {
            var ll = mouseEvent.latLng;
            var lat = ll.getLat(), lng = ll.getLng();
            if (geocoder) {
              geocoder.coord2Address(lng, lat, function (res, status) {
                var addr = '';
                if (status === kmaps.services.Status.OK && res && res[0]) {
                  addr = (res[0].road_address && res[0].road_address.address_name) || (res[0].address && res[0].address.address_name) || '';
                }
                cb({ lat: lat, lng: lng, place: addr || (lat.toFixed(4) + ', ' + lng.toFixed(4)) });
              });
            } else {
              cb({ lat: lat, lng: lng, place: lat.toFixed(4) + ', ' + lng.toFixed(4) });
            }
          });
        },
        searchPlaces: function (keyword) {
          return new Promise(function (resolve) {
            if (!placesSvc || !keyword) return resolve([]);
            placesSvc.keywordSearch(keyword, function (data, status) {
              if (status === kmaps.services.Status.OK && Array.isArray(data)) {
                resolve(data.slice(0, 8).map(function (d) {
                  return {
                    place: d.place_name,
                    address: d.road_address_name || d.address_name || '',
                    lat: parseFloat(d.y),
                    lng: parseFloat(d.x)
                  };
                }));
              } else {
                resolve([]);
              }
            });
          });
        }
      };
    }

    // Fallback: Leaflet + OpenStreetMap (키 미등록 로컬/테스트 환경에서도 즉시 동작)
    var L = await ensureLeaflet();
    if (!L) {
      mapDiv.innerHTML = '<div style="padding:24px;color:#475569;text-align:center;">지도를 불러오지 못했습니다. 네트워크 연결 또는 KAKAO_MAP_JS_KEY 환경변수를 확인해 주세요.</div>';
      return null;
    }
    providerBadge.textContent = '무료 지도 모드 (KAKAO_MAP_JS_KEY 등록 시 카카오맵 전환)';
    var lmap = L.map(mapDiv).setView([centerLat, centerLng], opts.leafletZoom || 12);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      maxZoom: 19,
      attribution: '&copy; OpenStreetMap'
    }).addTo(lmap);
    var lMarkers = [];
    var lTempPin = null;

    return {
      provider: 'leaflet',
      relayout: function () { setTimeout(function () { lmap.invalidateSize(); }, 60); },
      setCenter: function (lat, lng, level) {
        lmap.setView([lat, lng], level ? Math.max(10, 19 - level) : 14);
      },
      clearMarkers: function () {
        lMarkers.forEach(function (m) { lmap.removeLayer(m); });
        lMarkers = [];
      },
      setTempPin: function (lat, lng) {
        if (lTempPin) lmap.removeLayer(lTempPin);
        if (lat == null || lng == null) { lTempPin = null; return; }
        lTempPin = L.marker([lat, lng]).addTo(lmap);
      },
      addMarker: function (item) {
        var icon = L.divIcon({
          className: 'geo-leaflet-pin-wrap',
          html: '<div class="geo-pin-badge" style="background:' + (item.color || '#2563eb') + ';color:' + (item.textColor || '#0f172a') + ';">' + escapeHtml(item.label || '핀') + '</div>',
          iconSize: [null, null],
          iconAnchor: [24, 28]
        });
        var mk = L.marker([item.lat, item.lng], { icon: icon }).addTo(lmap);
        if (item.html) mk.bindPopup('<div class="geo-infowin">' + item.html + '</div>');
        mk.on('click', function () {
          if (typeof item.onClick === 'function') item.onClick(item);
        });
        lMarkers.push(mk);
        return {
          openInfo: function () {
            lmap.panTo([item.lat, item.lng]);
            mk.openPopup();
          },
          lat: item.lat,
          lng: item.lng
        };
      },
      fitPoints: function (pts) {
        if (!pts || !pts.length) return;
        if (pts.length === 1) { lmap.setView([pts[0].lat, pts[0].lng], 15); return; }
        var b = L.latLngBounds(pts.map(function (p) { return [p.lat, p.lng]; }));
        lmap.fitBounds(b, { padding: [32, 32] });
      },
      onClick: function (cb) {
        lmap.on('click', function (ev) {
          var lat = ev.latlng.lat, lng = ev.latlng.lng;
          cb({ lat: lat, lng: lng, place: lat.toFixed(4) + ', ' + lng.toFixed(4) });
        });
      },
      searchPlaces: function (keyword) {
        if (!keyword) return Promise.resolve([]);
        return fetch('https://nominatim.openstreetmap.org/search?format=json&limit=6&countrycodes=kr&q=' + encodeURIComponent(keyword))
          .then(function (r) { return r.ok ? r.json() : []; })
          .then(function (rows) {
            return (rows || []).map(function (d) {
              return {
                place: (d.display_name || '').split(',')[0],
                address: d.display_name || '',
                lat: parseFloat(d.lat),
                lng: parseFloat(d.lon)
              };
            });
          })
          .catch(function () { return []; });
      }
    };
  }

  // 서울 열린데이터 광장 프록시 호출 헬퍼
  async function fetchSeoulData(service, opts) {
    opts = opts || {};
    var qs = new URLSearchParams();
    qs.set('service', service);
    qs.set('start', String(opts.start || 1));
    qs.set('end', String(opts.end || 50));
    if (opts.p1) qs.set('p1', opts.p1);
    if (opts.p2) qs.set('p2', opts.p2);
    if (opts.p3) qs.set('p3', opts.p3);
    try {
      var res = await fetch('/api/seoul-data?' + qs.toString(), { cache: 'no-store' });
      var data = await res.json();
      return data;
    } catch (err) {
      return { ok: false, error: 'NETWORK_ERROR', message: '서울 공공데이터 서버에 연결할 수 없습니다.' };
    }
  }

  // ============================================================================
  // 1. 수업 메모보드(패들렛형) 지도 보드 컨트롤러
  // ============================================================================
  var MemoBoardMap = (function () {
    var adapter = null;
    var initializing = false;
    var pendingLoc = null;
    var markerHandles = {};
    var bound = false;

    function updateSelectedLocUi() {
      var locLabel = document.getElementById('mbMapSelectedLoc');
      var clearBtn = document.getElementById('mbMapClearLocBtn');
      if (locLabel) {
        if (pendingLoc) {
          locLabel.textContent = '선택 위치: ' + (pendingLoc.place || (pendingLoc.lat.toFixed(4) + ', ' + pendingLoc.lng.toFixed(4)));
          locLabel.classList.add('has-loc');
        } else {
          locLabel.textContent = '지도 클릭 또는 장소 검색으로 핀 위치를 지정하세요';
          locLabel.classList.remove('has-loc');
        }
      }
      if (clearBtn) clearBtn.style.display = pendingLoc ? 'inline-block' : 'none';
      if (adapter) {
        if (pendingLoc) adapter.setTempPin(pendingLoc.lat, pendingLoc.lng);
        else adapter.setTempPin(null, null);
      }
    }

    function bindUi() {
      if (bound) return;
      bound = true;
      var searchInp = document.getElementById('mbMapSearchInput');
      var searchBtn = document.getElementById('mbMapSearchBtn');
      var resBox = document.getElementById('mbMapSearchResults');
      var clearBtn = document.getElementById('mbMapClearLocBtn');

      async function doSearch() {
        var q = (searchInp && searchInp.value || '').trim();
        if (!q || !adapter) return;
        if (resBox) {
          resBox.style.display = 'block';
          resBox.innerHTML = '<div class="mb-map-res-item">장소 검색 중...</div>';
        }
        var items = await adapter.searchPlaces(q);
        if (!resBox) return;
        if (!items.length) {
          resBox.innerHTML = '<div class="mb-map-res-item">검색 결과가 없습니다.</div>';
          return;
        }
        resBox.innerHTML = '';
        items.forEach(function (it) {
          var btn = document.createElement('button');
          btn.type = 'button';
          btn.className = 'mb-map-res-item';
          btn.innerHTML = '<b>' + escapeHtml(it.place) + '</b><small>' + escapeHtml(it.address) + '</small>';
          btn.addEventListener('click', function () {
            pendingLoc = { lat: it.lat, lng: it.lng, place: it.place };
            updateSelectedLocUi();
            adapter.setCenter(it.lat, it.lng, 4);
            resBox.style.display = 'none';
          });
          resBox.appendChild(btn);
        });
      }

      if (searchBtn) searchBtn.addEventListener('click', doSearch);
      if (searchInp) searchInp.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') { e.preventDefault(); doSearch(); }
      });
      if (clearBtn) clearBtn.addEventListener('click', function () {
        pendingLoc = null;
        updateSelectedLocUi();
      });
    }

    async function ensureReady() {
      bindUi();
      var canvas = document.getElementById('memoBoardMapCanvas');
      if (!canvas) return null;
      if (adapter) {
        adapter.relayout();
        return adapter;
      }
      if (initializing) return null;
      initializing = true;
      try {
        adapter = await createMapAdapter(canvas, { lat: 37.5665, lng: 126.9780, level: 7 });
        if (adapter) {
          adapter.onClick(function (loc) {
            pendingLoc = loc;
            updateSelectedLocUi();
          });
        }
      } finally {
        initializing = false;
      }
      return adapter;
    }

    async function render(notesList, opts) {
      opts = opts || {};
      var mapWrap = document.getElementById('memoBoardMapWrap');
      var side = document.getElementById('memoBoardMapSide');
      if (!mapWrap || !side) return;
      var ad = await ensureReady();
      side.innerHTML = '';
      markerHandles = {};

      var pinnedPts = [];
      if (ad) ad.clearMarkers();

      if (!notesList || !notesList.length) {
        side.innerHTML = '<div class="mb-map-empty">아직 올라온 지도 메모가 없습니다.<br>장소를 검색하거나 지도를 클릭한 뒤 메모를 붙이거나, <b>[QR 학생 참여]</b>로 학생 핀을 모아 보세요.</div>';
        return;
      }

      notesList.forEach(function (item, idx) {
        var parsed = parseMapTag(item.text);
        var hasLoc = typeof parsed.lat === 'number' && !isNaN(parsed.lat);
        var handle = null;

        if (hasLoc && ad) {
          pinnedPts.push({ lat: parsed.lat, lng: parsed.lng });
          var shortAuthor = String(item.author || '익명').slice(0, 10);
          var infoHtml =
            '<div class="geo-iw-head"><b>' + escapeHtml(item.author || '익명') + '</b>' +
            (parsed.place ? '<span class="geo-iw-place">' + escapeHtml(parsed.place) + '</span>' : '') +
            '</div><div class="geo-iw-body">' + escapeHtml(parsed.cleanText) + '</div>';
          handle = ad.addMarker({
            lat: parsed.lat,
            lng: parsed.lng,
            color: item.color || '#fef3c7',
            textColor: '#0f172a',
            label: (idx + 1) + '. ' + shortAuthor,
            title: parsed.place || item.author,
            html: infoHtml
          });
          markerHandles[item.id] = handle;
        }

        var card = document.createElement('div');
        card.className = 'mb-map-note-card' + (hasLoc ? ' has-pin' : '');
        card.style.background = item.color || '#fef3c7';
        card.innerHTML =
          '<div class="mb-map-note-head">' +
            '<b>' + escapeHtml(item.author || '익명') + '</b>' +
            '<span class="mb-map-note-loc">' + (hasLoc ? escapeHtml(parsed.place || '핀 지정됨') : '위치 미지정') + '</span>' +
          '</div>' +
          '<div class="mb-map-note-body">' + escapeHtml(parsed.cleanText) + '</div>' +
          '<div class="mb-map-note-foot">' +
            '<button type="button" class="memo-sticky-btn" data-act="like">공감 ' + (Number(item.likes) || 0) + '</button>' +
            '<button type="button" class="memo-sticky-btn" data-act="del">삭제</button>' +
          '</div>';

        card.addEventListener('click', function (ev) {
          var btn = ev.target.closest('button[data-act]');
          if (btn) {
            ev.stopPropagation();
            if (btn.getAttribute('data-act') === 'like' && opts.onLike) opts.onLike(item.id);
            if (btn.getAttribute('data-act') === 'del' && opts.onDelete) opts.onDelete(item.id);
            return;
          }
          if (handle) {
            ad.setCenter(handle.lat, handle.lng, 4);
            handle.openInfo();
          }
        });
        side.appendChild(card);
      });

      if (ad && pinnedPts.length && !opts.skipFit) {
        ad.fitPoints(pinnedPts);
      }
      if (ad) ad.relayout();
    }

    function consumePendingLocation() {
      var loc = pendingLoc;
      pendingLoc = null;
      updateSelectedLocUi();
      return loc;
    }

    function relayout() {
      if (adapter) adapter.relayout();
    }

    return {
      render: render,
      relayout: relayout,
      consumePendingLocation: consumePendingLocation,
      getPendingLocation: function () { return pendingLoc; }
    };
  })();

  // ============================================================================
  // 2. 오늘의 알림장 — 서울시 자치구 실시간 대기질(미세먼지) 배지
  // ============================================================================
  var NoticeAir = (function () {
    var cachedRows = null;
    var cachedAt = 0;

    function getSavedGu() {
      try { return localStorage.getItem(LOCAL_GU_KEY) || '강남구'; } catch (e) { return '강남구'; }
    }
    function saveGu(gu) {
      try { localStorage.setItem(LOCAL_GU_KEY, gu); } catch (e) {}
    }

    function gradeClass(grade) {
      if (grade === '좋음') return 'air-good';
      if (grade === '보통') return 'air-ok';
      if (grade === '나쁨') return 'air-bad';
      if (grade === '매우나쁨') return 'air-worst';
      return 'air-muted';
    }

    async function loadAirRows(force) {
      if (!force && cachedRows && (Date.now() - cachedAt < 300000)) return { ok: true, rows: cachedRows };
      var res = await fetchSeoulData('RealtimeCityAir', { start: 1, end: 25 });
      if (res && res.ok && Array.isArray(res.rows)) {
        cachedRows = res.rows;
        cachedAt = Date.now();
      }
      return res;
    }

    function populateGuSelect(sel) {
      if (!sel || sel.options.length > 0) return;
      Object.keys(SEOUL_GU_COORDS).forEach(function (gu) {
        var opt = document.createElement('option');
        opt.value = gu;
        opt.textContent = gu;
        sel.appendChild(opt);
      });
      sel.value = getSavedGu();
    }

    async function refresh(force) {
      var sel = document.getElementById('noticeAirGuSelect');
      var badge = document.getElementById('noticeAirBadge');
      var bigBadge = document.getElementById('nbAirBadge');
      if (sel) populateGuSelect(sel);
      var gu = sel ? sel.value : getSavedGu();

      var res = await loadAirRows(force);
      if (!res || !res.ok) {
        var msg = (res && res.error === 'SEOUL_KEY_NOT_CONFIGURED')
          ? '대기질: Vercel SEOUL_OPEN_API_KEY 등록 시 자동 표시'
          : '대기질 정보 불러오기 대기';
        if (badge) { badge.className = 'notice-air-pill air-muted'; badge.textContent = msg; }
        if (bigBadge) { bigBadge.style.display = 'none'; }
        return;
      }

      var row = (res.rows || []).find(function (r) { return r.MSRSTE_NM === gu; }) || res.rows[0];
      if (!row) return;
      var pm10 = row.PM10 != null ? row.PM10 : '-';
      var pm25 = row.PM25 != null ? row.PM25 : '-';
      var grade = row.IDEX_NM || '보통';
      var text = row.MSRSTE_NM + ' 대기질 ' + grade + ' (미세 ' + pm10 + ' · 초미세 ' + pm25 + ')';
      var cls = gradeClass(grade);

      if (badge) {
        badge.className = 'notice-air-pill ' + cls;
        badge.textContent = text;
      }
      if (bigBadge) {
        bigBadge.style.display = 'inline-flex';
        bigBadge.className = 'nb-air-badge ' + cls;
        bigBadge.textContent = text;
      }
    }

    function init() {
      var sel = document.getElementById('noticeAirGuSelect');
      var btn = document.getElementById('noticeAirRefreshBtn');
      if (sel) {
        populateGuSelect(sel);
        sel.addEventListener('change', function () {
          saveGu(sel.value);
          refresh(false);
        });
      }
      if (btn) btn.addEventListener('click', function () { refresh(true); });
      var noticeBtn = document.getElementById('noticeBtn');
      if (noticeBtn) noticeBtn.addEventListener('click', function () { refresh(false); });
      var bigBtn = document.getElementById('noticeBigBtn');
      if (bigBtn) bigBtn.addEventListener('click', function () { refresh(false); });
    }

    return { init: init, refresh: refresh };
  })();

  // ============================================================================
  // 3. 서울 공공데이터 × 카카오맵 수업 탐구 도구
  // ============================================================================
  var SeoulDataExplorer = (function () {
    var mapAdapter = null;
    var currentService = 'bikeList';
    var currentRows = [];
    var currentColumns = [];
    var bound = false;

    function defaultSubwayDate() {
      var d = new Date(Date.now() - 4 * 86400000);
      var y = d.getFullYear();
      var m = String(d.getMonth() + 1).padStart(2, '0');
      var day = String(d.getDate()).padStart(2, '0');
      return '' + y + m + day;
    }

    function setStatus(msg, isErr) {
      var el = document.getElementById('seoulDataStatus');
      if (!el) return;
      el.textContent = msg || '';
      el.className = 'sd-status' + (isErr ? ' is-err' : '');
    }

    function exportCsv() {
      if (!currentRows.length || !currentColumns.length) {
        setStatus('내보낼 데이터가 없습니다. 먼저 데이터를 조회해 주세요.', true);
        return;
      }
      var lines = [];
      lines.push(currentColumns.map(function (c) { return '"' + String(c.label).replace(/"/g, '""') + '"'; }).join(','));
      currentRows.forEach(function (r) {
        lines.push(currentColumns.map(function (c) {
          var v = r[c.key] != null ? String(r[c.key]) : '';
          return '"' + v.replace(/"/g, '""') + '"';
        }).join(','));
      });
      var blob = new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = '서울공공데이터_' + currentService + '_' + new Date().toISOString().slice(0, 10) + '.csv';
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 4000);
      setStatus('CSV 파일로 저장했습니다. 엑셀·통그라미에서 바로 열어 분석할 수 있습니다.');
    }

    function renderTable(rows, cols) {
      var wrap = document.getElementById('seoulDataTableWrap');
      if (!wrap) return;
      if (!rows.length) {
        wrap.innerHTML = '<div class="sd-empty">표시할 데이터가 없습니다.</div>';
        return;
      }
      var html = '<table class="sd-table"><thead><tr>';
      cols.forEach(function (c) { html += '<th>' + escapeHtml(c.label) + '</th>'; });
      html += '</tr></thead><tbody>';
      rows.slice(0, 100).forEach(function (r) {
        html += '<tr>';
        cols.forEach(function (c) {
          html += '<td>' + escapeHtml(r[c.key] != null ? r[c.key] : '') + '</td>';
        });
        html += '</tr>';
      });
      html += '</tbody></table>';
      wrap.innerHTML = html;
    }

    async function runQuery() {
      var dateInp = document.getElementById('seoulDataDateInput');
      var filterInp = document.getElementById('seoulDataFilterInput');
      var filterTxt = (filterInp && filterInp.value || '').trim().toLowerCase();
      var summaryEl = document.getElementById('seoulDataSummary');

      setStatus('서울 열린데이터 광장에서 실시간 데이터를 불러오는 중...');
      var opts = { start: 1, end: 80 };
      if (currentService === 'CardSubwayStatsNew') {
        opts.p1 = ((dateInp && dateInp.value) || defaultSubwayDate()).replace(/[^0-9]/g, '').slice(0, 8);
      }

      var res = await fetchSeoulData(currentService, opts);
      if (!res || !res.ok) {
        setStatus((res && res.message) || '데이터를 불러오지 못했습니다. Vercel 환경변수 SEOUL_OPEN_API_KEY를 확인해 주세요.', true);
        renderTable([], []);
        if (summaryEl) summaryEl.textContent = '';
        return;
      }

      var rows = res.rows || [];
      if (filterTxt) {
        rows = rows.filter(function (r) {
          return JSON.stringify(r).toLowerCase().indexOf(filterTxt) >= 0;
        });
      }

      var cols = [];
      var mapPoints = [];

      if (currentService === 'bikeList') {
        cols = [
          { key: 'stationName', label: '대여소명' },
          { key: 'parkingBikeTotCnt', label: '대여가능(대)' },
          { key: 'rackTotCnt', label: '거치대(개)' },
          { key: 'shared', label: '거치율(%)' }
        ];
        rows.forEach(function (r) {
          var lat = parseFloat(r.stationLatitude), lng = parseFloat(r.stationLongitude);
          if (!isNaN(lat) && !isNaN(lng) && lat > 30) {
            var cnt = parseInt(r.parkingBikeTotCnt, 10) || 0;
            mapPoints.push({
              lat: lat, lng: lng,
              color: cnt >= 8 ? '#bbf7d0' : (cnt >= 3 ? '#fef08a' : '#fecaca'),
              label: cnt + '대',
              title: r.stationName,
              html: '<b>' + escapeHtml(r.stationName) + '</b><br>대여 가능: <b>' + cnt + '대</b> / 거치대 ' + escapeHtml(r.rackTotCnt) + '개'
            });
          }
        });
        var totalBikes = rows.reduce(function (a, r) { return a + (parseInt(r.parkingBikeTotCnt, 10) || 0); }, 0);
        if (summaryEl) summaryEl.textContent = '조회 대여소 ' + rows.length + '곳 · 대여 가능 자전거 총 ' + totalBikes + '대';
      } else if (currentService === 'CardSubwayStatsNew') {
        cols = [
          { key: 'USE_YMD', label: '일자' },
          { key: 'SBWY_ROUT_LN_NM', label: '호선' },
          { key: 'SBWY_STNS_NM', label: '역명' },
          { key: 'GTON_TNOPE', label: '승차인원' },
          { key: 'GTOFF_TNOPE', label: '하차인원' }
        ];
        var sumOn = rows.reduce(function (a, r) { return a + (parseInt(r.GTON_TNOPE, 10) || 0); }, 0);
        var avgOn = rows.length ? Math.round(sumOn / rows.length) : 0;
        if (summaryEl) summaryEl.textContent = '조회 ' + rows.length + '개 역 · 총 승차 ' + sumOn.toLocaleString() + '명 · 역당 평균 ' + avgOn.toLocaleString() + '명';
      } else if (currentService === 'culturalEventInfo') {
        cols = [
          { key: 'CODENAME', label: '분류' },
          { key: 'TITLE', label: '행사명' },
          { key: 'GUNAME', label: '자치구' },
          { key: 'PLACE', label: '장소' },
          { key: 'DATE', label: '기간' },
          { key: 'IS_FREE', label: '요금' }
        ];
        rows.forEach(function (r) {
          // 서울 문화행사 API는 LOT=위도(37.x), LAT=경도(126.x)가 뒤바뀌어 내려오는 경우가 있어 안전 판별
          var a = parseFloat(r.LAT), b = parseFloat(r.LOT);
          var lat = null, lng = null;
          if (!isNaN(a) && !isNaN(b)) {
            if (a > 33 && a < 40 && b > 124 && b < 132) { lat = a; lng = b; }
            else if (b > 33 && b < 40 && a > 124 && a < 132) { lat = b; lng = a; }
          }
          if (lat && lng) {
            mapPoints.push({
              lat: lat, lng: lng,
              color: '#dbeafe',
              label: String(r.CODENAME || '행사').slice(0, 6),
              title: r.TITLE,
              html: '<b>' + escapeHtml(r.TITLE) + '</b><br>' + escapeHtml(r.PLACE) + '<br>' + escapeHtml(r.DATE) + ' (' + escapeHtml(r.IS_FREE) + ')'
            });
          }
        });
        if (summaryEl) summaryEl.textContent = '서울시 문화행사 ' + rows.length + '건 조회됨 (지도 표시 ' + mapPoints.length + '건)';
      } else if (currentService === 'RealtimeCityAir') {
        cols = [
          { key: 'MSRRGN_NM', label: '권역' },
          { key: 'MSRSTE_NM', label: '자치구' },
          { key: 'PM10', label: '미세먼지(PM10)' },
          { key: 'PM25', label: '초미세(PM2.5)' },
          { key: 'IDEX_NM', label: '통합등급' }
        ];
        rows.forEach(function (r) {
          var ll = SEOUL_GU_COORDS[r.MSRSTE_NM];
          if (ll) {
            var g = r.IDEX_NM || '';
            mapPoints.push({
              lat: ll[0], lng: ll[1],
              color: g === '좋음' ? '#bbf7d0' : (g === '보통' ? '#dbeafe' : '#fecaca'),
              label: r.MSRSTE_NM + ' ' + (r.PM10 || '-'),
              title: r.MSRSTE_NM,
              html: '<b>' + escapeHtml(r.MSRSTE_NM) + '</b> (' + escapeHtml(g) + ')<br>미세먼지(PM10): ' + escapeHtml(r.PM10) + '<br>초미세먼지(PM2.5): ' + escapeHtml(r.PM25)
            });
          }
        });
        if (summaryEl) summaryEl.textContent = '서울시 25개 자치구 실시간 대기환경 (' + rows.length + '개 측정소)';
      }

      currentRows = rows;
      currentColumns = cols;
      renderTable(rows, cols);
      setStatus('총 ' + rows.length + '건 조회 완료 (전체 ' + (res.totalCount || rows.length) + '건 중)');

      var mapEl = document.getElementById('seoulDataMapCanvas');
      if (mapEl) {
        if (!mapAdapter) mapAdapter = await createMapAdapter(mapEl, { lat: 37.5665, lng: 126.9780, level: 8 });
        if (mapAdapter) {
          mapAdapter.clearMarkers();
          mapPoints.slice(0, 60).forEach(function (pt) { mapAdapter.addMarker(pt); });
          if (mapPoints.length) mapAdapter.fitPoints(mapPoints.slice(0, 60));
          mapAdapter.relayout();
        }
      }
    }

    function bind() {
      if (bound) return;
      bound = true;
      var modal = document.getElementById('seoulDataModal');
      if (!modal) return;
      var dateInp = document.getElementById('seoulDataDateInput');
      if (dateInp && !dateInp.value) dateInp.value = defaultSubwayDate();

      modal.querySelectorAll('[data-sd-service]').forEach(function (btn) {
        btn.addEventListener('click', function () {
          currentService = btn.getAttribute('data-sd-service') || 'bikeList';
          modal.querySelectorAll('[data-sd-service]').forEach(function (b) { b.classList.toggle('active', b === btn); });
          var dateWrap = document.getElementById('seoulDataDateWrap');
          if (dateWrap) dateWrap.style.display = currentService === 'CardSubwayStatsNew' ? 'inline-flex' : 'none';
          runQuery();
        });
      });

      var qBtn = document.getElementById('seoulDataQueryBtn');
      if (qBtn) qBtn.addEventListener('click', runQuery);
      var csvBtn = document.getElementById('seoulDataCsvBtn');
      if (csvBtn) csvBtn.addEventListener('click', exportCsv);
      var filterInp = document.getElementById('seoulDataFilterInput');
      if (filterInp) filterInp.addEventListener('keydown', function (e) { if (e.key === 'Enter') runQuery(); });

      modal.querySelectorAll('.close-btn[data-modal="seoulDataModal"]').forEach(function (cb) {
        cb.addEventListener('click', function () { modal.classList.remove('show'); });
      });
    }

    function open() {
      var modal = document.getElementById('seoulDataModal');
      if (!modal) return;
      bind();
      modal.classList.add('show');
      setTimeout(runQuery, 80);
    }

    return { open: open };
  })();

  window.TongGeoServices = {
    loadConfig: loadConfig,
    createMapAdapter: createMapAdapter,
    parseMapTag: parseMapTag,
    encodeMapTag: encodeMapTag,
    fetchSeoulData: fetchSeoulData,
    MemoBoardMap: MemoBoardMap,
    NoticeAir: NoticeAir
  };
  window.SeoulDataExplorerModule = SeoulDataExplorer;

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () { NoticeAir.init(); });
  } else {
    NoticeAir.init();
  }
})();
