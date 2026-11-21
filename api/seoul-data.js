// 서울 열린데이터 광장 HTTPS/CORS 보안 프록시
//  - 브라우저에서 http://openapi.seoul.go.kr:8088 직접 호출 시 발생하는
//    Mixed Content(HTTPS→HTTP) 및 CORS 차단을 서버에서 중계해 해결합니다.
//  - 인증키는 Vercel 환경변수 SEOUL_OPEN_API_KEY 에만 보관합니다.
//    (키 발급: https://data.seoul.go.kr → 인증키 신청)

const ALLOWED_SERVICES = new Set([
  'RealtimeCityAir',              // 서울시 권역별 실시간 대기환경 (미세먼지/초미세먼지)
  'bikeList',                     // 서울시 공공자전거(따릉이) 실시간 대여정보
  'CardSubwayStatsNew',           // 서울시 지하철 호선별 역별 승하차 인원 정보
  'culturalEventInfo',            // 서울시 문화행사 정보
  'ListNecessariesPricesService'  // 서울시 생필품 농수축산물 가격 정보
]);

const CACHE_SECONDS = {
  RealtimeCityAir: 180,
  bikeList: 60,
  CardSubwayStatsNew: 3600,
  culturalEventInfo: 1800,
  ListNecessariesPricesService: 3600
};

const SEOUL_ENV_CANDIDATES = [
  'SEOUL_OPEN_API_KEY',
  'SEOUL_API_KEY',
  'SEOUL_OPEN_DATA_KEY',
  'SEOUL_DATA_API_KEY',
  'SEOUL_OPENAPI_KEY',
  'NEXT_PUBLIC_SEOUL_OPEN_API_KEY',
  'VITE_SEOUL_OPEN_API_KEY'
];

function cleanEnv(val) {
  return String(val || '').trim().replace(/^["']+|["']+$/g, '').trim();
}

function getSeoulApiKey() {
  for (let i = 0; i < SEOUL_ENV_CANDIDATES.length; i++) {
    const val = cleanEnv(process.env[SEOUL_ENV_CANDIDATES[i]]);
    if (val) return val;
  }
  return '';
}

function send(res, status, body, cacheSec) {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', cacheSec
    ? 'public, s-maxage=' + cacheSec + ', stale-while-revalidate=' + cacheSec
    : 'no-store, max-age=0');
  return res.status(status).json(body);
}

function safeSegment(val) {
  const s = String(Array.isArray(val) ? val[0] : (val || '')).trim();
  if (!s || s.length > 60) return '';
  if (/[\\/?#%]/.test(s)) return '';
  return s;
}

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') {
    res.setHeader('Allow', 'GET');
    return send(res, 405, { ok: false, error: 'METHOD_NOT_ALLOWED' });
  }

  const key = getSeoulApiKey();
  if (!key) {
    return send(res, 503, {
      ok: false,
      error: 'SEOUL_KEY_NOT_CONFIGURED',
      message: 'Vercel 환경변수 SEOUL_OPEN_API_KEY가 등록되지 않았습니다.'
    });
  }

  const q = req.query || {};
  const service = String(Array.isArray(q.service) ? q.service[0] : (q.service || '')).trim();
  if (!ALLOWED_SERVICES.has(service)) {
    return send(res, 400, { ok: false, error: 'UNSUPPORTED_SERVICE' });
  }

  const start = Math.max(1, parseInt(q.start, 10) || 1);
  const end = Math.min(start + 199, Math.max(start, parseInt(q.end, 10) || (start + 49)));

  const segs = [
    'http://openapi.seoul.go.kr:8088',
    encodeURIComponent(key),
    'json',
    encodeURIComponent(service),
    String(start),
    String(end)
  ];

  ['p1', 'p2', 'p3'].forEach(function (pk) {
    const v = safeSegment(q[pk]);
    if (v) segs.push(encodeURIComponent(v));
  });

  const targetUrl = segs.join('/') + '/';

  const controller = new AbortController();
  const timer = setTimeout(function () { controller.abort(); }, 10000);

  try {
    const upstream = await fetch(targetUrl, { signal: controller.signal });
    clearTimeout(timer);
    const rawText = await upstream.text();
    let parsed = null;
    try {
      parsed = JSON.parse(rawText);
    } catch (e) {
      return send(res, 502, {
        ok: false,
        error: 'INVALID_UPSTREAM_JSON',
        message: '서울 열린데이터 광장 응답이 JSON 형식이 아닙니다.'
      });
    }

    // 서울 열린데이터 광장 응답 구조 정규화
    // 일반: { [service]: { list_total_count, RESULT, row: [...] } }
    // bikeList 특수: { rentBikeStatus: { list_total_count, RESULT, row: [...] } }
    const rootKey = parsed[service]
      ? service
      : (parsed.rentBikeStatus ? 'rentBikeStatus' : Object.keys(parsed)[0]);
    const block = parsed[rootKey];

    if (parsed.RESULT && parsed.RESULT.CODE && parsed.RESULT.CODE !== 'INFO-000') {
      return send(res, 200, {
        ok: false,
        service: service,
        code: parsed.RESULT.CODE,
        message: parsed.RESULT.MESSAGE || '조회 결과가 없습니다.',
        totalCount: 0,
        rows: []
      });
    }

    if (!block || typeof block !== 'object') {
      return send(res, 200, { ok: false, service: service, totalCount: 0, rows: [], message: '데이터가 없습니다.' });
    }

    const resultObj = block.RESULT || {};
    if (resultObj.CODE && resultObj.CODE !== 'INFO-000') {
      return send(res, 200, {
        ok: false,
        service: service,
        code: resultObj.CODE,
        message: resultObj.MESSAGE || '데이터 조회 오류',
        totalCount: 0,
        rows: []
      });
    }

    const rows = Array.isArray(block.row) ? block.row : [];
    const totalCount = Number(block.list_total_count) || rows.length;

    return send(res, 200, {
      ok: true,
      service: service,
      totalCount: totalCount,
      rows: rows
    }, CACHE_SECONDS[service] || 120);
  } catch (err) {
    clearTimeout(timer);
    return send(res, 504, {
      ok: false,
      error: 'UPSTREAM_TIMEOUT',
      message: '서울 열린데이터 광장 서버 응답이 지연되고 있습니다.'
    });
  }
};
