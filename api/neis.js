// NEIS 교육정보 개방 포털 프록시
//  - 인증키 없이 호출하면 NEIS 는 요청당 5건(샘플)만 돌려줍니다. 시간표는 하루 6~7교시,
//    학사일정은 연 200건 안팎이라 인증키가 있어야 전체가 조회됩니다.
//  - 인증키를 브라우저에 노출하지 않도록 서버 환경변수 NEIS_API_KEY 에만 두고 여기서 붙입니다.
//  - Vercel > Project > Settings > Environment Variables 에 NEIS_API_KEY 를 등록하세요.
//    (키 발급: https://open.neis.go.kr → 인증키 신청)
const DATASETS = new Set([
  'schoolInfo',
  'classInfo',
  'hisTimetable',
  'misTimetable',
  'elsTimetable',
  'spsTimetable',
  'SchoolSchedule',
  'mealServiceDietInfo'
]);

const PARAMS = new Set([
  'ATPT_OFCDC_SC_CODE', 'SD_SCHUL_CODE', 'SCHUL_NM', 'SCHUL_KND_SC_NM',
  'AY', 'SEM', 'GRADE', 'CLASS_NM',
  'ALL_TI_YMD', 'TI_FROM_YMD', 'TI_TO_YMD',
  'AA_YMD', 'AA_FROM_YMD', 'AA_TO_YMD',
  'MLSV_YMD', 'MLSV_FROM_YMD', 'MLSV_TO_YMD', 'MMEAL_SC_CODE'
]);

// 학교·학급 목록은 거의 바뀌지 않고, 시간표·급식·학사일정은 하루에도 갱신될 수 있습니다.
const CACHE_SECONDS = { schoolInfo: 86400, classInfo: 21600 };

function send(res, status, body, cacheSeconds) {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', cacheSeconds
    ? 'public, s-maxage=' + cacheSeconds + ', stale-while-revalidate=' + cacheSeconds
    : 'no-store, max-age=0');
  res.status(status).json(body);
}

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') { res.setHeader('Allow', 'GET'); return send(res, 405, { error: 'METHOD_NOT_ALLOWED' }); }

  // 환경변수에 붙여넣을 때 섞이기 쉬운 앞뒤 공백·줄바꿈·따옴표는 떼어 냅니다.
  const key = String(process.env.NEIS_API_KEY || '').trim().replace(/^["']+|["']+$/g, '').trim();
  if (!key) return send(res, 503, { error: 'NEIS_KEY_NOT_CONFIGURED' });

  const query = req.query || {};
  const setName = String(query.set || '');
  if (!DATASETS.has(setName)) return send(res, 400, { error: 'UNSUPPORTED_DATASET' });

  const url = new URL('https://open.neis.go.kr/hub/' + setName);
  url.searchParams.set('KEY', key);
  url.searchParams.set('Type', 'json');
  url.searchParams.set('pIndex', String(Math.max(1, parseInt(query.pIndex, 10) || 1)));
  url.searchParams.set('pSize', String(Math.min(1000, Math.max(1, parseInt(query.pSize, 10) || 100))));
  for (const name of Object.keys(query)) {
    if (!PARAMS.has(name)) continue;
    const value = String(Array.isArray(query[name]) ? query[name][0] : query[name]).trim();
    if (value && value.length <= 60) url.searchParams.set(name, value);
  }

  try {
    // 요청 헤더를 따로 붙이지 않습니다. NEIS 는 Accept: application/json 이 있으면 HTTP 500 을 돌려줍니다.
    const upstream = await fetch(url.toString());
    if (!upstream.ok) return send(res, 502, { error: 'NEIS_HTTP_' + upstream.status });
    const data = await upstream.json();
    // 키가 잘못됐을 때는 캐시하지 않습니다 (키를 고친 뒤에도 오류가 남는 일을 막기 위해).
    const code = data && data.RESULT && data.RESULT.CODE;
    const failed = code && code !== 'INFO-200';
    // NEIS 가 키를 거부하면, 저장된 값이 인증키 모양인지(값 자체는 드러내지 않고) 알려 줍니다.
    // NEIS 인증키는 open.neis.go.kr 에서 발급하는 32자리 영문·숫자입니다.
    if (code === 'ERROR-290') data.keyCheck = { length: key.length, lettersAndDigitsOnly: /^[0-9A-Za-z]+$/.test(key) };
    return send(res, 200, data, failed ? 0 : (CACHE_SECONDS[setName] || 600));
  } catch (error) {
    return send(res, 502, { error: 'NEIS_UNAVAILABLE' });
  }
};
