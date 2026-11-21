// 공개용 클라이언트 환경 설정 전달 (Vercel 환경변수 연동)
//  - KAKAO_MAP_JS_KEY : 카카오맵 JavaScript SDK 키 (도메인 화이트리스트로 보호되는 공개용 키)
//  - SEOUL_OPEN_API_KEY : 서버 프록시(/api/seoul-data) 전용 (절대 클라이언트에 원본 키를 노출하지 않음)

const KAKAO_ENV_CANDIDATES = [
  'KAKAO_MAP_JS_KEY',
  'KAKAO_MAP_API_KEY',
  'KAKAO_JS_KEY',
  'KAKAO_API_KEY',
  'KAKAO_APP_KEY',
  'KAKAO_JAVASCRIPT_KEY',
  'NEXT_PUBLIC_KAKAO_MAP_JS_KEY',
  'NEXT_PUBLIC_KAKAO_MAP_KEY',
  'NEXT_PUBLIC_KAKAO_JS_KEY',
  'VITE_KAKAO_MAP_JS_KEY',
  'VITE_KAKAO_MAP_KEY'
];

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

function pickEnv(candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const name = candidates[i];
    const value = cleanEnv(process.env[name]);
    if (value) return { name: name, value: value };
  }
  return { name: null, value: '' };
}

module.exports = function handler(req, res) {
  if (req.method !== 'GET') {
    res.setHeader('Allow', 'GET');
    return res.status(405).json({ error: 'METHOD_NOT_ALLOWED' });
  }

  const kakao = pickEnv(KAKAO_ENV_CANDIDATES);
  const seoul = pickEnv(SEOUL_ENV_CANDIDATES);

  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  return res.status(200).json({
    kakaoMapJsKey: kakao.value,
    hasKakaoMapKey: Boolean(kakao.value),
    hasSeoulApi: Boolean(seoul.value),
    matchedKakaoEnv: kakao.name,
    matchedSeoulEnv: seoul.name
  });
};
