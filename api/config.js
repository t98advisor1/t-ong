// 공개용 클라이언트 환경 설정 전달 (Vercel 환경변수 연동)
//  - KAKAO_MAP_JS_KEY : 카카오맵 JavaScript SDK 키 (도메인 화이트리스트로 보호되는 공개용 키)
//  - SEOUL_OPEN_API_KEY : 서버 프록시(/api/seoul-data) 전용 (절대 클라이언트에 원본 키를 노출하지 않음)

function cleanEnv(val) {
  return String(val || '').trim().replace(/^["']+|["']+$/g, '').trim();
}

module.exports = function handler(req, res) {
  if (req.method !== 'GET') {
    res.setHeader('Allow', 'GET');
    return res.status(405).json({ error: 'METHOD_NOT_ALLOWED' });
  }

  const kakaoMapJsKey = cleanEnv(process.env.KAKAO_MAP_JS_KEY);
  const hasSeoulApi = Boolean(cleanEnv(process.env.SEOUL_OPEN_API_KEY));

  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  return res.status(200).json({
    kakaoMapJsKey: kakaoMapJsKey,
    hasKakaoMapKey: Boolean(kakaoMapJsKey),
    hasSeoulApi: hasSeoulApi
  });
};
