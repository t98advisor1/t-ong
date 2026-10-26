# T-ong 교무수첩

담임·교과교사를 위한 디지털 교무수첩입니다. 기존 GAS 자리배치 앱을 바탕으로 UI와 기능을 확장했으며, 순수 HTML/CSS/바닐라 JavaScript로 구성합니다.

- 소스 버전: **5.17.0-preview.73**
- 운영 주소: https://t-ong.vercel.app/
- GitHub: https://github.com/t98advisor1/t-ong
- 주요 기능: 자리배치·출결·학생 기록·칭찬·교실운영·일정·즐겨찾기·외부 게임
- preview.73: 사이드바 **AI 비서 채팅**(사용자 본인 Claude API 키) 추가, API 키 발급 안내
- preview.72: 상단 **🤖 AI 메뉴**와 **AI 활용 안내** 추가, 자리배치 우측 메뉴 접기
- preview.71: 사이드 **AI 도우미**(구독 계정 연결) 추가, Google 캘린더 연결을 버튼 한 번으로 간소화
- preview.70: 게임 메뉴에 **온몸으로 찰칵! 포즈 챌린지** 추가

학번 가로·세로순 배치, 같은 성별 우선, 자리 복원·교탁 시점, 학생 팝업 출결 등록, 성장형 칭찬, 홈 백업, 일정 드래그·삭제와 Google 일정 전체 가져오기를 포함합니다.

## 구조

```
index.html            화면과 메뉴
style.css             테마 및 반응형 레이아웃
app.js                앱 기능과 로컬 저장 로직
manifest.webmanifest  PWA 설치 정보
sw.js                 서비스 워커 (오프라인 실행)
icon.svg / icon-maskable.svg
vercel.json           캐시 헤더
assets/               공개 테마·아이콘 이미지
guide/                사용 설명서 이미지
api/usage.js          익명 방문 통계 API
vendor/               AI 비서용 Anthropic SDK 브라우저 번들 (@anthropic-ai/sdk 0.132.1, esbuild로 묶음)
tools/                검사·공개 파일 미리보기·배포 준비 도구
```

## 데이터 저장

학생·학급·좌석·출결·상담 등 교무 자료는 기본적으로 현재 브라우저의 `localStorage`에 저장됩니다. 자리배치의 주요 저장 키는 `seatArrangement_v3`입니다.

익명 방문 통계만 Upstash Redis에 저장합니다. Today는 당일 순 방문 브라우저 수, Total은 누적 접속 횟수이며 통계에 학생 자료는 포함하지 않습니다. Google 동기화를 직접 켜면 선택한 일정 항목만 본인 Google 계정에도 저장됩니다. 교무 자료 전체를 기기 간 동기화하는 기능은 아닙니다.

- 브라우저마다 데이터가 따로입니다 (같은 PC의 Chrome ↔ Edge도 별개)
- 브라우저 "쿠키 및 사이트 데이터 삭제" 시 **전부 사라집니다**
- 시크릿 모드에서 쓴 데이터는 창을 닫으면 사라집니다

AI 비서는 사용자 본인의 Claude API 키(`tongAiApiKey_v1`)를 이 브라우저에만 저장하고, 대화는 브라우저에서 Anthropic API로 바로 보냅니다. 키와 대화 내용은 T-ong 서버와 데이터 백업 코드에 포함되지 않습니다.

> **정기적으로 `홈 → 데이터 백업`에서 자료를 백업해 따로 보관하세요.** 학생 원본 기록·백업, 개인 일정, 환경변수, OAuth 토큰, Android 서명키는 GitHub에 올리지 않습니다.

### 기존 GAS 버전 데이터 옮기기

도메인이 다르므로 자동 이전되지 않습니다.

1. 기존 GAS 주소에서 `설정 → 데이터 백업 → 코드 생성 → 복사`
2. 새 Vercel 주소에서 `설정 → 데이터 백업 → 가져오기`에 붙여넣기

## 로컬에서 실행

`file://`로 열면 절대경로가 동작하지 않으므로 반드시 HTTP 서버로 띄웁니다.

```bash
node tools/preview-server.js
```

이후 http://localhost:4321 접속. **localhost에서는 서비스 워커를 등록하지 않습니다** — 캐시가 수정사항을
가리는 것을 막기 위해서입니다.

이 미리보기 서버는 공개 앱 파일만 제공합니다. 방문 통계 API는 실행하지 않으며, Google OAuth 및 서버 API는 운영 환경과 별도 설정이 필요할 수 있습니다.

## 빌드 (난독화)

배포본은 `npm run build`로 만들어지는 `dist/` 폴더입니다.

```bash
npm install
npm run build
node tools/test-pose-game-link.js
node tools/test-pose-game-link.js --built
python -m http.server 4322 --directory dist   # 빌드 결과 확인
```

`build.js`가 하는 일:

| | |
|---|---|
| `app.js` | **javascript-obfuscator로 난독화** (식별자 16진수화, 문자열 배열 base64, 제어흐름 평탄화, 더미코드 삽입) |
| `style.css` | 주석 제거 + 공백 압축 |
| `index.html` | 주석 제거 + 들여쓰기 제거 |
| `sw.js` | **캐시 버전을 내용 해시로 자동 교체** — 수동으로 버전을 올릴 필요가 없습니다 |
| `guide/`, `assets/`, 아이콘, manifest | 그대로 복사 |

> 개발자 도구를 방해하는 옵션(`debugProtection`, `selfDefending`)은 일부러 켜지 않았습니다.
> 정상 사용자에게 피해가 갈 수 있고 보호 효과도 크지 않습니다.
>
> **난독화는 복사를 어렵게 만들 뿐 완전히 막지는 못합니다.** 브라우저에 전달된 코드는 원리상 항상
> 복원할 수 있습니다. 알고리즘을 확실히 숨기려면 서버로 옮겨야 하지만, 그러면 오프라인 동작과
> "서버에 아무것도 보내지 않는다"는 장점을 잃습니다.

## Vercel 배포

```bash
npx vercel deploy --prod
```

`vercel.json`에 `buildCommand: npm run build`, `outputDirectory: dist`가 설정되어 있어
**Vercel이 배포할 때마다 난독화 빌드를 자동으로 실행**합니다. 원본 소스는 업로드되지만 서빙되지 않습니다.

현재 운영 주소: **https://t-ong.vercel.app**

기존 **https://seattong.vercel.app**은 별도 서비스이며 이 프로젝트의 수정·배포 대상이 아닙니다. 배포 전에 `vercel project inspect --non-interactive`로 연결된 대상이 `t-ong`인지 확인합니다. `tools/prepare-tong-public-release.js`는 이미 연결된 T-ong 프로젝트 정보를 검사하고 공개 파일만 새 배포 준비 폴더로 복사합니다.

Android 앱은 같은 운영 웹사이트를 여는 TWA 방식입니다. `.well-known/assetlinks.json`의 인증서 지문은 공개 검증 정보이며 비밀 서명키가 아닙니다. 현재 테스트용 인증과 Play 정식 출시용 인증은 구분해야 합니다. 웹 메뉴 업데이트와 Google Play 업로드·심사·출시는 별도 절차입니다.

`dist/`, `tmp/`, `node_modules/`, 개인 문서·키·환경변수는 커밋하지 않습니다. 설명서의 v67 실습 이미지는 가상 학생·가상 일정으로 생성했습니다. 일부 기존 브라우저 회귀 검사는 제작 PC의 Playwright 경로를 사용하므로 다른 PC에서는 실행 환경을 맞춰야 합니다.

> 참고: `J:\내 드라이브`는 Google Drive 가상 드라이브라 `npm install`이 실패할 수 있습니다(EBADF/EPERM).
> 로컬에서 빌드해야 한다면 다른 경로에 설치한 뒤 `NODE_PATH`로 연결하세요. Vercel 빌드 서버에서는
> 문제없이 동작합니다.

## 알아둘 점

- 엑셀 기능은 `cdn.sheetjs.com`의 SheetJS를 사용합니다. 학교 망에서 CDN이 차단되면
  엑셀 업로드/다운로드가 동작하지 않으므로, 그 경우 라이브러리를 프로젝트 안에 내려받아
  `index.html`의 `<script src>`를 로컬 경로로 바꾸면 됩니다.
- 정적 배포이므로 **주소를 아는 사람은 누구나 접속**할 수 있습니다. 다만 데이터는 각자
  브라우저에만 저장되므로 다른 사람이 내 명렬표를 볼 수는 없습니다.
