/* 매뉴얼 영상 장(chapter) 목록 — 장마다 자막과 조작 순서를 적습니다.
 * id는 영상 파일 이름이 되고, title·summary는 매뉴얼 페이지에도 그대로 쓰입니다.
 */
const fs = require('fs');
const path = require('path');
const { shapes } = require('./demo-lib');

const SAMPLES = path.join(__dirname, 'samples');

async function openFromHub(page, d, inputId, file) {
  await d.click('#sbOpenFileHubBtn', 900);
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), d.click('.sb-filehub-card:has(#' + inputId + ')', 100)]);
  await chooser.setFiles(path.join(SAMPLES, file));
}

module.exports = [
  {
    id: '01-screen',
    title: '화면 구성과 도구 팔레트',
    summary: '위쪽 두 줄이 도구 팔레트입니다. 숨기면 화면 전체가 칠판이 됩니다.',
    async run(page, d) {
      await d.caption('스마트 판서는 <b>위쪽 두 줄</b>의 도구 팔레트와 넓은 칠판으로 되어 있습니다', 2600);
      await d.caption('첫 줄: <b>사용설명서 · 전체화면 · 노트함 · 열기 · 모드 · 저장</b>', 0);
      await d.highlightBox(4, 2, 1272, 40, 2600);
      await d.caption('둘째 줄: <b>선택 · 펜 · 도형 · 레이저 · 지우개 · 색 · 굵기 · 쪽 넘기기</b>', 0);
      await d.highlightBox(4, 44, 1272, 42, 2600);
      await d.caption('버튼 옆 작은 글자는 <b>단축키</b>입니다 — 색 동그라미의 숫자도 단축키', 0);
      await d.highlight('#sbPenMenuBtn', 1300);
      await d.highlight('.sb-color-row, .sb-color-dot[data-key="1"]', 1500).catch(() => d.wait(1500));
      await d.caption('<b>▲ 숨김</b>을 누르면 팔레트가 올라가 화면을 넓게 씁니다', 900);
      await d.click('#sbFoldPaletteBtn', 1600);
      await d.caption('다시 펼칠 때는 <b>Tab</b> 키나 위쪽 <b>▼ 펼치기</b> 버튼', 600);
      await d.press('Tab', 'Tab');
      await d.wait(900);
      await d.caption('<b>자동숨김</b>을 켜면 판서를 시작할 때 팔레트가 저절로 숨겨집니다', 900);
      await d.click('#sbAutoHideBtn', 700);
      await d.stroke(shapes.parabola(640, 520, 220, 260));
      await d.wait(900);
      await d.caption('마우스나 S펜을 <b>화면 맨 위</b>로 가져가면 다시 나타납니다', 500);
      await d.moveTo(640, 2, 30);
      await d.wait(1600);
      await d.caption('수업 화면을 크게 보여 줄 때는 <b>전체화면</b> (F 키)', 0);
      await d.highlight('#sbFullscreenBtn', 2200);
    }
  },
  {
    id: '02-pen',
    title: '펜으로 쓰기 · 색 · 굵기 · 되돌리기',
    summary: '펜 종류와 색, 굵기를 바꾸고 잘못 쓴 것은 되돌립니다.',
    async run(page, d) {
      await d.caption('<b>펜</b>이 기본으로 켜져 있어 바로 쓸 수 있습니다');
      await d.stroke(shapes.parabola(330, 470, 170, 230));
      await d.caption('색 동그라미를 누르면 <b>펜 색</b>이 바뀝니다', 800);
      await d.click('.sb-color-dot[data-color="#3b82f6"]');
      await d.stroke(shapes.circle(700, 330, 90));
      await d.caption('<b>+ / −</b> 버튼으로 굵기를 조절합니다', 800);
      await d.click('#sbSizeIncBtn', 350); await d.click('#sbSizeIncBtn', 350); await d.click('#sbSizeIncBtn', 500);
      await d.stroke(shapes.wave(560, 560, 330, 40));
      await d.caption('<b>펜 종류</b> 메뉴에서 만년필 · 볼펜 · 연필 · 붓펜 · <b>형광펜</b>을 고릅니다', 800);
      await d.click('#sbPenMenuBtn', 900);
      await d.click('#sbPenMenu [data-pen-style="highlighter"]', 700);
      await d.click('.sb-color-dot[data-color="#facc15"]', 500);
      await d.stroke([[600, 330], [800, 330]], 6);
      await d.caption('키보드 숫자 <b>1~6</b>으로도 색을 바로 바꿀 수 있습니다', 600);
      await d.press('p', 'P  펜');
      await d.press('2', '2  초록');
      await d.stroke([[230, 620], [520, 620]]);
      await d.caption('잘못 썼다면 <b>되돌리기</b> (Ctrl+Z)', 700);
      await d.click('#sbUndoBtn', 700); await d.click('#sbUndoBtn', 900);
      await d.caption('<b>다시 실행</b> (Ctrl+Y)으로 되살릴 수도 있습니다', 600);
      await d.click('#sbRedoBtn', 1600);
    }
  },
  {
    id: '03-shapes',
    title: '도형 · 직선 · 텍스트',
    summary: '도형 메뉴나 단축키 R·C·L·T로 반듯한 도형과 글자를 넣습니다.',
    async run(page, d) {
      await d.caption('<b>도형</b> 메뉴에 직선 · 사각형 · 원 · 텍스트가 있습니다', 800);
      await d.click('#sbMoreToolsBtn', 1000);
      await d.click('#sbMoreToolsMenu [data-sb-tool="rect"]', 500);
      await d.caption('끌어서 그리면 <b>반듯한 사각형</b>이 됩니다', 300);
      await d.drag(140, 200, 420, 380);
      await d.caption('단축키 <b>C</b> 원 · <b>L</b> 직선', 300);
      await d.press('c', 'C  원');
      await d.drag(520, 190, 760, 400);
      await d.press('l', 'L  직선');
      await d.drag(860, 300, 1160, 300);
      await d.press('3', '3  파랑');
      await d.drag(860, 230, 1160, 380);
      await d.caption('<b>T</b>를 누르고 화면을 누르면 글자를 입력합니다', 300);
      await d.setSize(8);
      await d.press('t', 'T  텍스트');
      await d.tapAt(200, 500, 400);
      await page.keyboard.type('정답 : x = 3', { delay: 90 });
      await page.keyboard.press('Enter');
      await d.wait(500);
      await d.caption('<b>P</b>를 누르면 다시 펜으로 돌아옵니다', 300);
      await d.press('p', 'P  펜');
      await d.setSize(5);
      await d.stroke([[200, 560], [470, 560]]);
      await d.caption('펜으로 그리다가 끝에서 <b>꾸욱 멈추면</b> 반듯한 도형으로 바뀝니다', 300);
      const wob = (pts, a) => pts.map(([x, y], i) => [x + Math.sin(i * 1.7) * a, y + Math.cos(i * 1.3) * a]);
      const holdStroke = async (pts) => { await d.moveTo(pts[0][0], pts[0][1], 10); await page.mouse.down(); for (const q of pts) { await page.mouse.move(q[0], q[1]); await page.waitForTimeout(12); } await page.waitForTimeout(900); await page.mouse.up(); await d.wait(500); };
      const edge = (a, b, n = 14) => Array.from({ length: n }, (_, i) => [a[0] + (b[0] - a[0]) * i / n, a[1] + (b[1] - a[1]) * i / n]);
      await holdStroke(wob(shapes.circle(640, 560, 70, 40), 5));
      await holdStroke(wob([].concat(edge([800, 490], [960, 495]), edge([960, 495], [955, 630]), edge([955, 630], [798, 625]), edge([798, 625], [802, 492]), [[802, 492]]), 3));
      await holdStroke(wob([].concat(edge([1050, 630], [1130, 490]), edge([1130, 490], [1210, 632]), edge([1210, 632], [1052, 628]), [[1052, 628]]), 3));
      await d.caption('별도 한 붓으로 그리고 <b>꾸욱</b> — 반듯한 별이 됩니다', 300);
      const starV = []; for (let i = 0; i <= 5; i++) { const ang = -Math.PI / 2 + i * 4 * Math.PI / 5; starV.push([1130 + 80 * Math.cos(ang), 300 + 80 * Math.sin(ang)]); }
      let starPts = []; for (let i = 0; i < 5; i++) starPts = starPts.concat(edge(starV[i], starV[i + 1])); starPts.push(starV[5]);
      await holdStroke(wob(starPts, 3));
      await d.wait(1400);
    }
  },
  {
    id: '04-stamps',
    title: '번호 스탬프 · 정답 가림막',
    summary: '①②③ 번호 스탬프를 찍고, 정답 가림막을 눌러 하나씩 공개합니다.',
    async run(page, d) {
      await d.caption('<b>S</b> 원형 번호 스탬프 — 누를 때마다 다음 번호가 찍힙니다', 300);
      await d.press('s', 'S  원형 스탬프');
      for (const [x, y] of [[170, 200], [170, 300], [170, 400], [170, 500]]) await d.tapAt(x, y, 450);
      await d.caption('잘못 찍어 <b>되돌리면</b>, 빠진 번호부터 다시 찍힙니다', 300);
      await d.press('Control+z', 'Ctrl + Z');
      await d.wait(500);
      await d.tapAt(320, 500, 900);
      await d.caption('중간 번호를 <b>지워도</b> 그 번호를 다음에 다시 씁니다', 300);
      await d.press('e', 'E  지우개');
      await d.stroke([[150, 300], [190, 300]], 4);
      await d.press('s', 'S  원형 스탬프');
      await d.tapAt(320, 300, 1100);
      await d.caption('<b>Q</b> 사각 번호 스탬프도 같은 번호 순서를 이어 씁니다', 300);
      await d.press('q', 'Q  사각 스탬프');
      await d.tapAt(470, 200, 500); await d.tapAt(470, 300, 800);
      await d.caption('번호를 1번부터 새로 시작하려면 도형 메뉴의 <b>스탬프 번호 초기화</b>', 800);
      await d.click('#sbMoreToolsBtn', 900);
      await d.click('#sbStampResetBtn', 700);
      await page.keyboard.press('Escape');
      await d.wait(300);
      await d.caption('정답을 가리고 싶을 때는 <b>정답 가림막</b>', 300);
      await d.typeText(760, 290, '정답 : 2, 4, 6');
      await d.click('#sbMoreToolsBtn', 900);
      await d.click('#sbMoreToolsMenu [data-sb-tool="cover"]', 400);
      await d.drag(720, 260, 1100, 350);
      await d.wait(600);
      await d.caption('수업 중에 가림막을 <b>누르면 정답이 열리고</b>, 다시 누르면 닫힙니다', 300);
      await page.keyboard.press('Escape');
      await d.wait(400);
      await d.tapAt(900, 305, 1500);
      await d.tapAt(900, 305, 1500);
    }
  },
  {
    id: '05-select',
    title: '선택·이동 · 레이저 · 지우개',
    summary: '쓴 것을 옮기고, 레이저로 가리키고, 지우개로 지웁니다.',
    async run(page, d) {
      await d.stroke(shapes.circle(330, 380, 80));
      await d.stroke(shapes.wave(560, 380, 320, 60));
      await d.caption('<b>V</b> 선택 — 쓴 것을 눌러 잡고 원하는 곳으로 옮깁니다', 300);
      await d.press('v', 'V  선택');
      await d.tapAt(250, 380, 500);
      await d.drag(330, 380, 520, 560);
      await d.wait(700);
      await d.caption('<b>Z</b> 레이저 포인터 — 잠깐 보였다가 사라져 설명할 때 좋습니다', 300);
      await d.press('z', 'Z  레이저');
      await d.stroke(shapes.circle(720, 380, 120, 60), 5);
      await d.stroke([[600, 520], [900, 520]], 5);
      await d.wait(900);
      await d.caption('<b>E</b> 지우개 — 문지른 선이 통째로 지워집니다', 300);
      await d.press('e', 'E  지우개');
      await d.stroke([[700, 300], [700, 460]], 6);
      await d.wait(700);
      await d.caption('갤럭시 <b>S펜 옆 버튼</b>을 누르면 지우개 ↔ 펜이 바로 바뀝니다', 2400);
      await d.caption('잘못 지웠다면 <b>Esc</b> — 방금 한 일을 되돌립니다 (Ctrl+Z와 같음)', 300);
      await d.press('Escape', 'Esc  되돌리기');
      await d.wait(900);
      await d.caption('<b>P</b>를 누르면 다시 펜 · <b>Backspace</b>는 판서를 저장하고 홈으로', 300);
      await d.press('p', 'P  펜');
      await d.wait(700);
      await d.press('s', 'S  원형 스탬프');
      for (const x of [880, 960, 1040]) await d.tapAt(x, 560, 300);
      await d.press('p', 'P  펜');
      await d.caption('한꺼번에 지울 때는 <b>지우개 ▾</b> 메뉴 (지우개를 한 번 더 눌러도 열림)', 800);
      await d.click('#sbEraserToolBtn .sb-caret', 1000);
      await d.caption('펜 글씨 <b>Shift+E</b> · 스탬프 <b>Shift+S</b> · 전체 필기 <b>Ctrl+E</b>', 600);
      await d.click('#sbEraseAllPenBtn', 1500);
      await d.caption('실수로 지웠다면 바로 <b>되돌리기</b> (Ctrl+Z) — 한꺼번에 되살아납니다', 300);
      await d.press('Control+z', 'Ctrl + Z');
      await d.wait(1500);
    }
  },
  {
    id: '06-pages',
    title: '쪽 추가 · 배경 바꾸기',
    summary: '+쪽으로 새 칠판을 만들고, 칠판·백지·모눈 배경을 고릅니다.',
    async run(page, d) {
      await d.stroke(shapes.parabola(640, 480, 200, 230));
      await d.caption('<b>+쪽</b>을 누르면 새 칠판 쪽이 생깁니다', 800);
      await d.click('#sbAddPageBtn', 900);
      await d.stroke(shapes.circle(640, 380, 100));
      await d.caption('<b>◀ ▶</b> 버튼이나 화살표 키로 쪽을 넘깁니다', 800);
      await d.click('#sbPrevPageBtn', 1200);
      await d.click('#sbNextPageBtn', 1200);
      await d.caption('<b>설정</b>에서 칠판 · 백지 · <b>모눈</b> 배경을 고릅니다', 800);
      await d.click('#sbBoardMoreBtn', 1000);
      await d.click('[data-sb-bg="grid"]', 900);
      await page.keyboard.press('Escape');
      await d.wait(300);
      await d.press('3', '3  파랑');
      await d.stroke([[340, 380], [940, 380]]);
      await d.stroke([[640, 160], [640, 600]]);
      await d.caption('단축키 <b>W</b> 백지 · <b>K</b> 칠판', 300);
      await d.press('w', 'W  백지');
      await d.wait(900);
      await d.press('k', 'K  칠판');
      await d.wait(1200);
    }
  },
  {
    id: '07-pdf',
    title: 'PDF 열어서 판서하기',
    summary: '교과서·활동지 PDF를 열어 그 위에 바로 씁니다.',
    async run(page, d) {
      await d.caption('<b>열기</b>를 누르고 <b>PDF 열기</b>에서 파일을 고릅니다', 300);
      await openFromHub(page, d, 'sbPdfInput', 'worksheet.pdf');
      await d.wait(2600);
      await d.caption('PDF 위에 바로 판서할 수 있습니다', 300);
      await d.stroke([[355, 168], [475, 168]]);
      await d.stroke(shapes.circle(652, 143, 42, 40));
      await d.click('.sb-color-dot[data-color="#3b82f6"]', 400);
      await d.typeText(300, 232, '꼭짓점 (2, −1)');
      await d.wait(600);
      await d.caption('<b>▶</b> 버튼으로 다음 쪽 — 쪽마다 판서가 따로 저장됩니다', 800);
      await d.click('#sbNextPageBtn', 1400);
      await d.typeText(300, 232, 'a가 커지면 폭이 좁아진다');
      await d.click('#sbPrevPageBtn', 1600);
    }
  },
  {
    id: '08-pptx',
    title: 'PPTX 발표 자료 · 클릭 애니메이션',
    summary: 'PPTX를 열면 클릭 애니메이션 순서대로 하나씩 보여 줍니다.',
    async run(page, d) {
      await d.caption('<b>열기 → PPTX 열기</b>로 발표 자료를 엽니다', 300);
      await openFromHub(page, d, 'sbPptxInput', 'lesson.pptx');
      await d.wait(2600);
      await d.caption('<b>다음 ▶</b> 또는 <b>Space · →</b> 키로 다음 장면', 800);
      await d.click('#sbPptxNextActBtn', 1200);
      await d.caption('파워포인트에서 넣은 <b>클릭 애니메이션</b>이 순서대로 나타납니다', 300);
      await d.press('Space', 'Space');
      await d.wait(700);
      await d.press('ArrowRight', '→');
      await d.wait(700);
      await d.press('ArrowRight', '→');
      await d.wait(700);
      await d.caption('슬라이드 위에도 바로 판서할 수 있습니다', 300);
      await d.stroke(shapes.circle(1010, 380, 150, 50));
      await d.caption('<b>전체</b>를 누르면 애니메이션을 건너뛰고 모두 보여 줍니다', 800);
      await d.click('#sbPptxNextActBtn', 900);
      await d.click('#sbPptxShowAllBtn', 1800);
    }
  },
  {
    id: '09-web',
    title: '유튜브 · 웹 사이트 위에 판서',
    summary: '유튜브나 수업용 웹 사이트를 보드 안에 띄우고 그 위에 씁니다.',
    async routes(ctx) {
      const html = fs.readFileSync(path.join(SAMPLES, 'web-demo.html'));
      await ctx.route('https://example.com/**', r => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: html }));
    },
    async run(page, d) {
      await d.caption('<b>열기</b> 창 아래쪽에 유튜브 · 웹 주소를 넣고 <b>링크 열기</b>', 300);
      await d.click('#sbOpenFileHubBtn', 900);
      await d.click('#sbFileHubUrlInput', 300);
      await page.keyboard.type('https://example.com/earth', { delay: 45 });
      await d.wait(400);
      await page.keyboard.press('Enter');
      await d.wait(2400);
      await d.caption('웹 화면 위에 <b>펜 필기</b>를 할 수 있습니다', 300);
      await d.stroke(shapes.circle(410, 512, 90, 50));
      await d.stroke([[620, 470], [760, 470]]);
      await d.caption('<b>① 원형스탬프</b>로 짚어 가며 설명합니다', 800);
      await d.click('#sbWebCircleStampBtn', 600);
      await d.tapAt(560, 470, 600); await d.tapAt(860, 470, 600); await d.tapAt(640, 330, 800);
      await d.caption('영상 재생·스크롤처럼 웹을 직접 다룰 때는 <b>웹 조작</b>', 800);
      await d.click('#sbWebInteractToggleBtn', 1200);
      await d.caption('지오지브라 · 데스모스 같은 자주 쓰는 사이트는 <b>바로가기 버튼</b>으로 엽니다', 0);
      await d.highlight('.sb-web-bar, #sbWebUrlInput', 2400).catch(() => d.wait(2400));
      await d.hideCaption();
      await d.click('#sbWebCloseBtn', 1200);
    }
  },
  {
    id: '13-zoom',
    title: '확대 · 손가락으로 이동',
    summary: '줌잇처럼 화면 한 부분을 크게 보여 주고, 태블릿에서는 손가락으로 확대·이동합니다.',
    contextOptions: { hasTouch: true },
    async run(page, d) {
      await d.typeText(110, 200, 'y = x² − 4x + 3');
      await d.stroke(shapes.parabola(330, 560, 170, 260));
      await d.press('3', '3  파랑');
      await d.typeText(900, 520, '꼭짓점 (2, −1)');
      await d.caption('<b>Ctrl + 1</b> 확대 — 커서를 움직이면 그쪽을 크게 보여 줍니다', 300);
      await d.moveTo(640, 360, 8);
      await d.press('Control+1', 'Ctrl + 1  확대');
      await d.moveTo(150, 250, 40);
      await d.wait(600);
      await d.moveTo(1150, 560, 50);
      await d.wait(700);
      await d.caption('마우스 <b>휠</b>로 배율을 바꿉니다', 300);
      await page.mouse.wheel(0, -120); await d.wait(400);
      await page.mouse.wheel(0, -120); await d.wait(600);
      await page.mouse.wheel(0, 120); await d.wait(500);
      await d.caption('보고 싶은 곳에서 <b>누르거나 Ctrl + 2</b> — 그 자리에서 판서합니다', 300);
      await d.tapAt(1150, 560, 600);
      await d.click('.sb-color-dot[data-color="#ef4444"]', 300);
      await d.stroke(shapes.circle(1000, 560, 120, 50));
      await d.wait(500);
      await d.caption('<b>M</b> (Move) — 확대한 화면을 끌어서 옮깁니다', 300);
      await d.press('m', 'M  Move');
      await d.drag(500, 400, 820, 330);
      await d.wait(500);
      await d.drag(800, 360, 420, 420);
      await d.wait(500);
      await d.press('m', 'M  펜으로');
      await d.caption('<b>Ctrl + 1</b>을 다시 누르면 100% 화면으로 돌아옵니다', 300);
      await d.press('Control+1', 'Ctrl + 1  끄기');
      await d.wait(900);
      await d.caption('태블릿에서는 <b>두 손가락</b>을 벌려 확대하면서 이동합니다', 300);
      await d.fingers([[560, 380], [700, 380]], [[420, 300], [840, 300]]);
      await d.caption('<b>세 손가락</b>으로 밀면 확대한 화면을 옮깁니다', 300);
      await d.fingers([[500, 420], [620, 420], [740, 420]], [[760, 300], [880, 300], [1000, 300]]);
      await d.wait(600);
      await d.caption('두 손가락을 오므리면 다시 작아집니다', 300);
      await d.fingers([[400, 360], [880, 360]], [[600, 360], [680, 360]]);
      await d.wait(1200);
    }
  },
  {
    id: '10-pin',
    title: '게시 모드 · 조각 고정',
    summary: '중요한 문제나 공식을 오려 화면 위에 계속 띄워 둡니다.',
    async run(page, d) {
      await d.typeText(150, 210, 'y = a(x − p)² + q');
      await d.stroke(shapes.parabola(320, 520, 150, 220));
      await d.caption('<b>게시</b> 모드 (Ctrl + `) — 띄워 둘 영역을 끌어서 고릅니다', 800);
      await d.click('[data-sb-mode="pin"]', 600);
      await d.drag(130, 180, 520, 540);
      await d.wait(900);
      await d.caption('고른 부분이 <b>조각</b>이 되어 화면 위에 떠 있습니다', 1400);
      await d.click('[data-sb-mode="draw"]', 500);
      await d.caption('<b>쪽을 넘겨도</b> 조각은 그대로 남아 있어 비교하며 설명할 수 있습니다', 800);
      await d.click('#sbAddPageBtn', 1000);
      await d.press('3', '3  파랑');
      await d.stroke(shapes.parabola(880, 560, 220, 300));
      await d.caption('조각은 끌어서 옮기고, 오른쪽 위 <b>×</b>로 닫습니다', 600);
      await d.drag(300, 300, 520, 330);
      await d.wait(800);
      await d.click('.sb-pin-close', 1600);
    }
  },
  {
    id: '11-save',
    title: '노트 저장 · 학급별 복사 · 순서 · PNG · PDF',
    summary: '판서를 학급(폴더)별 노트함에 저장하고, 다른 학급으로 복사하거나 순서를 바꾸고, 그림이나 PDF로 내보냅니다.',
    async run(page, d) {
      await d.stroke(shapes.parabola(400, 470, 160, 200));
      await d.press('3', '3  파랑');
      await d.stroke(shapes.circle(820, 360, 90));
      await d.caption('<b>노트함</b>에서 제목과 폴더(학급)를 정해 판서를 저장합니다', 800);
      await d.click('#sbOpenVaultBtn', 1000);
      for (const title of ['이차함수 1차시', '이차함수 2차시']) {
        await d.click('#sbVaultNoteTitleInput', 200);
        await page.fill('#sbVaultNoteTitleInput', '');
        await page.keyboard.type(title, { delay: 70 });
        await d.wait(300);
        await d.click('#sbVaultSaveNewBtn', 1300);
      }
      await d.caption('<b>◀ 앞으로 · 뒤로 ▶</b>로 노트 순서를 바꿉니다', 600);
      await d.click('[data-move-note][data-dir="1"] >> nth=0', 1200);
      await d.caption('<b>복사</b> — 같은 수업을 다른 학급 폴더로 복사합니다', 600);
      await d.click('[data-copy-note] >> nth=0', 900);
      await page.selectOption('#sbVaultCopyFolder', { index: 1 });
      await d.caption('<b>자료만 복사</b>를 고르면 판서는 비우고 PDF·PPTX·배경만 가져갑니다', 1800);
      await d.click('#sbVaultCopyOkBtn', 1800);
      await d.caption('저장한 노트는 <b>이 기기</b>에 남아, 다음 수업에 다시 불러올 수 있습니다', 1600);
      await page.keyboard.press('Escape');
      await d.wait(400);
      await d.caption('<b>PNG·PDF</b> — 지금 화면은 그림으로, 모든 쪽은 PDF 한 파일로 저장합니다', 800);
      await d.click('#sbSaveMenuBtn', 2400);
      await page.keyboard.press('Escape');
      await d.wait(600);
    }
  },
  {
    id: '12-edit',
    title: '편집 모드 · 손글씨를 글자로 (OCR)',
    summary: '쓴 부분을 골라 옮기거나 지우고, 손글씨를 글자로 바꿉니다.',
    async run(page, d) {
      await d.captionAt('top');
      await d.stroke([[330, 290], [330, 430]]);
      await d.stroke([[390, 310], [450, 285], [480, 330], [390, 430], [490, 430]]);
      await d.stroke(shapes.circle(620, 360, 70));
      await d.caption('<b>편집</b> 모드 (Ctrl+3) — 바꿀 판서를 끌어서 고릅니다', 800);
      await d.click('[data-sb-mode="edit"]', 600);
      await d.drag(270, 270, 700, 450);
      await d.wait(700);
      await d.caption('고른 판서로 <b>글씨로 변환(OCR)</b> · <b>게시 조각 고정</b> · <b>삭제</b>를 할 수 있습니다', 0);
      await d.highlight('#sbEditOcrBtn', 1600).catch(() => d.wait(1600));
      await d.highlight('#sbEditPinBtn', 1300).catch(() => d.wait(1300));
      await d.caption('손글씨 변환은 <b>설정 → AI 손글씨 OCR 설정</b>에서 무료 엔진이나 Gemini 키를 고릅니다', 2600);
      await d.caption('이번에는 고른 판서를 <b>삭제</b>해 봅니다', 600);
      await d.click('#sbEditDelBtn', 1600);
      await d.click('[data-sb-mode="draw"]', 900);
    }
  },
];
