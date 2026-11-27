/* 매뉴얼 영상 장(chapter) 목록 — 스마트 판서 v2.8
 * d.say('자막 HTML', '읽어 줄 문장') : 자막과 음성을 함께 냅니다. 음성 문장은 작은따옴표 문자열로 적어야 record.js가 미리 음성을 만듭니다.
 * d.titleCard('번호', '제목', '부제', '읽어 줄 문장') : 장 시작 화면.
 * 단축키는 음성에서 '컨트롤 일', '제트'처럼 소리 나는 대로 적습니다.
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
const wob = (pts, a) => pts.map(([x, y], i) => [x + Math.sin(i * 1.7) * a, y + Math.cos(i * 1.3) * a]);
const edge = (a, b, n = 14) => Array.from({ length: n }, (_, i) => [a[0] + (b[0] - a[0]) * i / n, a[1] + (b[1] - a[1]) * i / n]);
const poly = v => { let r = []; for (let i = 1; i < v.length; i++) r = r.concat(edge(v[i - 1], v[i])); r.push(v[v.length - 1]); return r; };
async function holdStroke(page, d, pts) {
  await d.moveTo(pts[0][0], pts[0][1], 10); await page.mouse.down();
  for (const q of pts) { await page.mouse.move(q[0], q[1]); await page.waitForTimeout(12); }
  await page.waitForTimeout(950); await page.mouse.up(); await d.wait(500);
}

module.exports = [
  {
    id: 'c01-start',
    title: '처음 시작 · 화면 구성',
    async run(page, d) {
      await d.titleCard('1장', '처음 시작 · 화면 구성', '열자마자 바로 쓰는 칠판', '1장. 처음 시작하기. 화면 구성을 알아봅니다.');
      await d.say('열자마자 <b>펜</b>이 켜져 있어 바로 쓸 수 있습니다', '스마트 판서를 열면 펜이 켜져 있어서 바로 쓸 수 있습니다.');
      await d.stroke(shapes.parabola(640, 520, 220, 260));
      await d.say('첫 줄: <b>홈 · 사용설명서 · 노트함 · 백업 · 열기 · 모드 · 저장 · 공유</b>', '위쪽 첫 줄에는 노트함, 열기, 모드, 저장과 공유 버튼이 있습니다.');
      await d.highlightBox(4, 2, 1272, 40, 2600);
      await d.say('둘째 줄: <b>선택 · 펜 · 도형 · 레이저 · 지우개 · 색 · 굵기 · 되돌리기 · 쪽</b>', '둘째 줄에는 펜, 도형, 레이저, 지우개, 색과 굵기가 있습니다.');
      await d.highlightBox(4, 44, 1272, 42, 2600);
      await d.say('버튼 옆 작은 글자는 <b>단축키</b> — 키보드로 더 빠르게', '버튼 옆의 작은 글자는 단축키입니다. 키보드로 더 빠르게 바꿀 수 있습니다.');
      await d.highlight('#sbPenMenuBtn', 1400);
      await d.say('<b>Tab</b> 키나 <b>▲ 숨김</b>으로 팔레트를 숨기면 화면 전체가 칠판', '탭 키를 누르면 도구 줄이 숨겨져 화면 전체를 칠판으로 씁니다.');
      await d.press('Tab', 'Tab  숨기기');
      await d.wait(1200);
      await d.press('Tab', 'Tab  펼치기');
      await d.say('<b>자동숨김</b>을 켜면 쓰기 시작할 때 저절로 숨고, 화면 맨 위로 가면 다시 나타납니다', '자동 숨김을 켜면 쓰기 시작할 때 저절로 숨고, 화면 맨 위로 가면 다시 나타납니다.');
      await d.click('#sbAutoHideBtn', 700);
      await d.stroke(shapes.wave(420, 300, 440, 50));
      await d.wait(700);
      await d.moveTo(640, 2, 30);
      await d.wait(1400);
      await d.click('#sbAutoHideBtn', 500);
      await d.say('수업 화면을 크게 보여 줄 때는 <b>전체화면</b> (F)', '수업 화면을 크게 보여 줄 때는 에프 키로 전체 화면을 켭니다.');
      await d.highlight('#sbFullscreenBtn', 2000);
      await d.say('모르는 단축키는 <b>?</b> 키로 언제든 확인', '단축키가 생각나지 않으면 물음표 키를 누르세요.');
      await d.press('?', '?  단축키 창');
      await d.wait(1800);
      await d.press('Escape', 'Esc  닫기');
    }
  },
  {
    id: 'c02-pen',
    title: '펜 · 색 · 굵기 · 되돌리기',
    async run(page, d) {
      await d.titleCard('2장', '펜 · 색 · 굵기 · 되돌리기', 'P 펜 · H 형광펜 · 1~6 색 · Z/X 되돌리기', '2장. 펜과 색, 굵기, 되돌리기입니다.');
      await d.say('<b>P</b> 펜 · 기본 굵기 2pt', '피 키는 펜입니다. 기본 굵기는 2포인트입니다.');
      await d.stroke(shapes.parabola(300, 470, 170, 230));
      await d.say('숫자 <b>1~6</b>으로 색을 바로 바꿉니다', '숫자 일부터 육까지 누르면 색이 바로 바뀝니다.');
      await d.press('3', '3  파랑');
      await d.stroke(shapes.circle(640, 330, 90));
      await d.press('2', '2  초록');
      await d.stroke(shapes.wave(520, 560, 300, 40));
      await d.say('<b>+ / −</b> 버튼으로 굵기를 조절합니다', '더하기와 빼기 버튼으로 굵기를 조절합니다.');
      await d.click('#sbSizeIncBtn', 300); await d.click('#sbSizeIncBtn', 300); await d.click('#sbSizeIncBtn', 400);
      await d.press('1', '1  빨강');
      await d.stroke([[880, 520], [1150, 520]]);
      await d.say('<b>펜 종류</b>: 만년필 · 볼펜 · <b>연필</b> · 붓펜 · 형광펜', '펜 메뉴에서 만년필, 볼펜, 연필, 붓펜을 고를 수 있습니다.');
      await d.click('#sbPenMenuBtn', 1000);
      await d.click('#sbPenMenu [data-pen-style="pencil"]', 500);
      await d.press('5', '5  흰색');
      await d.stroke(shapes.wave(860, 300, 300, 40));
      await d.say('<b>H</b> 형광펜 (기본 6pt) — 글자 위에 칠해 강조', '에이치 키는 형광펜입니다. 중요한 부분에 칠해서 강조합니다.');
      await d.press('h', 'H  형광펜');
      await d.press('4', '4  노랑');
      await d.stroke([[540, 330], [740, 330]], 6);
      await d.press('p', 'P  펜');
      await d.say('잘못 썼다면 <b>Z</b> (또는 Esc) — 한 번씩 되돌립니다', '잘못 썼다면 제트 키나 이에스씨 키로 되돌립니다.');
      await d.press('z', 'Z  되돌리기');
      await d.wait(500);
      await d.press('Escape', 'Esc  되돌리기');
      await d.say('<b>X</b> — 되돌린 것을 다시 살립니다', '엑스 키를 누르면 되돌린 것이 다시 살아납니다.');
      await d.press('x', 'X  다시 실행');
      await d.wait(500);
      await d.press('x', 'X  다시 실행');
      await d.say('Ctrl+Z · Ctrl+Y 도 똑같이 쓸 수 있습니다', '컨트롤 제트와 컨트롤 와이도 똑같이 쓸 수 있습니다.');
    }
  },
  {
    id: 'c03-shapes',
    title: '도형 · 텍스트 · 꾹 눌러 도형',
    async run(page, d) {
      await d.titleCard('3장', '도형 · 텍스트 · 꾹 눌러 도형', 'R 사각형 · C 원 · L 직선 · T 텍스트', '3장. 도형과 텍스트, 그리고 꾹 눌러 도형 만들기입니다.');
      await d.say('<b>도형</b> 메뉴에 직선 · 사각형 · 원 · 채움사각 · 텍스트', '도형 메뉴에는 직선, 사각형, 원, 텍스트가 있습니다.');
      await d.click('#sbMoreToolsBtn', 1100);
      await d.click('#sbMoreToolsMenu [data-sb-tool="rect"]', 400);
      await d.drag(140, 180, 400, 340);
      await d.say('단축키 <b>R</b> 사각형 · <b>C</b> 원 · <b>L</b> 직선', '단축키 알은 사각형, 씨는 원, 엘은 직선입니다.');
      await d.press('c', 'C  원');
      await d.drag(500, 170, 720, 360);
      await d.press('l', 'L  직선');
      await d.drag(830, 270, 1150, 270);
      await d.say('<b>T</b> 텍스트 — 누른 자리에 글자를 쓰고 Enter', '티 키를 누르고 화면을 누르면 글자를 입력합니다.');
      await d.setSize(8);
      await d.press('t', 'T  텍스트');
      await d.tapAt(150, 450, 400);
      await page.keyboard.type('정답 : x = 3', { delay: 90 });
      await page.keyboard.press('Enter');
      await d.wait(400);
      await d.press('p', 'P  펜');
      await d.setSize(5);
      await d.say('펜으로 그리다가 끝에서 <b>꾹 멈추면</b> 반듯한 도형으로 바뀝니다', '펜으로 그리다가 끝에서 꾹 멈추면 반듯한 도형으로 바뀝니다.');
      await holdStroke(page, d, wob(shapes.circle(560, 560, 70, 40), 5));
      await holdStroke(page, d, wob(poly([[700, 490], [860, 495], [855, 630], [698, 625], [702, 492]]), 3));
      await d.say('세모 · 별 · 직선도 꾹 — 형광펜에서도 됩니다', '세모와 별, 직선도 똑같이 됩니다. 형광펜에서도 쓸 수 있습니다.');
      await holdStroke(page, d, wob(poly([[950, 630], [1030, 490], [1110, 632], [952, 628]]), 3));
      const sv = []; for (let i = 0; i <= 5; i++) { const a = -Math.PI / 2 + i * 4 * Math.PI / 5; sv.push([1170 + 70 * Math.cos(a), 470 + 70 * Math.sin(a)]); }
      await holdStroke(page, d, wob(poly(sv), 3));
      await d.say('멈추지 않고 떼면 손글씨 그대로 남습니다', '멈추지 않고 바로 떼면 손글씨 그대로 남습니다.');
      await d.stroke(shapes.wave(420, 660, 240, 20));
    }
  },
  {
    id: 'c04-stamps',
    title: '번호 스탬프 · 자동 번호 · 가림막',
    async run(page, d) {
      await d.titleCard('4장', '번호 스탬프 · 자동 번호 · 가림막', 'S 원형 · Q 사각 · SS/0 처음부터 · A 자동 번호', '4장. 번호 스탬프와 자동 번호, 정답 가림막입니다.');
      await d.say('<b>S</b> 원형 번호 스탬프 — 누를 때마다 다음 번호', '에스 키는 원형 번호 스탬프입니다. 누를 때마다 다음 번호가 찍힙니다.');
      await d.press('s', 'S  원형 스탬프');
      for (const y of [200, 290, 380]) await d.tapAt(160, y, 450);
      await d.say('<b>Q</b> 사각 번호 스탬프', '큐 키는 사각 번호 스탬프입니다.');
      await d.press('q', 'Q  사각 스탬프');
      await d.tapAt(300, 200, 450); await d.tapAt(300, 290, 600);
      await d.say('<b>S 두 번(SS)</b> 또는 <b>0</b> — 번호를 1번부터 다시', '에스를 두 번 빠르게 누르거나 영을 누르면 번호가 1번부터 다시 시작합니다.');
      await d.wait(500);
      await page.keyboard.press('s'); await d.wait(120); await d.press('s', 'S S  처음부터');
      await d.tapAt(440, 200, 500); await d.tapAt(440, 290, 700);
      await d.say('<b>A</b> 자동 번호 — 밑줄이나 원을 그리면 왼쪽 위에 번호가 붙습니다', '에이 키를 켜면 밑줄이나 원을 그릴 때 왼쪽 위에 번호가 저절로 붙습니다.');
      await d.press('a', 'A  자동 번호 켜기');
      await d.press('l', 'L  직선');
      await d.drag(620, 260, 900, 260);
      await d.press('c', 'C  원');
      await d.drag(960, 180, 1160, 340);
      await d.press('a', 'A  자동 번호 끄기');
      await d.press('p', 'P  펜');
      await d.say('스탬프만 한꺼번에 지우기: <b>Ctrl+S</b> (Ctrl+Z로 되살림)', '스탬프만 한꺼번에 지울 때는 컨트롤 에스를 누릅니다.');
      await d.press('Control+s', 'Ctrl + S  스탬프 지우기');
      await d.wait(900);
      await d.press('Control+z', 'Ctrl + Z  되살리기');
      await d.say('정답을 가릴 때는 도형 메뉴의 <b>정답 가림막</b>', '정답을 가릴 때는 도형 메뉴의 정답 가림막을 씁니다.');
      await d.typeText(640, 520, '정답 : 2, 4, 6');
      await d.click('#sbMoreToolsBtn', 900);
      await d.click('#sbMoreToolsMenu [data-sb-tool="cover"]', 400);
      await d.drag(600, 490, 960, 570);
      await page.keyboard.press('Escape');
      await d.say('수업 중에 가림막을 <b>누르면 열리고</b>, 다시 누르면 닫힙니다', '수업 중에 가림막을 누르면 정답이 열리고, 다시 누르면 닫힙니다.');
      await d.tapAt(780, 530, 1500);
      await d.tapAt(780, 530, 1200);
    }
  },
  {
    id: 'c05-select',
    title: '선택 · 레이저 · 지우개',
    async run(page, d) {
      await d.titleCard('5장', '선택 · 레이저 · 지우개', 'V 선택 · D 레이저 · E 지우개', '5장. 선택해서 옮기기, 레이저, 지우개입니다.');
      await d.stroke(shapes.circle(300, 360, 80));
      await d.stroke(shapes.wave(520, 360, 320, 60));
      await d.say('<b>V</b> 선택 — 쓴 것을 잡아 원하는 곳으로 옮깁니다', '브이 키로 선택한 뒤 쓴 것을 끌어서 옮깁니다.');
      await d.press('v', 'V  선택');
      await d.drag(300, 280, 470, 520);
      await d.wait(600);
      await d.say('<b>D</b> 레이저 — 여러 번 그어도 3초 뒤 사라집니다', '디 키는 레이저입니다. 여러 번 그어도 3초 뒤에 사라집니다.');
      await d.press('d', 'D  레이저');
      await d.stroke(shapes.circle(700, 360, 110, 50), 5);
      await d.stroke([[600, 520], [900, 520]], 5);
      await d.wait(3200);
      await d.say('<b>E</b> 지우개 — 문지른 선이 통째로 지워집니다', '이 키는 지우개입니다. 문지른 선이 통째로 지워집니다.');
      await d.press('e', 'E  지우개');
      await d.stroke([[690, 250], [690, 460]], 6);
      await d.wait(600);
      await d.say('잘못 지웠다면 <b>Z</b> 또는 Esc — 지운 것도 되살아납니다', '잘못 지웠다면 제트 키로 되살립니다.');
      await d.press('z', 'Z  되돌리기');
      await d.wait(600);
      await d.say('갤럭시 <b>S펜 옆 버튼</b>을 누르면 지우개와 펜이 바로 바뀝니다', '갤럭시 에스펜 옆 버튼을 누르면 지우개와 펜이 바로 바뀝니다.');
      await d.press('p', 'P  펜');
      await d.press('s', 'S  원형 스탬프');
      for (const x of [900, 980, 1060]) await d.tapAt(x, 520, 300);
      await d.press('p', 'P  펜');
      await d.say('한꺼번에 지울 때는 <b>지우개 ▾</b> 메뉴', '한꺼번에 지울 때는 지우개 옆 메뉴를 엽니다.');
      await d.click('#sbEraserToolBtn .sb-caret', 1300);
      await d.say('펜 글씨 <b>Shift+E</b> · 스탬프 <b>Ctrl+S</b> · 전체 필기 <b>Ctrl+E</b>', '펜 글씨만, 스탬프만, 또는 전체 필기를 한 번에 지울 수 있습니다.');
      await d.click('#sbEraseAllPenBtn', 1500);
      await d.say('한꺼번에 지운 것도 <b>Ctrl+Z</b> 한 번에 되살아납니다', '한꺼번에 지운 것도 컨트롤 제트 한 번이면 되살아납니다.');
      await d.press('Control+z', 'Ctrl + Z');
    }
  },
  {
    id: 'c06-pages',
    title: '쪽 넘기기 · 배경',
    async run(page, d) {
      await d.titleCard('6장', '쪽 넘기기 · 배경', '+쪽 · ← → 넘기기 · K 칠판 · W 백지', '6장. 쪽을 추가하고 넘기기, 배경 바꾸기입니다.');
      await d.stroke(shapes.parabola(640, 480, 200, 230));
      await d.say('<b>+쪽</b>을 누르면 새 칠판 쪽이 생깁니다', '더하기 쪽 버튼을 누르면 새 쪽이 생깁니다.');
      await d.click('#sbAddPageBtn', 900);
      await d.stroke(shapes.circle(640, 380, 100));
      await d.say('<b>← →</b> 방향키나 ◀ ▶ 버튼으로 쪽을 넘깁니다', '왼쪽 오른쪽 방향키로 쪽을 넘깁니다.');
      await d.press('ArrowLeft', '←  이전 쪽');
      await d.wait(900);
      await d.press('ArrowRight', '→  다음 쪽');
      await d.say('<b>설정 ▾</b>에서 칠판 · 백지 · <b>모눈</b> 배경', '설정 메뉴에서 칠판, 백지, 모눈 배경을 고릅니다.');
      await d.click('#sbBoardMoreBtn', 1000);
      await d.click('[data-sb-bg="grid"]', 900);
      await page.keyboard.press('Escape');
      await d.press('3', '3  파랑');
      await d.stroke([[340, 380], [940, 380]]);
      await d.stroke([[640, 160], [640, 600]]);
      await d.say('단축키 <b>W</b> 백지 · <b>K</b> 칠판', '더블유는 백지, 케이는 칠판입니다.');
      await d.press('w', 'W  백지');
      await d.wait(1000);
      await d.press('k', 'K  칠판');
    }
  },
  {
    id: 'c07-pdf',
    title: 'PDF 열어서 판서',
    async run(page, d) {
      await d.titleCard('7장', 'PDF 열어서 판서', '교과서 · 활동지 위에 바로', '7장. 피디에프를 열어서 판서하기입니다.');
      await d.say('<b>열기</b> → <b>PDF 열기</b>에서 파일을 고릅니다 (끌어다 놓아도 됩니다)', '열기 버튼에서 피디에프 열기를 누르고 파일을 고릅니다.');
      await openFromHub(page, d, 'sbPdfInput', 'worksheet.pdf');
      await d.wait(2400);
      await d.say('PDF 위에 바로 판서합니다', '피디에프 위에 바로 판서할 수 있습니다.');
      await d.stroke([[355, 168], [475, 168]]);
      await d.stroke(shapes.circle(652, 143, 42, 40));
      await d.press('3', '3  파랑');
      await d.typeText(300, 232, '꼭짓점 (2, −1)');
      await d.say('<b>→</b> 다음 쪽 — 쪽마다 판서가 따로 남습니다', '오른쪽 방향키로 다음 쪽으로 가면 쪽마다 판서가 따로 남습니다.');
      await d.press('ArrowRight', '→  다음 쪽');
      await d.typeText(300, 232, 'a가 커지면 폭이 좁아진다');
      await d.press('ArrowLeft', '←  이전 쪽');
      await d.wait(1000);
    }
  },
  {
    id: 'c08-pptx',
    title: 'PPTX 발표 자료 · 애니메이션',
    async run(page, d) {
      await d.titleCard('8장', 'PPTX 발표 자료 · 애니메이션', 'Space · → 다음 장면', '8장. 피피티 발표 자료와 클릭 애니메이션입니다.');
      await d.say('<b>열기 → PPTX 열기</b>로 발표 자료를 엽니다', '열기에서 피피티 열기를 눌러 발표 자료를 엽니다.');
      await openFromHub(page, d, 'sbPptxInput', 'lesson.pptx');
      await d.wait(2400);
      await d.say('<b>Space</b>나 <b>→</b> — 클릭 애니메이션이 순서대로 나타납니다', '스페이스나 오른쪽 방향키를 누르면 클릭 애니메이션이 순서대로 나타납니다.');
      await d.press('Space', 'Space');
      await d.wait(900);
      await d.press('Space', 'Space');
      await d.wait(700);
      await d.press('ArrowRight', '→');
      await d.wait(700);
      await d.press('ArrowRight', '→');
      await d.say('<b>←</b> 이전 장면 · 슬라이드 위에도 바로 판서', '왼쪽 방향키는 이전 장면입니다. 슬라이드 위에도 바로 판서합니다.');
      await d.stroke(shapes.circle(1010, 380, 150, 50));
      await d.say('<b>전체</b>를 누르면 애니메이션 없이 모두 보여 줍니다', '전체 버튼을 누르면 애니메이션을 건너뛰고 모두 보여 줍니다.');
      await d.click('#sbPptxNextActBtn', 700);
      await d.click('#sbPptxShowAllBtn', 1600);
    }
  },
  {
    id: 'c09-web',
    title: '유튜브 · 웹 위에 판서',
    async routes(ctx) {
      const html = fs.readFileSync(path.join(SAMPLES, 'web-demo.html'));
      await ctx.route('https://example.com/**', r => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: html }));
    },
    async run(page, d) {
      await d.titleCard('9장', '유튜브 · 웹 위에 판서', '주소를 넣고 링크 열기', '9장. 유튜브와 웹 사이트 위에 판서하기입니다.');
      await d.say('<b>열기</b> 창 아래에 유튜브 · 웹 주소를 넣고 <b>Enter</b>', '열기 창 아래에 유튜브나 웹 주소를 넣고 엔터를 누릅니다.');
      await d.click('#sbOpenFileHubBtn', 900);
      await d.click('#sbFileHubUrlInput', 300);
      await page.keyboard.type('https://example.com/earth', { delay: 40 });
      await page.keyboard.press('Enter');
      await d.wait(2200);
      await d.say('웹 화면 위에 그대로 <b>펜 필기</b>', '웹 화면 위에 그대로 필기할 수 있습니다.');
      await d.stroke(shapes.circle(410, 512, 90, 50));
      await d.stroke([[620, 470], [760, 470]]);
      await d.say('<b>① 원형스탬프</b>로 짚어 가며 설명', '원형 스탬프로 짚어 가며 설명합니다.');
      await d.click('#sbWebCircleStampBtn', 500);
      await d.tapAt(560, 470, 500); await d.tapAt(860, 470, 500);
      await d.say('영상 재생 · 스크롤은 <b>웹 조작</b>을 켜고 합니다', '영상을 재생하거나 스크롤할 때는 웹 조작을 켭니다.');
      await d.click('#sbWebInteractToggleBtn', 1400);
      await d.say('막힌 사이트는 <b>새 탭</b>이나 <b>스냅샷 고정</b>으로', '띄워지지 않는 사이트는 새 탭이나 스냅샷 고정을 쓰세요.');
      await d.highlight('#sbWebSnapBtn', 1800).catch(() => d.wait(1800));
      await d.click('#sbWebCloseBtn', 900);
    }
  },
  {
    id: 'c10-zoom',
    title: '확대 · 손가락으로 이동',
    contextOptions: { hasTouch: true },
    async run(page, d) {
      await d.titleCard('10장', '확대 · 손가락으로 이동', 'Ctrl+1 확대 · 한 번 더 누르면 원래대로', '10장. 화면 확대와 손가락으로 이동하기입니다.');
      await d.typeText(110, 200, 'y = x² − 4x + 3');
      await d.stroke(shapes.parabola(330, 560, 170, 260));
      await d.press('h', 'H  형광펜');
      await d.press('4', '4  노랑');
      await d.say('형광펜을 쓰다가 <b>Ctrl+1</b> — 커서 쪽을 크게 보여 줍니다', '형광펜을 쓰다가 컨트롤 일을 누르면 커서 쪽을 크게 보여 줍니다.');
      await d.moveTo(640, 360, 8);
      await d.press('Control+1', 'Ctrl + 1  확대');
      await d.moveTo(160, 230, 40);
      await d.wait(600);
      await d.say('마우스 <b>휠</b>로 배율 · 누르면 그 자리에서 판서', '마우스 휠로 배율을 바꾸고, 누르면 그 자리에서 판서합니다.');
      await page.mouse.wheel(0, -120); await d.wait(500);
      await d.tapAt(170, 220, 500);
      await d.stroke([[120, 200], [400, 200]], 6);
      await d.say('<b>M</b> — 확대한 화면을 끌어서 옮깁니다', '엠 키를 누르면 확대한 화면을 끌어서 옮깁니다.');
      await d.press('m', 'M  Move');
      await d.drag(500, 400, 760, 330);
      await d.wait(400);
      await d.say('<b>Ctrl+1</b> 한 번 더 — 100% 화면, 쓰던 <b>형광펜 그대로</b>', '컨트롤 일을 한 번 더 누르면 원래 화면으로, 쓰던 형광펜 그대로 돌아옵니다.');
      await d.press('Control+1', 'Ctrl + 1  원래대로');
      await d.wait(600);
      await d.stroke([[600, 600], [900, 600]], 6);
      await d.say('태블릿: <b>두 손가락</b>으로 확대 · 이동', '태블릿에서는 두 손가락을 벌려 확대하면서 이동합니다.');
      await d.fingers([[560, 380], [700, 380]], [[420, 300], [840, 300]]);
      await d.say('<b>세 손가락</b>으로 밀면 화면 이동', '세 손가락으로 밀면 화면이 옮겨집니다.');
      await d.fingers([[500, 420], [620, 420], [740, 420]], [[740, 320], [860, 320], [980, 320]]);
      await d.say('두 손가락을 오므리면 다시 작아집니다', '두 손가락을 오므리면 다시 작아집니다.');
      await d.fingers([[400, 360], [880, 360]], [[600, 360], [680, 360]]);
    }
  },
  {
    id: 'c11-pin',
    title: '게시 모드 · 조각 고정',
    async run(page, d) {
      await d.titleCard('11장', '게시 모드 · 조각 고정', 'Ctrl+` 게시', '11장. 게시 모드로 화면 조각 고정하기입니다.');
      await d.typeText(150, 210, 'y = a(x − p)² + q');
      await d.stroke(shapes.parabola(320, 520, 150, 220));
      await d.say('<b>Ctrl+`</b> 게시 모드 — 띄워 둘 부분을 끌어서 고릅니다', '컨트롤과 숫자 일 왼쪽 키를 누르면 게시 모드입니다. 띄워 둘 부분을 끌어서 고릅니다.');
      await d.press('Control+`', 'Ctrl + `  게시');
      await d.drag(130, 180, 520, 540);
      await d.wait(800);
      await d.say('고른 부분이 <b>조각</b>이 되어 화면 위에 떠 있습니다', '고른 부분이 조각이 되어 화면 위에 떠 있습니다.');
      await d.press('Control+2', 'Ctrl + 2  판서');
      await d.say('<b>쪽을 넘겨도</b> 조각은 남아 비교하며 설명할 수 있습니다', '쪽을 넘겨도 조각은 남아 있어 비교하며 설명할 수 있습니다.');
      await d.click('#sbAddPageBtn', 900);
      await d.press('3', '3  파랑');
      await d.stroke(shapes.parabola(880, 560, 220, 300));
      await d.say('조각은 끌어서 옮기고 <b>×</b>로 닫습니다', '조각은 끌어서 옮기고 엑스 버튼으로 닫습니다.');
      await d.drag(300, 300, 520, 330);
      await d.wait(700);
      await d.click('.sb-pin-close', 1200);
    }
  },
  {
    id: 'c12-edit',
    title: '편집 모드 · 손글씨를 글자로',
    async run(page, d) {
      await d.titleCard('12장', '편집 모드 · 손글씨를 글자로', 'Ctrl+3 편집 · Delete 삭제', '12장. 편집 모드와 손글씨를 글자로 바꾸기입니다.');
      await d.captionAt('top');
      await d.stroke([[330, 290], [330, 430]]);
      await d.stroke([[390, 310], [450, 285], [480, 330], [390, 430], [490, 430]]);
      await d.stroke(shapes.circle(620, 360, 70));
      await d.say('<b>Ctrl+3</b> 편집 모드 — 바꿀 판서를 끌어서 고릅니다', '컨트롤 삼을 누르면 편집 모드입니다. 바꿀 판서를 끌어서 고릅니다.');
      await d.press('Control+3', 'Ctrl + 3  편집');
      await d.drag(270, 270, 700, 450);
      await d.wait(600);
      await d.say('고른 판서를 <b>글씨로 변환(OCR)</b> · <b>조각 고정</b> · <b>삭제</b>', '고른 판서는 글씨로 바꾸거나, 조각으로 고정하거나, 지울 수 있습니다.');
      await d.highlight('#sbEditOcrBtn', 1500).catch(() => d.wait(1500));
      await d.highlight('#sbEditPinBtn', 1200).catch(() => d.wait(1200));
      await d.say('손글씨 변환은 <b>설정 → AI 손글씨 OCR 설정</b>에서 엔진을 고릅니다', '손글씨 변환은 설정의 오씨알 설정에서 엔진을 고릅니다.');
      await d.wait(1500);
      await d.say('<b>Delete</b> — 고른 판서를 지웁니다', '딜리트 키를 누르면 고른 판서를 지웁니다.');
      await d.press('Delete', 'Delete  삭제');
      await d.press('Control+2', 'Ctrl + 2  판서');
    }
  },
  {
    id: 'c13-notes',
    title: '노트함 · 자동 저장 · 이어 쓰기',
    async run(page, d, ctx) {
      await d.titleCard('13장', '노트함 · 자동 저장 · 이어 쓰기', '1분마다 자동 저장 · 다시 열면 마지막 판서', '13장. 노트함과 자동 저장, 이어 쓰기입니다.');
      await d.stroke(shapes.parabola(400, 470, 160, 200));
      await d.press('3', '3  파랑');
      await d.stroke(shapes.circle(820, 360, 90));
      await d.say('판서는 <b>1분마다 자동 저장</b> — 따로 저장하지 않아도 됩니다', '판서는 1분마다 자동으로 저장됩니다.');
      await d.say('<b>노트함</b>에서 제목과 폴더(학급)를 정해 저장할 수도 있습니다', '노트함에서 제목과 학급 폴더를 정해 저장할 수도 있습니다.');
      await d.click('#sbOpenVaultBtn', 1000);
      for (const title of ['이차함수 1차시', '이차함수 2차시']) {
        await d.click('#sbVaultNoteTitleInput', 200);
        await page.fill('#sbVaultNoteTitleInput', '');
        await page.keyboard.type(title, { delay: 60 });
        await d.click('#sbVaultSaveNewBtn', 1200);
      }
      await d.say('노트는 <b>최근에 연 순서</b>로 정렬됩니다', '노트는 최근에 연 순서로 정렬됩니다.');
      await d.highlight('#sbVaultSortSelect', 1500).catch(() => d.wait(1500));
      await d.say('<b>복사</b> — 같은 수업을 다른 학급 폴더로 (자료만 / 판서까지)', '복사 버튼으로 같은 수업을 다른 학급 폴더로 복사합니다.');
      await d.click('[data-copy-note] >> nth=0', 900);
      await page.selectOption('#sbVaultCopyFolder', { index: 1 });
      await d.wait(1200);
      await d.click('#sbVaultCopyOkBtn', 1400);
      await page.keyboard.press('Escape');
      await d.say('<b>Backspace</b> — 저장하고 교무수첩으로 나갑니다', '백스페이스를 누르면 저장하고 교무수첩으로 나갑니다.');
      await d.keyBadgeAsync('Backspace  나가기');
      await d.wait(900);
      await page.keyboard.press('Backspace');
      await page.waitForLoadState('load').catch(() => {});
      await d.wait(1800);
      await d.say('다음에 스마트 판서를 열면 <b>마지막 판서로 바로</b> 돌아옵니다', '다음에 스마트 판서를 열면 마지막으로 쓰던 판서로 바로 돌아옵니다.');
      await page.goto(page.url().replace(/[^/]*$/, 'smart-board.html'));
      await page.waitForTimeout(1500);
      await d.moveTo(640, 360, 4);
      await d.wait(1500);
    }
  },
  {
    id: 'c14-backup',
    title: '백업 · 복원 (.json)',
    contextOptions: { acceptDownloads: true },
    async run(page, d) {
      await d.titleCard('14장', '백업 · 복원', '다른 기기 · 새 주소로 노트 옮기기', '14장. 노트 백업과 복원입니다.');
      await d.stroke(shapes.parabola(400, 470, 160, 200));
      await d.click('#sbOpenVaultBtn', 700);
      await page.fill('#sbVaultNoteTitleInput', '1학년 3반 확률');
      await d.click('#sbVaultSaveNewBtn', 900);
      await page.keyboard.press('Escape');
      await d.say('노트는 <b>이 기기의 브라우저</b>에 저장됩니다 — 옮길 때는 백업 파일로', '노트는 이 기기의 브라우저에 저장됩니다. 다른 기기로 옮길 때는 백업 파일을 씁니다.');
      await d.click('#sbBackupMenuBtn', 1300);
      await d.say('<b>전체 노트 백업(.json)</b> — 모든 폴더와 노트를 파일 하나로', '전체 노트 백업을 누르면 모든 노트가 파일 하나로 내려받아집니다.');
      const [dl] = await Promise.all([page.waitForEvent('download'), d.click('#sbMenuBackupAllBtn', 300)]);
      const file = path.join(require('os').tmpdir(), 'sb-demo-backup.json');
      await dl.saveAs(file);
      await d.wait(1200);
      await d.say('새 기기에서는 <b>백업 파일 불러오기 / 복원</b>을 누르고 파일을 고릅니다', '새 기기에서는 백업 파일 불러오기를 누르고 그 파일을 고릅니다.');
      await d.click('#sbBackupMenuBtn', 900);
      const [fc] = await Promise.all([page.waitForEvent('filechooser'), d.click('#sbMenuRestoreBtn', 300)]);
      await fc.setFiles(file);
      await d.wait(1500);
      await d.say('폴더와 노트가 그대로 돌아옵니다', '폴더와 노트가 그대로 돌아옵니다.');
      await d.click('#sbOpenVaultBtn', 2600);
      await d.click('#sbVaultModal [data-close="sbVaultModal"], #sbVaultModal .sb-modal-close, #sbVaultCloseBtn', 600).catch(() => page.keyboard.press('Escape'));
    }
  },
  {
    id: 'c15-share',
    title: '저장 · 공유 (PNG · ZIP · PDF)',
    contextOptions: { acceptDownloads: true, permissions: ['clipboard-read', 'clipboard-write'] },
    async run(page, d) {
      await d.titleCard('15장', '저장 · 공유', 'Ctrl+P 그림 저장 · 노션에 붙여넣기 · PDF 공유', '15장. 판서를 그림이나 피디에프로 저장하고 공유하기입니다.');
      await d.stroke(shapes.parabola(400, 470, 160, 200));
      await d.press('3', '3  파랑');
      await d.stroke(shapes.circle(820, 360, 90));
      await d.say('<b>Ctrl+P</b> — 지금 화면을 PNG 그림으로 저장', '컨트롤 피를 누르면 지금 화면이 그림 파일로 저장됩니다.');
      await d.press('Control+p', 'Ctrl + P  PNG 저장');
      await d.say('<b>PNG·PDF ▾</b>: 전체 쪽 PNG 묶음(ZIP) · 합본 PDF', '저장 메뉴에서는 모든 쪽을 그림 묶음이나 피디에프 한 파일로 저장합니다.');
      await d.click('#sbSaveMenuBtn', 2400);
      await page.keyboard.press('Escape');
      await d.say('<b>공유 ▾</b>: 태블릿 · 휴대폰은 공유 창으로 노션 · 카톡 · 드라이브에 바로', '공유 메뉴를 누르면 태블릿과 휴대폰에서는 공유 창이 열려 노션이나 카카오톡으로 바로 보냅니다.');
      await d.click('#sbShareMenuBtn', 2200);
      await d.say('컴퓨터에서는 <b>이미지 복사</b> → 노션에서 <b>Ctrl+V</b>', '컴퓨터에서는 이미지 복사를 누르고 노션에서 컨트롤 브이로 붙여 넣습니다.');
      await d.click('#sbCopyPngBtn', 1600);
      await d.say('여러 쪽은 <b>PDF로 공유</b> — 한 파일로 묶어 보냅니다', '여러 쪽은 피디에프로 공유하면 한 파일로 묶어서 보냅니다.');
      await d.click('#sbShareMenuBtn', 800);
      await d.highlight('#sbSharePdfBtn', 1800).catch(() => d.wait(1800));
      await page.keyboard.press('Escape');
    }
  },
  {
    id: 'c16-settings',
    title: '설정 · S펜 · 버전',
    async run(page, d) {
      await d.titleCard('16장', '설정 · S펜 · 버전', 'S펜 자동 · S펜 버튼 · 수정 이력', '16장. 설정과 에스펜, 버전 확인입니다.');
      await d.say('<b>설정 ▾</b>을 엽니다', '설정 메뉴를 엽니다.');
      await d.click('#sbBoardMoreBtn', 1000);
      await d.say('<b>S펜 자동</b> — S펜이 닿으면 손바닥은 무시하고 손가락은 확대 · 이동', '에스펜 자동을 켜 두면, 에스펜이 닿을 때 손바닥은 무시하고 손가락으로는 확대와 이동만 합니다.');
      await d.highlight('#sbSpenAutoBtn', 1800).catch(() => d.wait(1800));
      await d.say('<b>S펜 버튼</b> — 옆 버튼으로 지우개 ↔ 펜 바꾸기', '에스펜 버튼 설정을 켜면 옆 버튼으로 지우개와 펜을 바꿉니다.');
      await d.highlight('#sbSpenEraserModeBtn', 1800).catch(() => d.wait(1800));
      await d.say('<b>꽉찬화면</b> · <b>AI 손글씨 OCR 설정</b>도 여기에 있습니다', '꽉 찬 화면과 손글씨 변환 설정도 여기에 있습니다.');
      await d.highlight('#sbOcrCfgBtn', 1500).catch(() => d.wait(1500));
      await d.say('맨 아래에서 <b>버전</b>과 <b>수정 이력</b>을 확인합니다', '맨 아래에서 버전과 수정 이력을 확인할 수 있습니다.');
      await d.click('#sbChangelogBtn', 2600);
      await page.keyboard.press('Escape');
      await d.say('궁금할 때는 언제든 <b>사용설명서</b> 버튼', '궁금한 점이 있으면 언제든 사용설명서 버튼을 누르세요.');
      await d.highlight('#sbHelpBtn', 1800);
    }
  },
];
