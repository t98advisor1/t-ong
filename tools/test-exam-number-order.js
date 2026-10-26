const { chromium } = require('C:/Users/place/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');

const APP_URL = process.env.TONG_TEST_URL || 'http://127.0.0.1:4321/';
const CHROME = 'C:/Program Files/Google/Chrome/Application/chrome.exe';
const VERSION = require('../package.json').version;
const fs = require('fs');
const path = require('path');
const out = path.resolve(__dirname, '../tmp/seat-panel-v69');

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

async function runAndSkip(page, expectedHeadline) {
  await page.locator('.sp-actions [data-proxy="shuffleBtn"]').click();
  await page.locator('#lotterySplash.show').waitFor({ state: 'visible' });
  await page.locator('#splashSkipBtn').click();
  await page.locator('#lotteryPanel.show').waitFor({ state: 'visible' });
  assert((await page.locator('#lotteryHeadline').innerText()).includes(expectedHeadline), '배치 방향 안내 오류');
  await page.locator('#skipLotteryBtn').click();
  await page.locator('#lotteryPanel').waitFor({ state: 'hidden' });
}

async function seatMap(page) {
  return page.locator('#seatingArea .seat-card').evaluateAll(cards => {
    const map = {};
    cards.forEach(card => {
      const id = card.querySelector('.student-id');
      map[card.dataset.seatId] = id ? id.textContent.trim() : '';
    });
    return map;
  });
}

(async () => {
  const browser = await chromium.launch({ headless: true, executablePath: CHROME });
  fs.mkdirSync(out, { recursive: true });
  const context = await browser.newContext({ viewport: { width: 1024, height: 768 }, serviceWorkers: 'block' });
  const browserErrors = [];

  await context.addInitScript(() => {
    localStorage.clear();
    localStorage.setItem('seatGuide_v2_seen', '1');
    localStorage.setItem('seatSound_v1', '0');
    localStorage.setItem('seatArrangement_v3', JSON.stringify({
      settings: {},
      currentClassIndex: 0,
      allClassData: {
        0: {
          className: '학번순 테스트반',
          students: [
            { id: '10', name: '가학생', gender: '남' },
            { id: '2', name: '나학생', gender: '여' },
            { id: '1', name: '다학생', gender: '남' },
            { id: '6', name: '라학생', gender: '여' },
            { id: '3', name: '마학생', gender: '남' },
            { id: '4', name: '바학생', gender: '여' },
            { id: '7', name: '사학생', gender: '남' },
            { id: '5', name: '아학생', gender: '여' }
          ],
          currentSeats: {},
          lockedSeats: {},
          config: {
            totalDesks: 8,
            columns: 4,
            seatType: 'pair',
            layoutIndex: 0,
            rows: 2,
            sameGenderSeating: true,
            numberOrderMode: 'column',
            numberOrderSeating: true
          },
          relations: { close: { '1': ['10'] }, far: {} },
          studentMemos: {}, specialNotes: [], savedRecords: [],
          praise: {}, attendance: {}, attendReasons: {}, attendMemos: {}
        }
      }
    }));
  });

  await context.route('**/api/usage**', route => route.fulfill({
    status: 200,
    contentType: 'application/json',
    body: JSON.stringify({ todayVisitors: 1, totalVisits: 1 })
  }));

  const page = await context.newPage();
  page.on('console', message => {
    if (message.type() === 'error' && !message.text().includes('Failed to load resource: net::ERR_NETWORK_ACCESS_DENIED')) {
      browserErrors.push(message.text());
    }
  });
  page.on('pageerror', error => browserErrors.push(error.message));
  await page.goto(APP_URL, { waitUntil: 'domcontentloaded' });

  assert(
    (await page.locator('[data-help="version"] small').innerText()) === VERSION,
    '버전 표시가 package.json과 일치하지 않습니다'
  );

  await page.locator('.mode-tab[data-mode="seat"]').click();
  await page.locator('#seatPanel .sp-head').click();
  const orderOptions = page.locator('#numberOrderOption');
  await orderOptions.waitFor({ state: 'visible' });
  assert(await page.locator('input[name="numberOrderMode"][value="column"]').isChecked(), '세로순 설정이 복원되지 않았습니다');
  assert(await page.locator('#sameGenderSeating').isDisabled(), '학번순 모드에서 성별 우선 옵션이 비활성화되지 않았습니다');
  assert((await page.locator('.sp-actions [data-proxy="shuffleBtn"]').innerText()).includes('세로순'), '실행 버튼에 세로순이 표시되지 않습니다');
  assert((await page.locator('#spBody').innerText()).includes('학번 세로순'), '좌석 설정 요약에 세로순이 없습니다');

  await runAndSkip(page, '세로순');
  const vertical = await seatMap(page);
  assert(JSON.stringify(vertical) === JSON.stringify({
    '0-0': '1번', '0-1': '3번', '0-2': '5번', '0-3': '7번',
    '1-0': '2번', '1-1': '4번', '1-2': '6번', '1-3': '10번'
  }), '세로순 좌석 결과 오류: ' + JSON.stringify(vertical));

  await page.locator('input[name="numberOrderMode"][value="row"]').check();
  assert((await page.locator('.sp-actions [data-proxy="shuffleBtn"]').innerText()).includes('가로순'), '실행 버튼에 가로순이 표시되지 않습니다');
  await runAndSkip(page, '가로순');
  const horizontal = await seatMap(page);
  assert(JSON.stringify(horizontal) === JSON.stringify({
    '0-0': '1번', '0-1': '2번', '0-2': '3번', '0-3': '4번',
    '1-0': '5번', '1-1': '6번', '1-2': '7번', '1-3': '10번'
  }), '가로순 좌석 결과 오류: ' + JSON.stringify(horizontal));

  await page.locator('input[name="numberOrderMode"][value="random"]').check();
  assert(!(await page.locator('#sameGenderSeating').isDisabled()), '무작위 모드에서 같은 성별 옵션이 다시 활성화되지 않았습니다');
  assert((await page.locator('.sp-actions [data-proxy="shuffleBtn"]').innerText()).includes('추첨 시작'), '무작위 모드 실행 버튼 문구 오류');
  await page.locator('input[name="numberOrderMode"][value="row"]').check();

  const savedConfig = await page.evaluate(() => JSON.parse(localStorage.getItem('seatArrangement_v3')).allClassData['0'].config);
  assert(savedConfig.numberOrderMode === 'row', '선택한 배치 방향이 학급 설정에 저장되지 않았습니다');

  await page.screenshot({ path: path.join(out, 'number-order-tablet.png'), fullPage: true });
  await page.setViewportSize({ width: 390, height: 844 });
  assert(await orderOptions.isVisible(), '모바일에서 배치 순서 옵션이 보이지 않습니다');
  const pageWidth = await page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, client: document.documentElement.clientWidth }));
  assert(pageWidth.scroll <= pageWidth.client + 1, '모바일 화면에 가로 넘침이 있습니다: ' + JSON.stringify(pageWidth));
  await page.screenshot({ path: path.join(out, 'number-order-mobile.png'), fullPage: true });

  assert(browserErrors.length === 0, '브라우저 오류: ' + browserErrors.join(' | '));
  console.log(JSON.stringify({
    ok: true,
    version: VERSION,
    appUrl: APP_URL,
    vertical,
    horizontal,
    verified: [
      '짝꿍 자리 형태에서도 배치 순서 옵션 표시',
      '학번 세로순 앞→뒤 배치',
      '학번 가로순 왼쪽→오른쪽 배치',
      '숫자 기준 1·2·3…10번 정렬',
      '관계·성별 설정보다 학번순 우선',
      '무작위 전환 시 기존 성별 옵션 복원',
      '배치 방향 저장·복원',
      '모바일 가로 넘침 없음'
    ]
  }, null, 2));
  await browser.close();
})().catch(error => {
  console.error(error.stack || error);
  process.exitCode = 1;
});
