const { chromium } = require('C:/Users/place/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');

const APP_URL = process.env.TONG_TEST_URL || 'http://127.0.0.1:4321/';
const CHROME = 'C:/Program Files/Google/Chrome/Application/chrome.exe';

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

async function assertGrowth(page, name, icon, label, score) {
  const row = page.locator('#praiseList .praise-row').filter({ hasText: name }).first();
  await row.waitFor({ state: 'visible' });
  assert(await row.locator('.pr-stage-icon').innerText() === icon, name + ' 성장 아이콘 오류');
  const text = await row.locator('.pr-stage-text').innerText();
  assert(text.includes(label) && text.includes(score + '점'), name + ' 성장 단계/점수 오류: ' + text);
}

async function assertSeatGrowth(page, name, icon, score) {
  const card = page.locator('#seatingArea .seat-card').filter({ hasText: name }).first();
  await card.waitFor({ state: 'visible' });
  const badge = card.locator('.praise-badge');
  assert(await badge.count() === 1, name + ' 좌석 성장 배지가 없습니다');
  const text = (await badge.innerText()).replace(/\s/g, '');
  assert(text.includes(icon) && text.includes(String(score)), name + ' 좌석 성장 배지 오류: ' + text);
}

(async () => {
  const browser = await chromium.launch({ headless: true, executablePath: CHROME });
  const context = await browser.newContext({ viewport: { width: 1024, height: 768 } });
  const consoleErrors = [];

  await context.addInitScript(() => {
    localStorage.clear();
    localStorage.setItem('seatGuide_v2_seen', '1');
    localStorage.setItem('seatArrangement_v3', JSON.stringify({
      settings: {},
      currentClassIndex: 0,
      allClassData: {
        0: {
          className: '성장 테스트반',
          students: [
            { id: '1', name: '알학생', gender: '여' },
            { id: '2', name: '부화학생', gender: '남' },
            { id: '3', name: '병아리학생', gender: '여' },
            { id: '4', name: '닭학생', gender: '남' }
          ],
          currentSeats: {
            '0-0': { studentId: '1', name: '알학생' },
            '0-1': { studentId: '2', name: '부화학생' },
            '0-2': { studentId: '3', name: '병아리학생' },
            '0-3': { studentId: '4', name: '닭학생' }
          },
          lockedSeats: {},
          config: { totalDesks: 4, columns: 4, seatType: 'single', layoutIndex: 0, rows: 1 },
          relations: { close: {}, far: {} },
          studentMemos: {}, specialNotes: [], savedRecords: [],
          praise: { '2': 1, '3': 5, '4': 10 },
          attendance: {}, attendReasons: {}, attendMemos: {}
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
      consoleErrors.push(message.text());
    }
  });
  page.on('pageerror', error => consoleErrors.push(error.message));
  await page.goto(APP_URL, { waitUntil: 'domcontentloaded' });

  assert(
    (await page.locator('[data-help="version"] small').innerText()) === '5.17.0-preview.66',
    '버전 표시가 preview.66가 아닙니다'
  );

  await page.locator('#homeBtn').click();
  const backupShortcut = page.locator('#homeMenu [data-home="backup"]');
  await backupShortcut.waitFor({ state: 'visible' });
  assert((await backupShortcut.innerText()).includes('데이터 백업'), '홈 메뉴 백업 문구가 없습니다');
  await backupShortcut.click();
  await page.locator('#backupModal').waitFor({ state: 'visible' });
  await page.locator('#backupModal .close-btn').click();

  await page.locator('.mode-tab[data-mode="seat"]').click();
  await assertSeatGrowth(page, '알학생', '🥚', 0);
  await assertSeatGrowth(page, '부화학생', '🐣', 1);
  await assertSeatGrowth(page, '병아리학생', '🐥', 5);
  await assertSeatGrowth(page, '닭학생', '🐔', 10);

  await page.locator('#extraBtn').click();
  await page.locator('#extraMenu [data-open="praise"]').click();
  await page.locator('#praiseModal').waitFor({ state: 'visible' });
  await assertGrowth(page, '알학생', '🥚', '알', 0);
  await assertGrowth(page, '부화학생', '🐣', '부화 중', 1);
  await assertGrowth(page, '병아리학생', '🐥', '병아리', 5);
  await assertGrowth(page, '닭학생', '🐔', '닭', 10);

  let eggRow = page.locator('#praiseList .praise-row').filter({ hasText: '알학생' }).first();
  await eggRow.locator('.praise-name').click();
  await assertGrowth(page, '알학생', '🐣', '부화 중', 1);
  eggRow = page.locator('#praiseList .praise-row').filter({ hasText: '알학생' }).first();
  await eggRow.locator('.praise-minus').click();
  await assertGrowth(page, '알학생', '🥚', '알', 0);
  eggRow = page.locator('#praiseList .praise-row').filter({ hasText: '알학생' }).first();
  assert(await eggRow.locator('.praise-minus').isDisabled(), '0점 차감 버튼이 비활성화되지 않았습니다');

  await page.screenshot({ path: 'tools/preview66-praise-growth-tablet.png', fullPage: true });
  await page.locator('#praiseModal .close-btn').click();

  await page.setViewportSize({ width: 390, height: 844 });
  await page.locator('#homeBtn').click();
  await backupShortcut.waitFor({ state: 'visible' });
  const bounds = await backupShortcut.boundingBox();
  assert(bounds && bounds.x >= 0 && bounds.x + bounds.width <= 390, '모바일 홈 메뉴 백업 버튼이 화면 밖으로 벗어났습니다');
  await backupShortcut.click();
  await page.locator('#backupModal').waitFor({ state: 'visible' });
  await page.screenshot({ path: 'tools/preview66-home-backup-mobile.png', fullPage: true });

  assert(consoleErrors.length === 0, '브라우저 오류: ' + consoleErrors.join(' | '));
  console.log(JSON.stringify({
    ok: true,
    version: '5.17.0-preview.66',
    appUrl: APP_URL,
    verified: [
      '홈 하위메뉴 데이터 백업 진입',
      '모바일 홈 메뉴 버튼 화면 내 표시',
      '칭찬 0·1·5·10점 성장 단계',
      '좌석 카드 성장 배지',
      '칭찬 +1·−1 단계 전환',
      '0점 차감 방지'
    ]
  }, null, 2));
  await browser.close();
})().catch(error => {
  console.error(error.stack || error);
  process.exitCode = 1;
});
