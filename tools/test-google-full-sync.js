const { chromium } = require('C:/Users/place/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');

const APP_URL = process.env.TONG_TEST_URL || 'http://127.0.0.1:4321/';
const CHROME = 'C:/Program Files/Google/Chrome/Application/chrome.exe';

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

(async () => {
  const browser = await chromium.launch({ headless: true, executablePath: CHROME });
  const context = await browser.newContext({ viewport: { width: 1024, height: 768 } });
  const requests = [];
  const consoleErrors = [];
  const linkedUpdated = '2026-08-20T00:00:00.000Z';
  const linkedStamp = Date.parse(linkedUpdated);

  await context.addInitScript(({ linkedUpdated, linkedStamp }) => {
    localStorage.setItem('seatGuide_v2_seen', '1');
    localStorage.setItem('seatArrangement_v3', JSON.stringify({
      settings: { periodStart: '2026-09-01', periodEnd: '2026-12-31' },
      currentClassIndex: 0,
      allClassData: {
        0: {
          className: '테스트반',
          students: [{ id: 1, name: '테스트학생', gender: '남' }],
          currentSeats: {}, lockedSeats: {},
          config: { totalDesks: 30, columns: 5, seatType: 'pair', layoutIndex: 0, rows: 6 },
          relations: { close: {}, far: {} }, studentMemos: {}, specialNotes: [], savedRecords: []
        }
      }
    }));
    localStorage.setItem('seatCalendar_v1', JSON.stringify([{
      id: 'local-linked', title: '기존 연결 일정', start: '2026-10-05', end: '2026-10-05',
      property: '없음', target: '전체', memo: '', createdAt: linkedStamp, updatedAt: linkedStamp
    }]));
    localStorage.setItem('tongGoogleCalendarConfig_v1', JSON.stringify({
      clientId: '123456789-test.apps.googleusercontent.com', activeAccountId: 'acct-1',
      autoSync: false, includeConsult: false, includeMemo: false
    }));
    localStorage.setItem('tongGoogleCalendarSync_v1', JSON.stringify({
      'acct-1': {
        accountId: 'acct-1', email: 'teacher@example.com', calendarId: 'primary',
        calendarName: 'Google 기본 캘린더', syncMode: 'primary-v1', syncToken: 'stale-token',
        links: { 'local-linked': { googleId: 'google-linked', lastLocalUpdated: linkedStamp, lastGoogleUpdated: linkedUpdated } },
        pendingDeletes: [], lastSyncAt: linkedStamp
      }
    }));
    window.google = {
      accounts: {
        oauth2: {
          initTokenClient(options) {
            return {
              requestAccessToken() {
                setTimeout(() => options.callback({ access_token: 'fake-token', expires_in: 3600 }), 0);
              }
            };
          },
          revoke(_token, callback) { if (callback) callback(); }
        }
      }
    };
  }, { linkedUpdated, linkedStamp });

  await context.route('**/api/usage**', route => route.fulfill({
    status: 200, contentType: 'application/json', body: JSON.stringify({ today: 1, total: 1 })
  }));
  await context.route('https://openidconnect.googleapis.com/v1/userinfo', route => route.fulfill({
    status: 200, contentType: 'application/json', body: JSON.stringify({ sub: 'acct-1', email: 'teacher@example.com' })
  }));
  await context.route('https://www.googleapis.com/calendar/v3/**', async route => {
    const request = route.request();
    const url = new URL(request.url());
    requests.push({ method: request.method(), url: request.url(), pathname: url.pathname, search: url.search });
    if (request.method() !== 'GET') {
      return route.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify({ id: 'unexpected-write', updated: new Date().toISOString() }) });
    }
    if (url.searchParams.get('syncToken')) {
      return route.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify({ items: [], nextSyncToken: 'stale-token' }) });
    }
    const events = [
      {
        id: 'google-linked', status: 'confirmed', summary: '기존 연결 일정', updated: linkedUpdated,
        start: { date: '2026-10-05' }, end: { date: '2026-10-06' },
        extendedProperties: { private: { tongLocalId: 'local-linked', tongProperty: '없음', tongTarget: '전체' } }
      },
      {
        id: 'google-new', status: 'confirmed', summary: 'Google 신규 일정', updated: '2026-09-01T01:00:00.000Z',
        start: { date: '2026-10-12' }, end: { date: '2026-10-13' }
      }
    ];
    return route.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify({ items: events, nextSyncToken: 'fresh-token' }) });
  });

  const page = await context.newPage();
  page.on('console', msg => {
    if (msg.type() === 'error' && !msg.text().includes('Failed to load resource: net::ERR_NETWORK_ACCESS_DENIED')) {
      consoleErrors.push(msg.text());
    }
  });
  page.on('pageerror', error => consoleErrors.push(error.message));
  await page.goto(APP_URL, { waitUntil: 'domcontentloaded' });
  await page.locator('#calendarBtn').click();
  await page.locator('#calGooglePanel > summary').click();

  const incrementalResponse = page.waitForResponse(response => {
    const url = new URL(response.url());
    return url.pathname.endsWith('/calendars/primary/events') && url.searchParams.get('syncToken') === 'stale-token';
  });
  await page.locator('#googleCalConnectBtn').click();
  await incrementalResponse;
  await page.locator('#googleCalFullSyncBtn').waitFor({ state: 'visible' });
  await page.waitForFunction(() => !document.querySelector('#googleCalFullSyncBtn').disabled);

  const requestCut = requests.length;
  page.once('dialog', dialog => dialog.accept());
  const fullResponse = page.waitForResponse(response => {
    const url = new URL(response.url());
    return url.pathname.endsWith('/calendars/primary/events') && !url.searchParams.has('syncToken');
  });
  await page.locator('#googleCalFullSyncBtn').click();
  await fullResponse;
  await page.waitForFunction(() => document.querySelector('#toast').textContent.includes('전체 다시 가져오기 완료'));

  const after = requests.slice(requestCut);
  const fullGets = after.filter(call => call.method === 'GET' && call.pathname.endsWith('/calendars/primary/events'));
  const writes = after.filter(call => ['POST', 'PATCH', 'DELETE'].includes(call.method));
  assert(fullGets.length === 1, `전체 events.list 요청 수 오류: ${fullGets.length}`);
  const fullUrl = new URL(fullGets[0].url);
  assert(!fullUrl.searchParams.has('syncToken'), '전체 가져오기 요청에 syncToken이 포함됨');
  assert(fullUrl.searchParams.get('showDeleted') === 'true', 'showDeleted=true 누락');
  assert(fullUrl.searchParams.get('singleEvents') === 'true', 'singleEvents=true 누락');
  assert(writes.length === 0, `불필요한 Google 쓰기 요청 발생: ${JSON.stringify(writes)}`);

  const result = await page.evaluate(() => ({
    events: JSON.parse(localStorage.getItem('seatCalendar_v1') || '[]'),
    states: JSON.parse(localStorage.getItem('tongGoogleCalendarSync_v1') || '{}'),
    toast: document.querySelector('#toast').textContent,
    version: document.querySelector('[data-help="version"] small').textContent,
    fullButtonText: document.querySelector('#googleCalFullSyncBtn').textContent
  }));
  assert(result.version === '5.17.0-preview.66', `버전 오류: ${result.version}`);
  assert(result.fullButtonText.trim() === 'Google 일정 전체 다시 가져오기', '버튼 문구 오류');
  assert(result.events.filter(event => event.id === 'local-linked').length === 1, '기존 연결 일정이 중복되거나 사라짐');
  assert(result.events.filter(event => event.id === 'gcal-google-new').length === 1, 'Google 신규 일정 가져오기 실패 또는 중복');
  assert(result.states['acct-1'].syncToken === 'fresh-token', '새 syncToken 저장 실패');
  assert(result.states['acct-1'].links['local-linked'].googleId === 'google-linked', '기존 링크 보존 실패');
  assert(result.states['acct-1'].links['gcal-google-new'].googleId === 'google-new', '신규 Google 링크 저장 실패');
  assert(result.toast.includes('Google에서 1개') && result.toast.includes('Google로 0개'), `완료 문구 오류: ${result.toast}`);

  await page.screenshot({ path: 'tools/preview66-google-full-sync-tablet.png', fullPage: true });
  await page.locator('#calGooglePanel').screenshot({ path: 'tools/preview66-google-full-sync-panel-tablet.jpg', type: 'jpeg', quality: 55 });
  await page.setViewportSize({ width: 390, height: 844 });
  await page.locator('#googleCalFullSyncBtn').scrollIntoViewIfNeeded();
  await page.screenshot({ path: 'tools/preview66-google-full-sync-mobile.png', fullPage: true });
  await page.locator('#calGooglePanel').screenshot({ path: 'tools/preview66-google-full-sync-panel-mobile.jpg', type: 'jpeg', quality: 55 });

  assert(consoleErrors.length === 0, `브라우저 오류: ${consoleErrors.join(' | ')}`);
  console.log(JSON.stringify({
    ok: true,
    fullRequest: fullGets[0],
    writes: writes.length,
    eventIds: result.events.map(event => event.id),
    syncToken: result.states['acct-1'].syncToken,
    toast: result.toast,
    screenshots: ['tools/preview66-google-full-sync-tablet.png', 'tools/preview66-google-full-sync-mobile.png']
  }, null, 2));
  await browser.close();
})().catch(error => {
  console.error(error.stack || error);
  process.exitCode = 1;
});
