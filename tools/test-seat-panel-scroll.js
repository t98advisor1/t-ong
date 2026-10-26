// Isolated synthetic-data regression test; never reuses a personal browser profile.
const fs = require('fs');
const path = require('path');
const assert = require('assert/strict');
const { chromium } = require('C:/Users/place/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');

const APP_URL = process.env.TONG_TEST_URL || 'http://127.0.0.1:4321/';
const APP_ORIGIN = new URL(APP_URL).origin;
const OUT = path.resolve(__dirname, '../tmp/seat-panel-v69');
const AUDIT = process.argv.includes('--audit');
const BUILT = process.argv.includes('--built');
const DIST = path.resolve(__dirname, '../dist');
const SHOT_PREFIX = BUILT ? 'built-' : APP_ORIGIN.startsWith('https://') ? 'live-' : '';
const CHROME = 'C:/Program Files/Google/Chrome/Application/chrome.exe';
const VIEWPORTS = AUDIT
  ? [[1280, 600], [1024, 634], [962, 601]]
  : [[390, 844], [844, 390], [962, 601], [1024, 634], [1280, 390], [1280, 600], [1280, 680], [1344, 680], [1280, 800], [1440, 900]];

function seedSyntheticData({ origin }) {
  if (location.origin !== origin || sessionStorage.getItem('tongPanelSyntheticSeeded') === '1') return;
  sessionStorage.setItem('tongPanelSyntheticSeeded', '1');
  localStorage.clear();
  localStorage.setItem('seatGuide_v2_seen', '1');
  localStorage.setItem('seatSound_v1', '0');
  localStorage.setItem('seatPanel_collapsed', '0');
  const students = Array.from({ length: 30 }, (_, i) => ({
    id: String(i + 1), name: '가상학생' + String(i + 1).padStart(2, '0'), gender: i % 2 ? '여' : '남'
  }));
  const seats = {};
  students.forEach((student, i) => {
    seats[Math.floor(i / 6) + '-' + (i % 6)] = { studentId: student.id, name: student.name };
  });
  localStorage.setItem('seatArrangement_v3', JSON.stringify({
    settings: {}, currentClassIndex: 0, allClassData: {
      0: {
        className: '모바일 패널 가상 테스트반', students,
        currentSeats: seats, lockedSeats: {},
        config: { totalDesks: 30, columns: 5, seatType: 'pair', layoutIndex: 0, rows: 5,
          sameGenderSeating: false, numberOrderMode: 'random', numberOrderSeating: false },
        relations: { close: {}, far: {} }, studentMemos: {}, specialNotes: [], savedRecords: [],
        praise: {}, attendance: {}, attendReasons: {}, attendMemos: {}
      }
    }
  }));
}

async function metrics(page) {
  return page.evaluate(() => {
    function rect(node) {
      if (!node) return null;
      const r = node.getBoundingClientRect();
      return { x: r.x, y: r.y, width: r.width, height: r.height, right: r.right, bottom: r.bottom };
    }
    const panel = document.getElementById('seatPanel');
    const teacher = document.getElementById('screenTeacherViewBtn');
    const stage = document.querySelector('.stage');
    return {
      viewport: { width: innerWidth, height: innerHeight },
      document: { scrollWidth: document.documentElement.scrollWidth, clientWidth: document.documentElement.clientWidth,
        scrollHeight: document.documentElement.scrollHeight, scrollY },
      panel: { bounds: rect(panel), overflowY: getComputedStyle(panel).overflowY,
        clientHeight: panel.clientHeight, scrollHeight: panel.scrollHeight, scrollTop: panel.scrollTop },
      stage: { bounds: rect(stage), overflowY: getComputedStyle(stage).overflowY },
      teacher: { bounds: rect(teacher), text: teacher.textContent.trim() }
    };
  });
}

async function wheelPanel(page, delta = 600) {
  const panel = await page.locator('#seatPanel').boundingBox();
  const vp = page.viewportSize();
  if (!panel || panel.x >= vp.width || panel.x + panel.width <= 0 || panel.y >= vp.height) {
    return { attempted: false, reason: 'panel has no reachable viewport intersection' };
  }
  const before = await page.locator('#seatPanel').evaluate(node => node.scrollTop);
  await page.mouse.move(Math.min(vp.width - 6, Math.max(6, panel.x + panel.width / 2)),
    Math.min(vp.height - 8, Math.max(8, panel.y + Math.min(140, panel.height / 2))));
  await page.mouse.wheel(0, delta);
  await page.waitForTimeout(200);
  const after = await page.locator('#seatPanel').evaluate(node => node.scrollTop);
  return { attempted: true, before, after, delta: after - before };
}

async function touchScroll(page, context) {
  // Chromium CDP emits actual touch input; no JS scrollTop mutation.
  await page.locator('#seatPanel').scrollIntoViewIfNeeded();
  const panel = await page.locator('#seatPanel').boundingBox();
  const vp = page.viewportSize();
  if (!panel) return { attempted: false };
  const x = Math.min(vp.width - 12, Math.max(12, panel.x + panel.width - 20));
  const bottom = Math.min(vp.height - 25, panel.y + panel.height - 20);
  const top = Math.max(30, bottom - 250);
  const before = await metrics(page);
  const cdp = await context.newCDPSession(page);
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x, y: bottom }] });
  for (let step = 1; step <= 8; step++) {
    await cdp.send('Input.dispatchTouchEvent', { type: 'touchMove',
      touchPoints: [{ x, y: bottom - (bottom - top) * step / 8 }] });
    await page.waitForTimeout(20);
  }
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
  await page.waitForTimeout(250);
  await cdp.detach();
  const after = await metrics(page);
  return { attempted: true, documentDelta: after.document.scrollY - before.document.scrollY,
    panelDelta: after.panel.scrollTop - before.panel.scrollTop };
}

async function storedConfig(page) {
  return page.evaluate(() => JSON.parse(localStorage.getItem('seatArrangement_v3')).allClassData['0'].config);
}

async function seatMapSnapshot(page) {
  return page.evaluate(() => ({
    saved: JSON.stringify(JSON.parse(localStorage.getItem('seatArrangement_v3')).allClassData['0'].currentSeats),
    rendered: JSON.stringify(Array.from(document.querySelectorAll('#seatingArea .seat-card'))
      .map(card => ({ id: card.dataset.seatId,
        name: card.querySelector('.student-name')?.textContent.trim() || '',
        number: card.querySelector('.student-id')?.textContent.trim() || '' }))
      .sort((a, b) => a.id.localeCompare(b.id)))
  }));
}

async function assertReachable(page, selector, name) {
  const node = page.locator(selector);
  await node.scrollIntoViewIfNeeded();
  const box = await node.boundingBox();
  const vp = page.viewportSize();
  assert(box && box.x >= -1 && box.y >= -1 && box.x + box.width <= vp.width + 1 && box.y + box.height <= vp.height + 1,
    name + ' is clipped: ' + JSON.stringify(box));
  // Trial click checks event interception and clipping without activating print/download.
  await node.click({ trial: true, timeout: 5000 });
  return box;
}

async function testViewport(browser, width, height) {
  const label = width + 'x' + height;
  const isMobile = width <= 900;
  const context = await browser.newContext({ viewport: { width, height },
    serviceWorkers: 'block', hasTouch: true, isMobile });
  const errors = [];
  const externalRequests = [];
  await context.addInitScript(seedSyntheticData, { origin: APP_ORIGIN });
  await context.route('**/*', async route => {
    const request = route.request();
    const url = new URL(request.url());
    if (url.pathname.startsWith('/api/usage')) {
      return route.fulfill({ status: 200, contentType: 'application/json',
        body: JSON.stringify({ todayVisitors: 1, totalVisits: 1, today: 1, total: 1 }) });
    }
    if (url.origin === APP_ORIGIN) {
      if (!BUILT) return route.continue();
      let publicPath;
      try { publicPath = decodeURIComponent(url.pathname); } catch { return route.fulfill({ status: 400, body: '' }); }
      if (publicPath === '/') publicPath = '/index.html';
      const allowed = /^\/(index\.html|style\.css|app\.js|sw\.js|manifest\.webmanifest|icon(?:-maskable)?\.svg)$/.test(publicPath)
        || /^\/(assets|guide)\//.test(publicPath)
        || publicPath === '/.well-known/assetlinks.json';
      const candidate = path.resolve(DIST, '.' + publicPath);
      if (!allowed || !candidate.startsWith(DIST + path.sep) || !fs.existsSync(candidate) || !fs.statSync(candidate).isFile()) {
        return route.fulfill({ status: 404, body: '' });
      }
      const realPath = fs.realpathSync(candidate);
      const realDist = fs.realpathSync(DIST);
      if (!realPath.startsWith(realDist + path.sep)) return route.fulfill({ status: 403, body: '' });
      const types = { '.html': 'text/html', '.css': 'text/css', '.js': 'application/javascript',
        '.webmanifest': 'application/manifest+json', '.json': 'application/json', '.svg': 'image/svg+xml',
        '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.webp': 'image/webp' };
      return route.fulfill({ status: 200, contentType: types[path.extname(candidate)] || 'application/octet-stream', body: fs.readFileSync(candidate) });
    }
    externalRequests.push(url.origin + url.pathname);
    const type = request.resourceType();
    return route.fulfill({ status: 200, contentType: type === 'script' ? 'application/javascript' : 'text/plain',
      body: type === 'script' ? '/* blocked external request in isolated UI test */' : '' });
  });
  const page = await context.newPage();
  page.on('pageerror', error => errors.push(error.message));
  page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
  const result = { viewport: label, mode: AUDIT ? 'audit-current' : 'regression', built: BUILT, errors, externalRequests };
  try {
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded' });
    await page.locator('.mode-tab[data-mode="seat"]').click();
    await page.locator('#seatPanel').waitFor({ state: 'visible' });
    await page.waitForTimeout(200);
    result.initial = await metrics(page);
    result.wheel = await wheelPanel(page);
    result.afterWheel = await metrics(page);
    await page.screenshot({ path: path.join(OUT, SHOT_PREFIX + (AUDIT ? 'audit-' : 'after-') + label + '.png'), fullPage: true });
    if (AUDIT) return result;

    const head = page.locator('#seatPanel .sp-head');
    const settings = page.locator('#spSettings');
    assert.equal(await head.getAttribute('aria-expanded'), 'false', 'seat settings should initially be collapsed');
    assert.equal(await settings.isVisible(), false, 'settings group is not collapsed');
    for (const selector of ['#spBody', '#sameGenderSeating', '#numberOrderOption', '#spEditBtn']) {
      assert.equal(await page.locator(selector).isVisible(), false, selector + ' remains outside the collapsed group');
    }
    await head.click();
    assert(await settings.isVisible(), 'seat settings cannot expand');
    const choices = await page.locator('.sp-order-choices label').evaluateAll(nodes => nodes.map(node => {
      const r = node.getBoundingClientRect(); return { x:r.x, y:r.y, bottom:r.bottom, height:r.height };
    }));
    assert.equal(choices.length, 3);
    assert(choices.every(choice => choice.height >= 44), 'order choices need touch-sized targets');
    assert(choices[0].bottom <= choices[1].y && choices[1].bottom <= choices[2].y, 'order choices are not vertical');
    assert(choices.every(choice => Math.abs(choice.x - choices[0].x) <= 1), 'order choices do not align vertically');
    result.verticalOrderChoices = choices;
    assert((await page.locator('#numberOrderSummary').innerText()).includes('무작위'), 'summary misses current random mode');
    assert.equal(await page.locator('input[name="numberOrderMode"]').count(), 3, 'existing radio inputs changed');
    assert.equal(await page.locator('#numberOrderOption').count(), 1, 'existing number fieldset disappeared');

    if (width === 1280 && height === 600) {
      result.expanded = await metrics(page);
      result.expandedWheel = await wheelPanel(page);
      assert(result.expanded.panel.scrollHeight > result.expanded.panel.clientHeight, 'test needs overflowing desktop panel');
      assert(result.expandedWheel.attempted && result.expandedWheel.delta > 0, 'real mouse wheel did not scroll right panel');
    }
    if (isMobile || (width > 1080 && height <= 700)) {
      result.touch = await touchScroll(page, context);
      assert(result.touch.attempted && (result.touch.documentDelta > 0 || result.touch.panelDelta > 0),
        'actual touch drag did not scroll the mobile document or panel');
    }
    const mapBefore = await seatMapSnapshot(page);
    result.teacherToggleBounds = await assertReachable(page, '#screenTeacherViewBtn', 'teacher viewpoint');
    await page.screenshot({ path: path.join(OUT, SHOT_PREFIX + 'teacher-accessible-' + label + '.png') });
    await page.locator('#screenTeacherViewBtn').click();
    assert.equal(await page.locator('#screenTeacherViewBtn').getAttribute('aria-pressed'), 'true');
    assert((await page.locator('#screenTeacherViewBtn').innerText()).includes('학생 시점'), 'student return button missing');
    await page.locator('#screenTeacherViewBtn').click();
    assert.equal(await page.locator('#screenTeacherViewBtn').getAttribute('aria-pressed'), 'false');
    assert.deepEqual(await seatMapSnapshot(page), mapBefore, 'viewpoint altered seat assignments');

    result.backupBounds = await assertReachable(page, '#spBackupBtn', 'data backup');
    await page.locator('#spBackupBtn').click();
    await page.locator('#backupModal.show').waitFor({ state: 'visible' });
    await page.locator('#backupModal [data-modal="backupModal"]').click();
    const printDrop = page.locator('.sp-actions .sp-drop');
    await printDrop.locator('.drop-t').click();
    await printDrop.locator('.sp-menu.open').waitFor({ state: 'visible' });
    result.studentPrintBounds = await assertReachable(page, '.sp-actions [data-proxy="studentViewBtn"]', 'student print option');
    result.teacherPrintBounds = await assertReachable(page, '.sp-actions [data-proxy="teacherViewBtn"]', 'teacher print option');
    await page.screenshot({ path: path.join(OUT, SHOT_PREFIX + 'print-accessible-' + label + '.png') });
    await printDrop.locator('.drop-t').click();

    await page.locator('input[name="numberOrderMode"][value="row"]').check();
    assert((await page.locator('#numberOrderSummary').innerText()).includes('학번 가로순'), 'summary did not update to row mode');
    assert.equal((await storedConfig(page)).numberOrderMode, 'row', 'row mode not persisted');
    await head.focus();
    await head.press('Enter');
    assert.equal(await head.getAttribute('aria-expanded'), 'false');
    assert.equal(await settings.isVisible(), false);
    await head.press('Enter');
    assert(await page.locator('input[name="numberOrderMode"][value="row"]').isChecked());
    await page.locator('input[name="numberOrderMode"][value="column"]').check();
    assert((await page.locator('#numberOrderSummary').innerText()).includes('학번 세로순'), 'summary did not update to column mode');
    assert.equal((await storedConfig(page)).numberOrderMode, 'column');
    await page.reload({ waitUntil: 'domcontentloaded' });
    await page.locator('.mode-tab[data-mode="seat"]').click();
    assert.equal(await head.getAttribute('aria-expanded'), 'true', 'expanded preference was not restored');
    assert.equal((await storedConfig(page)).numberOrderMode, 'column', 'reload reset saved config');
    assert((await page.locator('#numberOrderSummary').innerText()).includes('학번 세로순'), 'saved column summary not restored');
    assert(await page.locator('input[name="numberOrderMode"][value="column"]').isChecked());
    await page.locator('input[name="numberOrderMode"][value="random"]').check();
    assert((await page.locator('#numberOrderSummary').innerText()).includes('무작위'));
    assert(!(await page.locator('#sameGenderSeating').isDisabled()));
    await head.click();
    assert.equal(await settings.isVisible(), false);
    await page.reload({ waitUntil: 'domcontentloaded' });
    await page.locator('.mode-tab[data-mode="seat"]').click();
    assert.equal(await settings.isVisible(), false, 'collapsed preference was not restored');
    result.final = await metrics(page);
    assert(result.final.document.scrollWidth <= result.final.document.clientWidth + 1,
      'horizontal document overflow: ' + JSON.stringify(result.final.document));
    assert.equal(errors.length, 0, 'browser errors: ' + errors.join(' | '));
    result.ok = true;
    await page.locator('#screenTeacherViewBtn').scrollIntoViewIfNeeded();
    await page.screenshot({ path: path.join(OUT, SHOT_PREFIX + 'verified-' + label + '.png'), fullPage: true });
    return result;
  } catch (error) {
    result.ok = false;
    result.failure = error.stack || String(error);
    await page.screenshot({ path: path.join(OUT, SHOT_PREFIX + 'failed-' + label + '.png'), fullPage: true }).catch(() => {});
    return result;
  } finally {
    await context.close();
  }
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const browser = await chromium.launch({ headless: true, executablePath: CHROME });
  const results = [];
  try {
    for (const [width, height] of VIEWPORTS) {
      const result = await testViewport(browser, width, height);
      results.push(result);
      console.log(JSON.stringify(result));
    }
  } finally { await browser.close(); }
  const report = { appUrl: APP_URL, audit: AUDIT, built: BUILT, syntheticDataOnly: true,
    serviceWorkers: 'blocked', externalNetwork: 'mocked', results };
  fs.writeFileSync(path.join(OUT, SHOT_PREFIX + (AUDIT ? 'audit-current.json' : 'regression-report.json')), JSON.stringify(report, null, 2));
  if (results.some(result => result.ok === false)) process.exitCode = 1;
})().catch(error => { console.error(error.stack || error); process.exitCode = 1; });
