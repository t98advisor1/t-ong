const { chromium } = require('C:/Users/place/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');

const APP_URL = process.env.TONG_TEST_URL || 'http://127.0.0.1:4321/';
const CHROME = 'C:/Program Files/Google/Chrome/Application/chrome.exe';
const STUDENT_ID = '1';
const STUDENT_NAME = '김가람';

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

async function getDateKeys(page) {
  return page.evaluate(() => {
    function key(offset) {
      const date = new Date();
      date.setDate(date.getDate() + offset);
      return date.getFullYear() + '-' +
        String(date.getMonth() + 1).padStart(2, '0') + '-' +
        String(date.getDate()).padStart(2, '0');
    }
    return { today: key(0), previous: key(-1) };
  });
}

async function attendanceSnapshot(page, dateKey, classIndex) {
  return page.evaluate(({ dateKey, classIndex, studentId }) => {
    const stored = JSON.parse(localStorage.getItem('seatArrangement_v3') || '{}');
    const classes = stored.allClassData || {};
    const cls = classes[classIndex] || classes[String(classIndex)] || {};
    const attendance = (cls.attendance || {})[dateKey] || {};
    const reasons = (cls.attendReasons || {})[dateKey] || {};
    const memos = (cls.attendMemos || {})[dateKey] || {};
    const owns = (object, key) => Object.prototype.hasOwnProperty.call(object, key);
    return {
      state: attendance[studentId],
      reason: reasons[studentId],
      memo: memos[studentId],
      hasState: owns(attendance, studentId),
      hasReason: owns(reasons, studentId),
      hasMemo: owns(memos, studentId)
    };
  }, { dateKey, classIndex, studentId: STUDENT_ID });
}

async function waitForAttendance(page, dateKey, expected) {
  await page.waitForFunction(({ dateKey, studentId, expected }) => {
    const stored = JSON.parse(localStorage.getItem('seatArrangement_v3') || '{}');
    const cls = ((stored.allClassData || {})['0']) || {};
    const attendance = (cls.attendance || {})[dateKey] || {};
    const reasons = (cls.attendReasons || {})[dateKey] || {};
    const memos = (cls.attendMemos || {})[dateKey] || {};
    if (expected.removed) {
      return !Object.prototype.hasOwnProperty.call(attendance, studentId) &&
        !Object.prototype.hasOwnProperty.call(reasons, studentId) &&
        !Object.prototype.hasOwnProperty.call(memos, studentId);
    }
    if (expected.state !== undefined && attendance[studentId] !== expected.state) return false;
    if (expected.reason !== undefined && reasons[studentId] !== expected.reason) return false;
    if (expected.memo !== undefined && memos[studentId] !== expected.memo) return false;
    return true;
  }, { dateKey, studentId: STUDENT_ID, expected });
}

async function enterSeatView(page) {
  await page.locator('.mode-tab[data-mode="seat"]').click();
  await page.locator('#stage').waitFor({ state: 'visible' });
  await page.locator('#seatingArea .seat-card:not(.empty)').first().waitFor({ state: 'visible' });
}

async function openStudentFromSeat(page) {
  await enterSeatView(page);
  const card = page.locator('#seatingArea .seat-card').filter({ hasText: STUDENT_NAME }).first();
  assert(await card.count() === 1, '좌석표에서 테스트 학생 카드를 찾지 못했습니다');
  await card.click();
  await page.locator('#studentModal').waitFor({ state: 'visible' });
  await page.locator('#studentAttendanceQuick').waitFor({ state: 'visible' });
  assert(
    (await page.locator('#selectedStudentName').innerText()).includes(STUDENT_NAME),
    '좌석에서 클릭한 학생과 열린 팝업의 학생이 다릅니다'
  );
}

async function closeStudentModal(page) {
  await page.locator('#studentModal .close-btn').click();
  await page.locator('#studentModal').waitFor({ state: 'hidden' });
}

async function assertAttendanceView(page, expected) {
  await page.locator('.mode-tab[data-mode="attend"]').click();
  await page.locator('#attendView').waitFor({ state: 'visible' });
  const row = page.locator('#avList .attend-item').filter({ hasText: STUDENT_NAME }).first();
  await row.waitFor({ state: 'visible' });

  assert(
    await row.locator('.at-sel.state').inputValue() === expected.state,
    '출결관리의 학생 상태가 좌석 팝업과 일치하지 않습니다'
  );
  assert(
    await row.locator('.at-sel.reason').inputValue() === expected.reason,
    '출결관리의 사유가 좌석 팝업과 일치하지 않습니다'
  );
  assert(
    await row.locator('.at-memo').inputValue() === expected.memo,
    '출결관리의 메모가 좌석 팝업과 일치하지 않습니다'
  );

  const summary = await page.locator('#avSummary').innerText();
  if (expected.state === '결석') {
    assert(/결석\s*1/.test(summary), '출결관리 상단 결석 집계가 1명이 아닙니다: ' + summary);
  } else {
    assert(/결석\s*0/.test(summary), '출석 복귀 후 결석 집계가 0명이 아닙니다: ' + summary);
  }
}

async function assertStateSelected(page, state) {
  const button = page.locator(
    '#studentAttendanceQuick [data-attendance-state="' + state + '"]'
  );
  await button.waitFor({ state: 'visible' });
  const selected = await button.evaluate(element => {
    return element.getAttribute('aria-pressed') === 'true' ||
      element.classList.contains('active') ||
      element.classList.contains('selected') ||
      element.checked === true;
  });
  assert(selected, '학생 팝업에서 ' + state + ' 상태가 선택되어 보이지 않습니다');
}

async function assertTodayAttendanceAfterViewingPreviousDate(page, dates) {
  const dateInput = page.locator('#avDate');
  await dateInput.fill(dates.previous);
  await dateInput.dispatchEvent('change');
  assert(
    await dateInput.inputValue() === dates.previous,
    '출결관리 조회 날짜가 전날로 바뀌지 않았습니다'
  );

  const previousRow = page.locator('#avList .attend-item').filter({ hasText: STUDENT_NAME }).first();
  await previousRow.waitFor({ state: 'visible' });
  assert(
    await previousRow.locator('.at-sel.state').inputValue() === '지각',
    '전날 출결관리 화면에 시드한 지각 기록이 표시되지 않습니다'
  );

  // 출결관리에서 과거 날짜를 보았더라도 자리배치는 수업 중 사용하는 오늘 출결을 표시해야 합니다.
  await enterSeatView(page);
  const seatCard = page.locator('#seatingArea .seat-card').filter({ hasText: STUDENT_NAME }).first();
  const badge = seatCard.locator('.attend-badge');
  await badge.waitFor({ state: 'visible' });
  assert(
    (await badge.innerText()).trim() === '결',
    '전날 조회 후 자리배치로 돌아오자 오늘 결석 배지가 유지되지 않았습니다'
  );

  await seatCard.click();
  await page.locator('#studentModal').waitFor({ state: 'visible' });
  const quick = page.locator('#studentAttendanceQuick');
  await quick.waitFor({ state: 'visible' });
  assert(
    (await quick.innerText()).includes('오늘 출결'),
    '좌석 학생 팝업에 오늘 출결이라는 대상 날짜 안내가 없습니다'
  );
  await assertStateSelected(page, '결석');
  assert(
    await page.locator('#studentAttendanceReason').inputValue() === '인정',
    '전날 조회 후 좌석 팝업이 오늘 출결 사유를 표시하지 않습니다'
  );
  assert(
    await page.locator('#studentAttendanceMemo').inputValue() === '병원 진료',
    '전날 조회 후 좌석 팝업이 오늘 출결 메모를 표시하지 않습니다'
  );

  await page.screenshot({
    path: 'tools/preview63-seat-attendance-tablet.png',
    fullPage: true
  });
  await closeStudentModal(page);
}

async function assertIsolation(page, dates) {
  const prior = await attendanceSnapshot(page, dates.previous, 0);
  assert(
    prior.state === '지각' && prior.reason === '기타' && prior.memo === '전날 기록',
    '오늘 출결 변경이 같은 학급의 이전 날짜 기록을 훼손했습니다'
  );

  const otherClass = await attendanceSnapshot(page, dates.today, 1);
  assert(
    otherClass.state === '조퇴' &&
      otherClass.reason === '인정' &&
      otherClass.memo === '다른 학급 기록',
    '동일 학생번호를 가진 다른 학급의 출결 기록이 변경되었습니다'
  );
}

async function assertMobileNoHorizontalOverflow(page) {
  const result = await page.evaluate(() => {
    const quick = document.querySelector('#studentAttendanceQuick');
    const modal = document.querySelector('#studentModal .modal-content');
    if (!quick || !modal) return { missing: true };

    const viewportWidth = document.documentElement.clientWidth;
    const quickRect = quick.getBoundingClientRect();
    const modalRect = modal.getBoundingClientRect();
    const controls = Array.from(quick.querySelectorAll(
      '[data-attendance-state], select, input, textarea, button'
    )).map((element, index) => {
      const rect = element.getBoundingClientRect();
      return {
        index,
        tag: element.tagName,
        left: rect.left,
        right: rect.right,
        width: rect.width
      };
    });

    return {
      missing: false,
      viewportWidth,
      documentOverflow: Math.max(
        document.documentElement.scrollWidth,
        document.body ? document.body.scrollWidth : 0
      ) - viewportWidth,
      modalOverflow: modal.scrollWidth - modal.clientWidth,
      quickOverflow: quick.scrollWidth - quick.clientWidth,
      quickLeft: quickRect.left,
      quickRight: quickRect.right,
      modalLeft: modalRect.left,
      modalRight: modalRect.right,
      outOfBounds: controls.filter(control =>
        control.left < -1 || control.right > viewportWidth + 1
      )
    };
  });

  assert(!result.missing, '모바일 검사 중 빠른 출결 영역 또는 학생 팝업을 찾지 못했습니다');
  assert(result.documentOverflow <= 1, '390px 화면에 문서 가로 overflow가 있습니다: ' + JSON.stringify(result));
  assert(result.modalOverflow <= 1, '학생 팝업 내부에 가로 overflow가 있습니다: ' + JSON.stringify(result));
  assert(result.quickOverflow <= 1, '빠른 출결 영역에 가로 overflow가 있습니다: ' + JSON.stringify(result));
  assert(
    result.modalLeft >= -1 && result.modalRight <= result.viewportWidth + 1,
    '학생 팝업이 390px 화면 밖으로 벗어납니다: ' + JSON.stringify(result)
  );
  assert(
    result.quickLeft >= -1 && result.quickRight <= result.viewportWidth + 1,
    '빠른 출결 영역이 390px 화면 밖으로 벗어납니다: ' + JSON.stringify(result)
  );
  assert(
    result.outOfBounds.length === 0,
    '빠른 출결 컨트롤이 390px 화면 밖으로 벗어납니다: ' + JSON.stringify(result.outOfBounds)
  );
}

(async () => {
  const browser = await chromium.launch({ headless: true, executablePath: CHROME });
  const context = await browser.newContext({ viewport: { width: 1024, height: 768 } });
  const consoleErrors = [];

  await context.addInitScript(() => {
    if (sessionStorage.getItem('tongSeatAttendanceTestSeeded') === '1') return;
    sessionStorage.setItem('tongSeatAttendanceTestSeeded', '1');

    function dateKey(offset) {
      const date = new Date();
      date.setDate(date.getDate() + offset);
      return date.getFullYear() + '-' +
        String(date.getMonth() + 1).padStart(2, '0') + '-' +
        String(date.getDate()).padStart(2, '0');
    }

    const today = dateKey(0);
    const previous = dateKey(-1);
    localStorage.clear();
    localStorage.setItem('seatGuide_v2_seen', '1');
    localStorage.setItem('seatArrangement_v3', JSON.stringify({
      settings: { periodStart: previous, periodEnd: today },
      currentClassIndex: 0,
      allClassData: {
        0: {
          className: '테스트 1반',
          students: [
            { id: '1', name: '김가람', gender: '여' },
            { id: '2', name: '이도윤', gender: '남' }
          ],
          currentSeats: {
            '0-0': { studentId: '1', name: '김가람' },
            '0-1': { studentId: '2', name: '이도윤' }
          },
          lockedSeats: {},
          config: { totalDesks: 4, columns: 3, seatType: 'pair', layoutIndex: 0, rows: 1 },
          relations: { close: {}, far: {} },
          studentMemos: {},
          specialNotes: [],
          savedRecords: [],
          attendance: { [previous]: { '1': '지각' } },
          attendReasons: { [previous]: { '1': '기타' } },
          attendMemos: { [previous]: { '1': '전날 기록' } }
        },
        1: {
          className: '테스트 2반',
          students: [{ id: '1', name: '박가람', gender: '여' }],
          currentSeats: { '0-0': { studentId: '1', name: '박가람' } },
          lockedSeats: {},
          config: { totalDesks: 4, columns: 3, seatType: 'pair', layoutIndex: 0, rows: 1 },
          relations: { close: {}, far: {} },
          studentMemos: {},
          specialNotes: [],
          savedRecords: [],
          attendance: { [today]: { '1': '조퇴' } },
          attendReasons: { [today]: { '1': '인정' } },
          attendMemos: { [today]: { '1': '다른 학급 기록' } }
        }
      }
    }));
  });

  await context.route('**/api/usage**', route => route.fulfill({
    status: 200,
    contentType: 'application/json',
    body: JSON.stringify({ today: 1, total: 1 })
  }));

  const page = await context.newPage();
  page.on('console', message => {
    if (message.type() === 'error') consoleErrors.push(message.text());
  });
  page.on('pageerror', error => consoleErrors.push(error.message));

  await page.goto(APP_URL, { waitUntil: 'domcontentloaded' });
  const dates = await getDateKeys(page);

  // 좌석 학생 클릭 → 결석/사유/메모가 각각 즉시 저장되어야 합니다.
  await openStudentFromSeat(page);
  await page.locator(
    '#studentAttendanceQuick [data-attendance-state="결석"]'
  ).click();
  await waitForAttendance(page, dates.today, { state: '결석' });

  await page.locator('#studentAttendanceReason').selectOption('인정');
  await waitForAttendance(page, dates.today, { state: '결석', reason: '인정' });

  const memo = page.locator('#studentAttendanceMemo');
  await memo.fill('병원 진료');
  await memo.press('Tab');
  await waitForAttendance(page, dates.today, {
    state: '결석',
    reason: '인정',
    memo: '병원 진료'
  });

  await assertIsolation(page, dates);
  await closeStudentModal(page);

  // 출결관리 화면에 같은 데이터와 요약이 즉시 반영되어야 합니다.
  await assertAttendanceView(page, {
    state: '결석',
    reason: '인정',
    memo: '병원 진료'
  });

  // 전날 기록을 조회하고 자리배치로 돌아와도 좌석과 팝업은 오늘 결석을 유지해야 합니다.
  await assertTodayAttendanceAfterViewingPreviousDate(page, dates);

  // 새로고침 후에도 localStorage와 팝업 UI가 유지되어야 합니다.
  await page.reload({ waitUntil: 'domcontentloaded' });
  const persisted = await attendanceSnapshot(page, dates.today, 0);
  assert(
    persisted.state === '결석' &&
      persisted.reason === '인정' &&
      persisted.memo === '병원 진료',
    '새로고침 후 출결 기록이 유지되지 않았습니다: ' + JSON.stringify(persisted)
  );

  await openStudentFromSeat(page);
  await assertStateSelected(page, '결석');
  assert(
    await page.locator('#studentAttendanceReason').inputValue() === '인정',
    '새로고침 후 팝업 사유가 복원되지 않았습니다'
  );
  assert(
    await page.locator('#studentAttendanceMemo').inputValue() === '병원 진료',
    '새로고침 후 팝업 메모가 복원되지 않았습니다'
  );

  // 결석 상태의 빠른 출결 팝업이 모바일에서도 가로로 벗어나지 않아야 합니다.
  await page.setViewportSize({ width: 390, height: 844 });
  await assertMobileNoHorizontalOverflow(page);
  await page.screenshot({
    path: 'tools/preview63-seat-attendance-mobile.png',
    fullPage: true
  });
  await page.setViewportSize({ width: 1024, height: 768 });

  // 출석으로 되돌리면 오늘의 attendance/reasons/memos 학생 키가 모두 없어야 합니다.
  await page.locator(
    '#studentAttendanceQuick [data-attendance-state="출석"]'
  ).click();
  await waitForAttendance(page, dates.today, { removed: true });

  const removed = await attendanceSnapshot(page, dates.today, 0);
  assert(
    !removed.hasState && !removed.hasReason && !removed.hasMemo,
    '출석 복귀 후 출결·사유·메모 키가 남아 있습니다: ' + JSON.stringify(removed)
  );
  await assertIsolation(page, dates);
  await closeStudentModal(page);

  await enterSeatView(page);
  const seatCard = page.locator('#seatingArea .seat-card').filter({ hasText: STUDENT_NAME }).first();
  assert(
    await seatCard.locator('.attend-badge').count() === 0,
    '출석 복귀 후 좌석 카드 출결 배지가 남아 있습니다'
  );
  await assertAttendanceView(page, { state: '', reason: '', memo: '' });

  assert(consoleErrors.length === 0, '브라우저 오류: ' + consoleErrors.join(' | '));

  const version = await page.locator('[data-help="version"] small').innerText();
  console.log(JSON.stringify({
    ok: true,
    version,
    appUrl: APP_URL,
    student: STUDENT_ID + '번 ' + STUDENT_NAME,
    date: dates.today,
    verified: [
      '좌석 학생 클릭',
      '결석·사유·메모 즉시 저장',
      '출결관리 양방향 화면 반영',
      '전날 조회 후 좌석·팝업은 오늘 출결 유지',
      '새로고침 유지',
      '출석 복귀 시 세 저장 키 삭제',
      '날짜·학급 격리',
      '모바일 390x844 가로 overflow 없음'
    ]
  }, null, 2));

  await browser.close();
})().catch(error => {
  console.error(error.stack || error);
  process.exitCode = 1;
});
