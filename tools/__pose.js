/* 스크린샷 촬영 전용 — 각 장면을 실제 UI 조작으로 재현합니다. 배포 제외(.vercelignore). */
(function () {
  var shot = new URLSearchParams(location.search).get('s') || '04';
  window.print = function () {};   // 인쇄 대화상자 차단

  function $(id) { return document.getElementById(id); }
  function click(id) { var el = $(id); if (el) el.click(); }
  function hide(id) { var el = $(id); if (el) el.classList.remove('show'); }

  function css(text) {
    var st = document.createElement('style');
    st.textContent = text;
    document.head.appendChild(st);
  }

  var annoStyled = false;
  function annoStyle() {
    if (annoStyled) return;
    annoStyled = true;
    css('.shot-ring{position:fixed;border:3px solid #ff4d6d;border-radius:12px;' +
        'box-shadow:0 0 0 6px rgba(255,77,109,0.20);z-index:2000;pointer-events:none}' +
        '.shot-ring.blue{border-color:#3b82f6;box-shadow:0 0 0 6px rgba(59,130,246,0.20)}' +
        '.shot-tag{position:fixed;background:#ff4d6d;color:#fff;font-weight:700;font-size:13px;' +
        'padding:6px 12px;border-radius:8px;z-index:2001;white-space:nowrap;' +
        'box-shadow:0 4px 14px rgba(255,77,109,0.4)}' +
        '.shot-tag.blue{background:#3b82f6;box-shadow:0 4px 14px rgba(59,130,246,0.4)}' +
        '.shot-badge{position:fixed;z-index:2002;width:26px;height:26px;border-radius:50%;' +
        'background:#ff4d6d;color:#fff;font-weight:800;font-size:13px;display:flex;' +
        'align-items:center;justify-content:center;box-shadow:0 3px 10px rgba(0,0,0,0.3)}' +
        '.shot-badge.blue{background:#3b82f6}');
  }

  // annotate(대상, 라벨, {where:'below'|'above'|'left', color:'blue', badge:1})
  function annotate(target, label, opt) {
    opt = opt || {};
    var el = typeof target === 'string' ? ($(target) || document.querySelector(target)) : target;
    if (!el) return;
    annoStyle();
    var r = el.getBoundingClientRect();
    var tint = opt.color === 'blue' ? ' blue' : '';

    var ring = document.createElement('div');
    ring.className = 'shot-ring' + tint;
    ring.style.left = (r.left - 6) + 'px';
    ring.style.top = (r.top - 6) + 'px';
    ring.style.width = (r.width + 12) + 'px';
    ring.style.height = (r.height + 12) + 'px';
    document.body.appendChild(ring);

    if (opt.badge) {
      var b = document.createElement('div');
      b.className = 'shot-badge' + tint;
      b.textContent = opt.badge;
      b.style.left = (r.left - 15) + 'px';
      b.style.top = (r.top - 15) + 'px';
      document.body.appendChild(b);
    }

    if (!label) return;
    var tag = document.createElement('div');
    tag.className = 'shot-tag' + tint;
    tag.textContent = label;
    document.body.appendChild(tag);
    var tw = tag.getBoundingClientRect().width;
    var where = opt.where || 'below';
    var x = r.left + r.width / 2 - tw / 2;
    var y = where === 'above' ? r.top - 40 : r.bottom + 12;
    if (where === 'left') { x = r.left - tw - 16; y = r.top + r.height / 2 - 16; }
    if (where === 'right') { x = r.right + 16; y = r.top + r.height / 2 - 16; }
    tag.style.left = Math.max(8, Math.min(window.innerWidth - tw - 8, x)) + 'px';
    tag.style.top = Math.max(8, y) + 'px';
  }

  function firstSeatCard(index) {
    var cards = document.querySelectorAll('.seat-card:not(.empty)');
    return cards[index || 0];
  }


  function showSeatStage() {
    var seatTab = document.querySelector('.mode-tab[data-mode="seat"]');
    if (seatTab) seatTab.classList.add('active');
    var wizard = document.getElementById('setupWizard');
    var stage = document.getElementById('stage');
    var attend = document.getElementById('attendView');
    var actionBar = document.querySelector('.action-bar');
    if (wizard) { wizard.classList.add('hidden'); wizard.style.display = 'none'; }
    if (attend) { attend.hidden = true; attend.classList.add('mode-screen-hidden'); attend.style.display = 'none'; }
    if (stage) { stage.hidden = false; stage.classList.remove('mode-screen-hidden'); stage.style.display = 'flex'; }
    if (actionBar) { actionBar.hidden = true; actionBar.style.display = 'none'; }
    window.scrollTo(0, 0);
  }
  var POSES = {
    // 최신 메인 화면의 네 영역
    '00': function () {
      annotate('.app-header', '① 상단 메뉴 · 학급 선택', { where: 'below', badge: 1 });
      annotate('.side-rail', '② 설정 · 기록 메뉴', { where: 'right', badge: 2, color: 'blue' });
      annotate('.classroom', '③ 좌석표', { where: 'below', badge: 3 });
      annotate('.seat-panel', '④ 좌석 설정 · 추첨', { where: 'left', badge: 4, color: 'blue' });
    },

    // 상단 탭과 드롭다운
    '0A': function () {
      click('homeBtn');
      annotate('homeBtn', '첫 화면 · 사용 설명서', { where: 'below', badge: 1 });
      document.querySelectorAll('.mode-tab').forEach(function (el, i) {
        annotate(el, null, { badge: i + 2, color: i === 0 ? '' : 'blue' });
      });
      annotate('.header-right', '학급과 적용 기간', { where: 'below' });
    },

    // 추가 기능 메뉴
    '0B': function () {
      click('extraBtn');
      setTimeout(function () {
        annotate('extraMenu', '수업 도구를 한곳에서 엽니다', { where: 'below', badge: 1 });
      }, 80);
    },

    // 일정관리 ① — 전체화면 달력
    '22': function () {
      click('calendarBtn');
      setTimeout(function () {
        annotate('.cal-month-pane', '① 월간 일정 확인', { where: 'right', badge: 1 });
        annotate('#calGrid .cal-day.selected', '② 날짜 더블클릭으로 일정 입력', { where: 'right', badge: 2, color: 'blue' });
        annotate('calHelpBtn', '③ 일정관리 캐러셀 사용법', { where: 'below', badge: 3 });
      }, 140);
    },

    // 일정관리 ② — 속성과 상담 입력
    '22A': function () {
      click('calendarBtn');
      setTimeout(function () {
        click('calAddBtn');
        var property = document.getElementById('calPropertyInput');
        if (property) { property.value = '상담'; property.dispatchEvent(new Event('change')); }
        var student = document.getElementById('calStudentInput'); if (student && student.options.length > 3) student.selectedIndex = 3;
        click('calAddPropertyBtn');
        var customProperty = document.getElementById('calNewPropertyInput'); if (customProperty) customProperty.value = '연수';
        annotate('.cal-property-add', '① 새 속성을 직접 입력해 추가', { where: 'left', badge: 1 });
        annotate('.cal-consult-fields', '② 상담 대상·상태·후속 조치', { where: 'left', badge: 2, color: 'blue' });
      }, 160);
    },

    // 일정관리 ③ — 가져오기·내보내기
    '22B': function () {
      click('calendarBtn');
      setTimeout(function () {
        click('calTransferBtn');
        setTimeout(function () {
          click('calSampleBtn');
          annotate('#calTransferPanel', '시트 가져오기 · ICS · 상담기록 엑셀/복사', { where: 'above', badge: 1 });
        }, 100);
      }, 140);
    },

    // 교실 게임 메뉴
    '24': function () {
      click('gameBtn');
      setTimeout(function () {
        annotate('gameMenu', '수업 전환과 참여를 위한 교실 게임', { where: 'left', badge: 1 });
      }, 100);
    },

    // 도움말 메뉴
    '25': function () {
      click('helpBtn');
      setTimeout(function () {
        annotate('helpMenu', '버전 · 업데이트 · 개발자 응원 · 학교정보 · 문의', { where: 'left', badge: 1, color: 'blue' });
      }, 100);
    },

    // 도움말의 오늘의 학교 실험 기능
    '26': function () {
      click('helpBtn');
      setTimeout(function () {
        var item = document.querySelector('[data-help="school"]');
        if (item) item.click();
        setTimeout(function () {
          var date = document.getElementById('schoolTodayDate');
          if (date) { date.value = '2026-04-01'; date.dispatchEvent(new Event('change')); }
        }, 120);
      }, 100);
    },
    // 그룹형 즐겨찾기
    '23': function () {
      click('favBtn');
      setTimeout(function () {
        var categories = document.querySelectorAll('#favMenu .fav-category');
        categories.forEach(function (category, index) { category.open = index < 2; });
        annotate(categories[0], '① 수학관련 사이트 + 직접 추가', { where: 'left', badge: 1 });
        annotate(categories[1], '② 수업 및 클래스 + 직접 추가', { where: 'left', badge: 2, color: 'blue' });
        annotate(categories[categories.length - 1], '③ 나만의 즐겨찾기는 맨 아래', { where: 'below', badge: 3 });
      }, 100);
    },
    // 좌측 접기 메뉴
    '0C': function () {
      click('railToggle');
      setTimeout(function () {
        annotate('sideRail', '학기 초 설정과 기록 메뉴', { where: 'right', badge: 1 });
        annotate('railBackup', '저장 기록 · 백업 · 초기화', { where: 'right', badge: 2, color: 'blue' });
      }, 80);
    },

    // 우측 좌석 설정 패널
    '0D': function () {
      showSeatStage();
      annotate('.sp-roster', '① 명렬표 등록', { where: 'left', badge: 1 });
      annotate('spBody', '② 현재 좌석 조건', { where: 'left', badge: 2, color: 'blue' });
      annotate('spEditBtn', '③ 조건 바꾸기', { where: 'left', badge: 3, color: 'blue' });
      annotate('.sp-actions', '④ 추첨 · 저장 · 출력', { where: 'left', badge: 4 });
    },

    '01': function () { var seat=document.querySelector('.mode-tab[data-mode="seat"]'); if(seat)seat.click(); },

    // 전체화면 타이머
    '17': function () {
      var open = document.querySelector('[data-open="timer"]');
      if (open) open.click();
      var c = document.querySelector('#tsPresets .ts-preset[data-ts-min="5"]');
      if (c) c.click();
    },

    // 설정 메뉴 (번호 매김)
    '18': function () { click('settingsBtn'); },

    // 설정 탭 2 (자리 배치 안내)
    '19': function () {
      click('settingsBtn');
      var t = document.querySelectorAll('.set-tab')[1];
      if (t) t.click();
    },

    // 출결 입력 — 날짜·요약·상태·사유·메모·내보내기
    '20': function () {
      var tab = document.querySelector('.mode-tab[data-mode="attend"]');
      if (tab) tab.click();
      setTimeout(function () {
        annotate('.av-date', '① 날짜를 바꾸면 지난 출결도 조회·수정', { where: 'below', badge: 1 });
        annotate('#avSummary', '② 오늘 출결을 즉시 집계', { where: 'below', badge: 2, color: 'blue' });
        annotate('#avList', '③ 상태 · 사유 · 메모를 학생별 기록', { where: 'above', badge: 3 });
        annotate('.av-actions', '④ 월별 엑셀 · 구글시트 내보내기', { where: 'below', badge: 4, color: 'blue' });
      }, 120);
    },

    // 자동 검증 전용 — 태블릿에서 출결 → 자리배치 왕복 시 한 화면만 표시
    'tabcheck': function () {
      var stage = document.getElementById('stage');
      var attend = document.getElementById('attendView');
      var attendTab = document.querySelector('.mode-tab[data-mode="attend"]');
      var seatTab = document.querySelector('.mode-tab[data-mode="seat"]');
      if (attendTab) attendTab.click();
      setTimeout(function () {
        document.body.dataset.attendStageDisplay = getComputedStyle(stage).display;
        document.body.dataset.attendViewDisplay = getComputedStyle(attend).display;
        if (seatTab) seatTab.click();
        setTimeout(function () {
          document.body.dataset.seatStageDisplay = getComputedStyle(stage).display;
          document.body.dataset.seatViewDisplay = getComputedStyle(attend).display;
        }, 120);
      }, 120);
    },

    // 출결관리 학생 이름 → 종합 기록 → 관찰·세특
    '20B': function () {
      var tab = document.querySelector('.mode-tab[data-mode="attend"]');
      if (tab) tab.click();
      setTimeout(function () {
        var student = document.querySelector('.attend-student-link');
        if (student) student.click();
        setTimeout(function () {
          click('studentSpecialQuickBtn');
          setTimeout(function () {
            annotate('studentPraiseMinusBtn', '① 출결 확인 중 칭찬 +1 / −1', { where: 'below', badge: 1 });
            annotate('#specialTab .special-note-form', '② 날짜 · 교과 · 기록 유형 · 관찰 내용', { where: 'left', badge: 2, color: 'blue' });
          }, 100);
        }, 100);
      }, 120);
    },
    // 1인 1역
    '21': function () {
      click('dutyBtn');
      setTimeout(function () { click('dutyAssignBtn'); }, 150);
    },

    // 1인 1역 엑셀 연동
    '21A': function () {
      click('dutyBtn');
      setTimeout(function () {
        annotate('dutyTemplateBtn', '① 두 가지 양식 받기', { where: 'below', badge: 1 });
        annotate('dutyUploadBtn', '② 작성한 엑셀 올리기', { where: 'below', badge: 2, color: 'blue' });
        annotate('.duty-hint', '좌석표 · 학생 카드 · 통합 기록에 즉시 연동', { where: 'below', badge: 3 });
      }, 100);
    },

    '02': function () {
      click('settingsBtn'); click('openRosterBtn'); hide('settingsModal');
    },

    '03': function () {
      click('settingsBtn'); click('openExcelUploadBtn'); hide('settingsModal');
    },

    '04': function () { showSeatStage(); /* 좌석 화면 그대로 */ },

    // 추첨 시작 스플래시 (카운트다운 화면)
    '05S': function () { click('shuffleBtn'); },

    '05': function () {
      click('shuffleBtn');
      setTimeout(function () { click('splashSkipBtn'); }, 120);
    },

    '06': function () {
      var card = firstSeatCard(4);
      if (card) card.click();
    },

    '07': function () {
      var card = firstSeatCard(4);
      if (card) card.click();
      var tab = document.querySelector('.tab-btn[data-tab="relations"]');
      if (tab) tab.click();
    },

    '08': function () {
      var seatTab = document.querySelector('.mode-tab[data-mode="seat"]');
      if (seatTab) seatTab.click();
      var wizard = document.getElementById('setupWizard');
      var stage = document.getElementById('stage');
      if (wizard) { wizard.classList.add('hidden'); wizard.style.display = 'none'; }
      if (stage) { stage.hidden = false; stage.classList.remove('mode-screen-hidden'); stage.style.display = 'flex'; }
      var actionBar = document.querySelector('.action-bar');
      if (actionBar) { actionBar.hidden = true; actionBar.style.display = 'none'; }
      var locked = document.querySelector('.seat-card.locked .lock-btn');
      if (locked) annotate(locked.closest('.seat-card'), '🔒 고정된 자리 — 다시 추첨해도 유지됩니다');
    },

    '09': function () { click('historyBtn'); },

    '10': function () {
      click('railToggle');
      setTimeout(function () {
        click('railBackup');
        setTimeout(function () {
          annotate('[data-proxy="downloadHistoryBtn"]', '학생별 좌석 · 짝 이력을 엑셀로 저장', { where: 'right' });
        }, 80);
      }, 80);
    },

    // 인쇄 스타일은 @media print 안에만 있으므로, 화면에서도 같은 모습이 나오도록 규칙을 복제합니다
    '11': function () {
      click('studentViewBtn');
      css('body > *:not(#printArea){display:none !important}' +
          '#printArea{display:block !important;position:static !important;padding:28px;background:#fff}' +
          '.print-header{text-align:center;margin-bottom:20px}' +
          '.print-header h1{font-size:1.8rem;margin-bottom:8px}' +
          '.print-header p{font-size:1rem;color:#555}' +
          '.print-seating{display:flex;flex-direction:column;gap:8px;margin:20px 0}' +
          '.print-row{display:flex;gap:18px;justify-content:center}' +
          '.print-group{display:flex;gap:4px}' +
          '.print-seat{width:80px;height:56px;border:1.5px solid #333;display:flex;flex-direction:column;' +
            'align-items:center;justify-content:center;font-size:0.9rem;border-radius:4px}' +
          '.print-seat.male{background:#e3f2fd}' +
          '.print-seat.female{background:#fce4ec}' +
          '.print-seat.empty{background:#f5f5f5}' +
          '.print-seat.locked{border:2.5px solid #ff9800}' +
          '.print-seat-id{font-size:0.75rem;color:#666}' +
          '.print-seat-name{font-weight:600;font-size:1rem}' +
          '.print-desk{text-align:center;padding:10px 20px;background:#667eea;color:#fff;max-width:160px;' +
            'margin:16px auto;border-radius:6px;font-size:1rem;font-weight:600}' +
          '.print-desk-gap{height:16px}' +
          '.print-legend{display:flex;justify-content:center;gap:20px;margin-top:20px;font-size:0.9rem}' +
          '.print-legend-item{display:flex;align-items:center;gap:8px}' +
          '.print-legend-box{width:22px;height:22px;border:1px solid #333;border-radius:3px}' +
          '.print-legend-box.male{background:#e3f2fd}' +
          '.print-legend-box.female{background:#fce4ec}' +
          '.print-footer{display:flex;justify-content:space-between;margin-top:30px;font-size:0.85rem;color:#666;padding:0 20px}');
    },

    // 설정 메뉴 6개 항목
    '13': function () {
      click('settingsBtn');
      var items = document.querySelectorAll('#settingsModal .settings-item');
      items.forEach(function (el, i) { annotate(el, null, { badge: i + 1, color: i === 5 ? '' : 'blue' }); });
    },

    // 모둠 배치 좌석표
    '14': function () { showSeatStage(); /* 시드가 모둠 설정으로 렌더 */ },

    // 알림장
    '15': function () {
      click('noticeBtn');
      setTimeout(function () { annotate('#noticeModal .modal-content', '작성 후 크게 보기로 교실 TV에 표시', { where: 'left', color: 'blue' }); }, 80);
    },

    // 칭찬 포인트
    '16': function () {
      click('praiseBtn');
      var card = document.querySelector('.seat-card .praise-badge');
      if (card) annotate(card.closest('.seat-card'), '좌석표에도 표시', { where: 'below' });
    },

    '12': function () {
      click('settingsBtn'); click('openBackupBtn'); hide('settingsModal');
      click('exportBtn');
    }
  };

  window.addEventListener('load', function () {
    setTimeout(function () {
      var fn = POSES[shot];
      if (fn) fn();
    }, 260);
  });
})();
