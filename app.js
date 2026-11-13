(function() {
'use strict';

var APP_VERSION = '5.17.0-preview.73';
var STORAGE_KEY = 'seatArrangement_v3';
var MAX_CLASSES = 10;

var ALL_LAYOUTS = {
  3: [[2,2], [1,2], [2,1]],
  4: [[2,2], [2,2], [1,2,1]],
  5: [[2,2,2], [2,3], [3,2], [2,1,2]],
  6: [[2,2,2], [3,3], [2,2,2]],
  7: [[2,2,2,2], [2,3,2], [3,2,2], [2,2,3]]
};
var GROUP_SIZES = [4, 5, 6];
var TRIPLE_LAYOUTS = {
  3: [[3,3,3]],
  4: [[3,3,3,3]],
  5: [[3,3,3,3,3]],
  6: [[3,3,3,3,3,3]],
  7: [[3,3,3,3,3,3,3]]
};

var students = [];
var currentSeats = {};
var lockedSeats = {};
var relations = { close: {}, far: {} };
var studentMemos = {};
var savedRecords = [];
var config = { totalDesks: 30, columns: 5, seatType: 'pair', layoutIndex: 0, rows: 6 };
var settings = { periodStart: '', periodEnd: '' };
var allClassData = {};
var currentClassIndex = 0;

var selectedStudent = null;
var selectedEmptySeatId = null;
var tempRelations = { close: [], far: [] };
var pickerMode = 'close';
var isLotteryRunning = false;
var bouncingBalls = [];
var lotteryAnimationId = null;
var draggedSeat = null;
var assignSeatStudent = null;
var editingClassIndex = 0;
var parsedExcelData = null;
var isTeacherScreenView = false;
var toastHideTimer = null;

function sortStudentsByName(list) {
  return (Array.isArray(list) ? list.slice() : []).sort(function(a, b) {
    var an = String(a && a.name || '').trim();
    var bn = String(b && b.name || '').trim();
    var compared = 0;
    try {
      compared = an.localeCompare(bn, 'ko-KR', { sensitivity: 'base', numeric: true });
    } catch (e) {
      compared = an.localeCompare(bn);
    }
    if (compared !== 0) return compared;
    try {
      return String(a && a.id || '').localeCompare(String(b && b.id || ''), undefined, { numeric: true });
    } catch (e2) {
      return String(a && a.id || '').localeCompare(String(b && b.id || ''));
    }
  });
}

function sortStudentsByNumber(list) {
  return (Array.isArray(list) ? list.slice() : []).sort(function(a, b) {
    var aid = String(a && a.id || '').trim();
    var bid = String(b && b.id || '').trim();
    var compared = 0;
    try {
      compared = aid.localeCompare(bid, 'ko-KR', { sensitivity: 'base', numeric: true });
    } catch (e) {
      compared = aid.localeCompare(bid);
    }
    if (compared !== 0) return compared;
    return String(a && a.name || '').localeCompare(String(b && b.name || ''), 'ko-KR');
  });
}

// 추첨 연출 (v3.1)
var LOTTERY_BASE_MS = 13000;                                  // '보통' 속도 기준 총 진행 시간
var SPEED_MULTIPLIERS = { slow: 1.9, normal: 1, fast: 0.45 };
var lotterySpeed = 'normal';
var GUIDE_SEEN_KEY = 'seatGuide_v2_seen';
var SOUND_PREF_KEY = 'seatSound_v1';
var guideIndex = 0;

// ===================================================================
// CloudSheetSync — 구글 계정 영구 로그인 및 교사용 단일 통합 시트 양방향 동기화
// - 개인 기기 특성을 고려하여 한 번 로그인 시 수동 로그아웃 전까지 영구 유지
// - 교사 각자의 구글 시트 웹앱(Apps Script) URL로 데이터를 안전하게 분리 저장
// - 오프라인/미로그인 시 로컬 대기열에 안전하게 쌓고, 재로그인 시 일괄 자동 전송
// ===================================================================
var CloudSheetSync = (function () {
  var AUTH_KEY = 'tong_google_auth';
  var SHEET_URL_KEY = 'tong_gsheet_url';
  var PENDING_KEY = 'tong_pending_sync_v1';
  var LAST_SYNC_KEY = 'tong_last_sync_time';
  var autoSyncTimer = null;
  var isSyncing = false;

  function el(id) { return document.getElementById(id); }

  function getAuth() {
    try {
      var item = localStorage.getItem(AUTH_KEY);
      return item ? JSON.parse(item) : null;
    } catch (e) {
      return null;
    }
  }

  function setAuth(data) {
    try {
      if (data) {
        localStorage.setItem(AUTH_KEY, JSON.stringify(data));
      } else {
        localStorage.removeItem(AUTH_KEY);
      }
    } catch (e) {}
  }

  function getSheetUrl() {
    try {
      return localStorage.getItem(SHEET_URL_KEY)
        || localStorage.getItem('SR_GAS_MASTER_URL')
        || localStorage.getItem('swk_gsheet_url')
        || localStorage.getItem('swk_gsheet_homeroom_url')
        || localStorage.getItem('swk_gsheet_subject_url')
        || '';
    } catch (e) {
      return '';
    }
  }

  function setSheetUrl(url) {
    try {
      var u = (url || '').trim();
      var keys = [SHEET_URL_KEY, 'SR_GAS_MASTER_URL', 'swk_gsheet_url', 'swk_gsheet_homeroom_url', 'swk_gsheet_subject_url'];
      keys.forEach(function (k) {
        if (u) localStorage.setItem(k, u);
        else localStorage.removeItem(k);
      });
    } catch (e) {}
  }

  function getPending() {
    try {
      var p = localStorage.getItem(PENDING_KEY);
      return p ? JSON.parse(p) : [];
    } catch (e) {
      return [];
    }
  }

  function savePending(list) {
    try {
      localStorage.setItem(PENDING_KEY, JSON.stringify(list || []));
    } catch (e) {}
  }

  function parseJwt(token) {
    try {
      var base64Url = token.split('.')[1];
      var base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
      var jsonPayload = decodeURIComponent(atob(base64).split('').map(function (c) {
        return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
      }).join(''));
      return JSON.parse(jsonPayload);
    } catch (e) {
      return null;
    }
  }

  function metaClientId() {
    var node = document.querySelector('meta[name="google-oauth-client-id"]');
    return node ? String(node.content || '').trim() : '';
  }

  function updateUI() {
    var auth = getAuth();
    var sheetUrl = getSheetUrl();
    var pending = getPending();
    var lastSync = localStorage.getItem(LAST_SYNC_KEY);

    var badge = el('headerSyncBtn');
    var badgeText = el('headerSyncText');
    var avatar = el('syncUserAvatar');
    var nameEl = el('syncUserName');
    var emailEl = el('syncUserEmail');
    var permBadge = el('syncPermBadge');
    var loginSec = el('syncLoginSection');
    var logoutBtn = el('syncLogoutBtn');
    var stateText = el('syncStateText');
    var lastTimeText = el('syncLastTimeText');
    var pendingText = el('syncPendingCountText');
    var urlInput = el('tongSheetUrlInput');

    if (urlInput && !urlInput.value && sheetUrl) {
      urlInput.value = sheetUrl;
    }

    if (pendingText) {
      pendingText.textContent = pending.length + '건';
    }

    if (lastTimeText) {
      lastTimeText.textContent = lastSync ? new Date(parseInt(lastSync, 10)).toLocaleTimeString('ko-KR') : '기록 없음';
    }

    if (isSyncing) {
      if (badge) {
        badge.className = 'header-sync-btn sync-busy';
      }
      if (badgeText) badgeText.textContent = '동기화 중...';
      if (stateText) stateText.textContent = '클라우드 동기화 진행 중...';
      return;
    }

    if (auth && auth.email) {
      // 영구 로그인 활성 상태
      if (permBadge) permBadge.hidden = false;
      if (loginSec) loginSec.hidden = true;
      if (logoutBtn) logoutBtn.hidden = false;

      if (avatar) {
        if (auth.picture) {
          avatar.innerHTML = '<img src="' + auth.picture + '" alt="프로필" referrerpolicy="no-referrer">';
        } else {
          avatar.textContent = (auth.name || 'G').charAt(0);
        }
      }
      if (nameEl) nameEl.textContent = auth.name || '선생님';
      if (emailEl) emailEl.textContent = auth.email;

      if (sheetUrl) {
        if (badge) badge.className = 'header-sync-btn sync-online';
        if (badgeText) badgeText.textContent = (auth.name ? auth.name + 'T' : 'Google') + ' | 동기화 켜짐';
        if (stateText) stateText.textContent = '정상 동기화 (영구 유지)';
      } else {
        if (badge) badge.className = 'header-sync-btn sync-offline';
        if (badgeText) badgeText.textContent = (auth.name ? auth.name + 'T' : 'Google') + ' | 시트 등록 필요';
        if (stateText) stateText.textContent = '시트 웹앱 URL 입력 필요';
      }
    } else {
      // 로컬 단독 / 미로그인 상태
      if (permBadge) permBadge.hidden = true;
      if (loginSec) loginSec.hidden = false;
      if (logoutBtn) logoutBtn.hidden = true;

      if (avatar) {
        avatar.innerHTML = 'G';
      }
      if (nameEl) nameEl.textContent = '이 브라우저에만 저장 중';
      if (emailEl) emailEl.textContent = 'Google 통합 로그인 시 계정과 캘린더가 함께 연결됩니다';

      if (badge) badge.className = 'header-sync-btn sync-offline';
      if (badgeText) {
        badgeText.textContent = pending.length > 0 ? ('미연결 (' + pending.length + '건)') : 'Google 연동';
      }
      if (stateText) {
        stateText.textContent = '미연결 (현재 기기만 저장)';
      }
      renderGoogleButton();
    }

    updateCalendarSyncSubUI(auth);
    updateDashSyncHubUI();
  }

  function updateCalendarSyncSubUI(auth) {
    var calAccText = el('syncCalAccountText');
    var calBadge = el('syncCalModeBadge');
    var useMainBtn = el('syncUseMainCalAccountBtn');
    var connectNowBtn = el('syncCalConnectNowBtn');
    if (!calAccText) return;

    if (typeof GoogleCalendarSync !== 'undefined' && GoogleCalendarSync.getAccountSummary) {
      var info = GoogleCalendarSync.getAccountSummary();
      if (info.customCalAccount) {
        if (calBadge) {
          calBadge.textContent = '캘린더 별도 계정';
          calBadge.style.background = '#fef3c7';
          calBadge.style.color = '#92400e';
        }
        if (useMainBtn) useMainBtn.hidden = !(auth && auth.email && info.email !== auth.email);
      } else {
        if (calBadge) {
          calBadge.textContent = '통합 연동';
          calBadge.style.background = '#e0e7ff';
          calBadge.style.color = '#3730a3';
        }
        if (useMainBtn) useMainBtn.hidden = true;
      }

      if (info.connected && info.email) {
        calAccText.innerHTML = '연결된 캘린더: <b>' + info.email + '</b> (' + (info.customCalAccount ? '별도 지정 계정' : '기본 계정 통합') + ')';
        if (connectNowBtn) connectNowBtn.textContent = '캘린더 지금 동기화';
      } else if (info.email) {
        calAccText.innerHTML = '캘린더 대상 계정: <b>' + info.email + '</b> · <span style="color:#b45309;">권한 연결 대기</span>';
        if (connectNowBtn) connectNowBtn.textContent = '캘린더 권한 연결';
      } else if (auth && auth.email) {
        calAccText.innerHTML = '기본 계정(<b>' + auth.email + '</b>)으로 캘린더 통합 연동 가능';
        if (connectNowBtn) connectNowBtn.textContent = '캘린더 연결 · 동기화';
      } else {
        calAccText.textContent = 'Google 로그인 한 번으로 교무수첩과 캘린더가 함께 동기화됩니다.';
        if (connectNowBtn) connectNowBtn.textContent = '캘린더 연결';
      }
    }
  }

  function renderGoogleButton() {
    var container = el('googleBtnContainer');
    if (!container) return;
    container.innerHTML = '';

    // 1순위: 한 번의 클릭으로 교무수첩 영구 로그인 + Google 캘린더 OAuth 토큰을 동시에 획득하는 통합 로그인 버튼
    var unifiedBtn = document.createElement('button');
    unifiedBtn.type = 'button';
    unifiedBtn.className = 'btn';
    unifiedBtn.style.cssText = 'width:100%; display:flex; align-items:center; justify-content:center; gap:8px; padding:10px 14px; font-weight:700; font-size:0.86rem; border-radius:999px; background:#ffffff; color:#1f2937; border:1px solid #cbd5e1; box-shadow:0 1px 3px rgba(0,0,0,0.08); cursor:pointer; margin-bottom:6px;';
    unifiedBtn.innerHTML = '<span style="font-weight:900; color:#4285F4; font-size:1rem;">G</span><span>Google 통합 로그인 (교무수첩·캘린더 동시 연동)</span>';
    unifiedBtn.addEventListener('click', function () {
      if (typeof GoogleCalendarSync !== 'undefined' && GoogleCalendarSync.loginUnified) {
        GoogleCalendarSync.loginUnified();
      }
    });
    container.appendChild(unifiedBtn);

    var gisWrap = document.createElement('div');
    gisWrap.style.cssText = 'display:flex; justify-content:center;';
    container.appendChild(gisWrap);

    if (window.google && google.accounts && google.accounts.id) {
      try {
        google.accounts.id.renderButton(gisWrap, {
          theme: 'outline',
          size: 'medium',
          text: 'signin_with',
          shape: 'pill',
          locale: 'ko',
          width: 240
        });
      } catch (e) {
        console.warn('GIS renderButton error:', e);
      }
    }
  }

  function setUnifiedAuthFromUser(user, triggerCalendarSync) {
    if (!user || !user.email) return;
    var existing = getAuth();
    var auth = {
      email: user.email,
      name: user.name || (existing && existing.name) || user.email.split('@')[0],
      picture: user.picture || (existing && existing.picture) || '',
      sub: user.sub || (existing && existing.sub) || '',
      loggedInAt: Date.now()
    };
    setAuth(auth);
    updateUI();
    if (getSheetUrl()) {
      syncNow(false);
    }
    if (triggerCalendarSync && typeof GoogleCalendarSync !== 'undefined' && GoogleCalendarSync.syncWithAccount) {
      GoogleCalendarSync.syncWithAccount(auth.email, auth.sub);
    }
  }

  function handleCredential(res) {
    if (!res || !res.credential) return;
    var payload = parseJwt(res.credential);
    if (!payload || !payload.email) {
      showToast('구글 로그인 토큰을 분석하지 못했습니다.', 'error');
      return;
    }
    var auth = {
      email: payload.email,
      name: payload.name || payload.email.split('@')[0],
      picture: payload.picture || '',
      sub: payload.sub || '',
      loggedInAt: Date.now()
    };
    setAuth(auth);
    updateUI();
    showToast(auth.name + ' 선생님 계정으로 로그인되었습니다 (캘린더 자동 연동 진행).', 'success');

    // 기본 Google 계정으로 일정관리 캘린더 자동 동기화 연동
    if (typeof GoogleCalendarSync !== 'undefined' && GoogleCalendarSync.syncWithAccount) {
      GoogleCalendarSync.syncWithAccount(auth.email, auth.sub);
    }

    if (getSheetUrl()) {
      syncNow(false);
    }
  }

  function initGIS() {
    var clientId = metaClientId() || '777531321307-609ro9rson4sqjqakvjjogib885u26n3.apps.googleusercontent.com';
    function setup() {
      if (!window.google || !google.accounts || !google.accounts.id) return;
      try {
        google.accounts.id.initialize({
          client_id: clientId,
          callback: handleCredential,
          auto_select: false,
          cancel_on_tap_outside: true
        });
        if (!getAuth()) {
          renderGoogleButton();
        }
      } catch (e) {
        console.warn('Google Identity initialize error:', e);
      }
    }

    if (window.google && google.accounts && google.accounts.id) {
      setup();
    } else {
      var checkCount = 0;
      var interval = setInterval(function () {
        checkCount++;
        if (window.google && google.accounts && google.accounts.id) {
          clearInterval(interval);
          setup();
        } else if (checkCount > 30) {
          clearInterval(interval);
        }
      }, 300);
    }
  }

  function queueRecord(item) {
    if (!item) return;
    var list = getPending();
    list.push(item);
    savePending(list);
    updateUI();

    if (getAuth() && getSheetUrl()) {
      if (autoSyncTimer) clearTimeout(autoSyncTimer);
      autoSyncTimer = setTimeout(function () {
        syncNow(false);
      }, 4000);
    }
  }

  // ------------------------------------------------------------------
  // 시트 동기화 상태 — 이 기기가 '시트에 이미 있다'고 알고 있는 것
  //  obs : 받아 봤거나 보낸 관찰 기록(지운 기록이 Pull 로 되살아나지 않게)
  //  rows: 출결·칭찬·1인1역·학생메모의 시트 줄 이력. 키별로 ['날짜|내용', ...] (시트에 쌓인 순서)
  // 시트는 줄을 덧붙이기만 하는 기록장이라, 줄이 쌓인 순서를 '나중에 쓴 것이 이긴다'의 기준으로 씁니다.
  // ------------------------------------------------------------------
  var SYNC_STATE_KEY = 'tongSheetSyncState_v1';
  // Apps Script 가 [6_수업관찰_기록] 탭으로 보내는 기록 유형. 그 밖의 유형은 '수업관찰'로 보내고 내용 앞에 [유형]을 붙입니다.
  var SHEET_OBS_CATEGORIES = ['수업관찰', '교과세특', '수업참여', '수업 참여', '발표·토론', '실험·실습', '과제·보고서'];
  var NOTE_CATEGORIES = ['수업 참여', '학업 역량', '탐구 활동', '발표·토론', '협업·소통', '기타'];
  var SHEET_ATT_STATES = ['결석', '지각', '조퇴', '결과'];
  var SHEET_ATT_REASONS = ['질병', '인정', '미인정', '기타'];
  // 예전 전송분은 내용 뒤에 " (과목: 수학 (2-2))" 가 붙어 시트에 저장됩니다. (반 이름에 괄호가 한 겹 들어 있어도 인식)
  var SUBJECT_SUFFIX_RE = /\s\(과목: [^()]*\(((?:[^()]|\([^()]*\))*)\)\)$/;
  var REVISION_RE = /\s\[수정 \d+\]$/;
  var pushInFlight = null;

  function getSyncState(sheetUrl) {
    var state = null;
    try { state = JSON.parse(localStorage.getItem(SYNC_STATE_KEY) || 'null'); } catch (e) {}
    // 다른 시트로 바꾸면 그 시트에 대해서는 아는 것이 없으므로 새로 시작합니다.
    if (!state || typeof state !== 'object' || state.sheet !== sheetUrl) state = { sheet: sheetUrl };
    if (!state.obs || typeof state.obs !== 'object') state.obs = {};
    if (!state.rows || typeof state.rows !== 'object') state.rows = {};
    return state;
  }
  function saveSyncState(state) {
    try { localStorage.setItem(SYNC_STATE_KEY, JSON.stringify(state)); } catch (e) {}
  }

  function hashKey(text) {
    var h = 5381;
    for (var i = 0; i < text.length; i++) h = ((h << 5) + h + text.charCodeAt(i)) | 0;
    return (h >>> 0).toString(36) + '.' + text.length.toString(36);
  }

  function syncClasses() {
    var list = [];
    for (var i = 0; i < MAX_CLASSES; i++) {
      var d = allClassData[i];
      if (d) list.push({ index: i, data: d, name: d.className || ((i + 1) + '반') });
    }
    return list;
  }
  // 현재 보고 있는 반의 학생 메모는 전역 변수(studentMemos)가 원본입니다.
  function memosOf(cls) {
    if (cls.index === currentClassIndex) return studentMemos;
    if (!cls.data.studentMemos) cls.data.studentMemos = {};
    return cls.data.studentMemos;
  }

  function stripSubjectSuffix(content) {
    var text = String(content || '').trim();
    var m;
    while ((m = text.match(SUBJECT_SUFFIX_RE))) text = text.slice(0, m.index).trim();
    return text;
  }
  // 같은 기록인지 비교할 때 쓰는 내용: 시트가 덧붙인 " (과목: … (반))" · " (반)" 을 떼어 냅니다.
  function noteCompareText(content, classNames) {
    var text = String(content || '').trim();
    for (;;) {
      var m = text.match(SUBJECT_SUFFIX_RE);
      if (m) { text = text.slice(0, m.index).trim(); continue; }
      var hit = '';
      for (var i = 0; i < classNames.length; i++) {
        var suffix = ' (' + classNames[i] + ')';
        if (classNames[i] && text.length > suffix.length && text.slice(-suffix.length) === suffix) { hit = suffix; break; }
      }
      if (!hit) break;
      text = text.slice(0, -hit.length).trim();
    }
    return text;
  }
  function noteKey(studentId, date, compareText) {
    return String(studentId || '') + '|' + String(date || '') + '|' + compareText;
  }

  function rowsKey(kind, className, studentId, date) {
    return kind + '|' + className + '|' + (date || '') + '|' + studentId;
  }
  function entryDetail(entry) {
    var text = String(entry || '');
    var at = text.indexOf('|');
    return (at >= 0 ? text.slice(at + 1) : text).replace(REVISION_RE, '');
  }

  // 시트의 마지막 값과 다를 때만 한 줄을 보냅니다.
  function queueRow(state, key, date, detail, base, items, commits) {
    var hist = state.rows[key] || [];
    if (hist.length && entryDetail(hist[hist.length - 1]) === detail) return;
    var sendDetail = detail;
    // Apps Script 는 (날짜·학번·구분·내용)이 같은 줄을 다시 받지 않으므로, 예전 값으로 되돌아갈 때는 [수정 n]을 붙여 구분합니다.
    if (hist.indexOf(date + '|' + detail) >= 0) sendDetail = detail + ' [수정 ' + (hist.length + 1) + ']';
    var item = { detail: sendDetail };
    Object.keys(base).forEach(function (name) { item[name] = base[name]; });
    items.push(item);
    commits.push({ key: key, entry: date + '|' + sendDetail });
  }

  // 모든 반의 기록을 시트로 보낼 형태로 모읍니다. commits 는 전송에 성공했을 때 동기화 상태에 반영할 줄입니다.
  function collectCurrentClassRecords(state) {
    state = state || getSyncState(getSheetUrl());
    var records = [];
    var commits = [];
    var nowStr = formatDate(new Date());
    var classes = syncClasses();
    var classNames = classes.map(function (c) { return c.name; });

    classes.forEach(function (cls) {
      var d = cls.data;
      var cName = cls.name;
      var roster = d.students || [];
      var studentMap = {};
      roster.forEach(function (st) { studentMap[String(st.id)] = st.name || ''; });

      // 1) 관찰·세특 기록 → [6_수업관찰_기록] · [7_AI세특_작성]
      if (Array.isArray(d.specialNotes)) {
        var sent = {};
        d.specialNotes.forEach(function (sn) {
          var content = stripSubjectSuffix(sn.content);
          if (!content) return;
          // 예전 Pull 이 만든 사본(내용 뒤에 과목·반이 붙은 것)은 원본과 같은 기록이므로 한 번만 보냅니다.
          var key = noteKey(sn.studentId, sn.date, noteCompareText(sn.content, classNames));
          if (sent[key]) return;
          sent[key] = true;
          var category = sn.category || '수업관찰';
          var known = SHEET_OBS_CATEGORIES.indexOf(category) >= 0;
          records.push({
            timestamp: sn.date ? (sn.date + ' 12:00:00') : (nowStr + ' 12:00:00'),
            studentId: sn.studentId || '',
            studentName: sn.studentName || studentMap[String(sn.studentId)] || '',
            className: cName,
            subject: sn.subject || '',
            category: known ? category : '수업관찰',
            detail: known ? content : ('[' + category + '] ' + content),
            memo: sn.subject ? ('과목: ' + sn.subject + ' (' + cName + ')') : cName
          });
        });
      }

      // 2) 학생별 일반 관찰 메모 → [8_출결_칭찬_기록]
      var memos = memosOf(cls);
      roster.forEach(function (st) {
        var key = rowsKey('M', cName, st.id);
        var memoText = String(memos[st.id] || '').trim();
        if (!memoText && !state.rows[key]) return;
        queueRow(state, key, nowStr, memoText || '(메모 삭제)', {
          timestamp: nowStr + ' 12:00:00', studentId: st.id, studentName: st.name || '', className: cName, category: '학생관찰', memo: cName
        }, records, commits);
      });

      // 3) 출결 (출석이 아닌 상태) → [8_출결_칭찬_기록]
      var attendance = d.attendance || {};
      var markedKeys = {};
      Object.keys(attendance).forEach(function (dk) {
        var attObj = attendance[dk] || {};
        var rzObj = (d.attendReasons && d.attendReasons[dk]) || {};
        var memObj = (d.attendMemos && d.attendMemos[dk]) || {};
        roster.forEach(function (st) {
          var status = attObj[st.id];
          if (!status || status === '출석') return;
          var key = rowsKey('A', cName, st.id, dk);
          markedKeys[key] = true;
          queueRow(state, key, dk, status + (rzObj[st.id] ? ' (' + rzObj[st.id] + ')' : ''), {
            timestamp: dk + ' 09:00:00', studentId: st.id, studentName: st.name || '', className: cName, category: '출결',
            memo: memObj[st.id] || (cName + ' 출결 기록')
          }, records, commits);
        });
      });
      // 시트에는 결석·지각 등으로 올라가 있는데 이 기기에서 출석으로 고친 날 → 정정 줄을 보내 다른 기기에도 전합니다.
      var attPrefix = 'A|' + cName + '|';
      Object.keys(state.rows).forEach(function (key) {
        if (key.indexOf(attPrefix) !== 0 || markedKeys[key]) return;
        var rest = key.slice(attPrefix.length);
        var cut = rest.indexOf('|');
        var dk = rest.slice(0, cut);
        var sid = rest.slice(cut + 1);
        if (!Object.prototype.hasOwnProperty.call(studentMap, sid)) return;
        var hist = state.rows[key];
        if (!hist.length || entryDetail(hist[hist.length - 1]).indexOf('출석') === 0) return;
        queueRow(state, key, dk, '출석 (정정)', {
          timestamp: dk + ' 09:00:00', studentId: sid, studentName: studentMap[sid], className: cName, category: '출결', memo: cName + ' 출결 정정'
        }, records, commits);
      });

      // 4) 칭찬 포인트 (누적 점수가 바뀌었을 때만) → [8_출결_칭찬_기록] · [3_창체_행특]
      var praise = d.praise || {};
      roster.forEach(function (st) {
        var key = rowsKey('P', cName, st.id);
        var score = Number(praise[st.id]) || 0;
        if (score <= 0 && !state.rows[key]) return;
        var stage = score >= 10 ? '닭 단계' : (score >= 5 ? '병아리 단계' : '부화 중 단계');
        queueRow(state, key, nowStr, score > 0 ? ('+' + score + '점 (' + stage + ')') : '+0점', {
          timestamp: nowStr + ' 15:00:00', studentId: st.id, studentName: st.name || '', className: cName, category: '칭찬',
          memo: cName + ' 칭찬 포인트 누적 ' + score + '점'
        }, records, commits);
      });

      // 5) 1인 1역 (역할이 바뀌었을 때만) → [8_출결_칭찬_기록] · [3_창체_행특]
      var duties = d.duties || {};
      roster.forEach(function (st) {
        var key = rowsKey('D', cName, st.id);
        var roleName = String(duties[st.id] || '').trim();
        if (!roleName && !state.rows[key]) return;
        queueRow(state, key, nowStr, roleName ? ('1인1역: ' + roleName) : '1인1역: (해제)', {
          timestamp: nowStr + ' 09:00:00', studentId: st.id, studentName: st.name || '', className: cName, category: '학급역할',
          memo: roleName ? (cName + ' 학급 1인1역 역할 수행 (' + roleName + ')') : (cName + ' 1인1역 해제')
        }, records, commits);
      });
    });

    return { items: records, commits: commits };
  }

  // 시트에 이미 있는 줄 하나를 그대로 다시 만듭니다. (Apps Script 가 중복으로 보고 새 줄을 만들지 않음)
  function knownRowAsItem(state) {
    var categoryOf = { A: '출결', P: '칭찬', D: '학급역할', M: '학생관찰' };
    var keys = Object.keys(state.rows);
    for (var i = 0; i < keys.length; i++) {
      var parts = keys[i].split('|');
      var hist = state.rows[keys[i]];
      if (parts.length !== 4 || !categoryOf[parts[0]] || !hist || !hist.length) continue;
      var entry = String(hist[hist.length - 1]);
      var at = entry.indexOf('|');
      if (at < 0) continue;
      return { timestamp: entry.slice(0, at) + ' 09:00:00', studentId: parts[3], studentName: '', className: parts[1], category: categoryOf[parts[0]], detail: entry.slice(at + 1), memo: '' };
    }
    return null;
  }

  // 이 기기의 기록을 시트로 보냅니다. 성공하면 보낸 건수를 돌려주고, 실패하면 오류를 던집니다.
  // 이미 보내는 중이면 그 전송이 끝나기를 기다립니다. (같은 줄을 두 번 보내지 않도록)
  // rosterOnly: 보낼 기록이 없어도 명렬표를 등록하기 위해 요청할지 여부
  function pushToSheet(sheetUrl, rosterOnly) {
    if (pushInFlight) return pushInFlight;
    pushInFlight = (async function () {
      var auth = getAuth();
      var state = getSyncState(sheetUrl);
      var built = collectCurrentClassRecords(state);
      var allItems = getPending().concat(built.items);

      var seen = {};
      var uniqueItems = [];
      for (var k = 0; k < allItems.length; k++) {
        var it = allItems[k];
        var key = (it.timestamp || '') + '_' + (it.studentId || '') + '_' + (it.category || '') + '_' + (it.detail || '');
        if (!seen[key]) {
          seen[key] = true;
          uniqueItems.push(it);
        }
      }

      var profileObj = collectTeacherProfile();
      var profileSig = hashKey(JSON.stringify({
        school: profileObj.school,
        isNonHomeroom: profileObj.isNonHomeroom,
        homeroom: profileObj.homeroom,
        subjectClasses: profileObj.subjectClasses,
        timetable: profileObj.timetable,
        customClubs: profileObj.customClubs,
        sheetMemos: profileObj.sheetMemos
      }));
      var profileChanged = (state.profile !== profileSig);
      var rosterList = Array.isArray(students) ? students : [];
      var rosterSig = hashKey(currentClassIndex + ':' + rosterList.map(function (st) { return st.id + '=' + st.name; }).join(','));
      var rosterChanged = (rosterList.length > 0 && state.roster !== rosterSig);
      if (!uniqueItems.length && !profileChanged && !rosterChanged) {
        return 0;  // 기록·명렬표·시간표 설정 모두 변동 없음 → 요청하지 않음
      }

      var curClass = allClassData[currentClassIndex] || {};
      var payload = {
        action: 'sync_tong_notebook',
        userEmail: auth ? auth.email : '',
        userName: auth ? auth.name : '',
        className: curClass.className || ((currentClassIndex + 1) + '반'),
        students: rosterList,
        teacherProfile: profileObj,
        items: uniqueItems
      };

      var res = await fetch(sheetUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload)
      });
      var data = null;
      try { data = await res.json(); } catch (e) {}
      if (!(res.ok || (data && data.success))) {
        throw new Error(data && data.message ? data.message : '동기화 응답 실패');
      }

      built.commits.forEach(function (commit) {
        if (!state.rows[commit.key]) state.rows[commit.key] = [];
        state.rows[commit.key].push(commit.entry);
      });
      state.roster = rosterSig;
      state.profile = profileSig;
      saveSyncState(state);
      savePending([]);
      localStorage.setItem(LAST_SYNC_KEY, String(Date.now()));
      return uniqueItems.length || ((profileChanged || rosterChanged) ? 1 : 0);
    })();
    var settle = function () { pushInFlight = null; };
    pushInFlight.then(settle, settle);
    return pushInFlight;
  }

  async function syncNow(manual) {
    var sheetUrl = getSheetUrl();
    if (!sheetUrl) {
      if (manual) {
        showToast('동기화할 구글 시트 웹앱 URL을 먼저 입력하고 저장해 주세요.', 'error');
        toggleMenu(true);
      }
      return;
    }

    isSyncing = true;
    updateUI();

    try {
      var count = await pushToSheet(sheetUrl, true);
      isSyncing = false;
      updateUI();
      if (window.parent && window.parent !== window) {
        try { window.parent.postMessage({ type: 'tong-notebook-synced', count: count }, '*'); } catch (e) {}
      }
      if (manual) {
        showToast(count
          ? '이 기기의 기록을 구글 시트로 보냈습니다 (' + count + '건). 다른 기기에서는 [시트에서 가져오기]를 누르세요.'
          : '새로 보낼 기록이 없습니다. 시트가 이미 최신입니다.', 'success', null, null, 4000);
      }
    } catch (err) {
      console.warn('Sync error:', err);
      isSyncing = false;
      updateUI();
      if (manual) {
        showToast('구글 시트 연결 실패: URL 및 Apps Script 권한을 확인해 주세요.', 'error');
      }
    }
  }

  function toggleMenu(forceOpen) {
    var menu = el('headerSyncMenu');
    var btn = el('headerSyncBtn');
    if (!menu) return;
    var isOpen = !menu.hidden;
    var target = (typeof forceOpen === 'boolean') ? forceOpen : !isOpen;
    menu.hidden = !target;
    if (btn) btn.setAttribute('aria-expanded', String(target));
    if (target) {
      updateUI();
    }
  }

  function logout() {
    if (confirm('Google 계정 연동을 해제하시겠습니까?\n(로컬에 저장된 자리배치와 명렬표는 안전하게 보존됩니다.)')) {
      setAuth(null);
      updateUI();
      showToast('로그아웃되었습니다. 로컬 저장 모드로 전환됩니다.', 'info');
    }
  }

  function init() {
    var btn = el('headerSyncBtn');
    if (btn) {
      btn.addEventListener('click', function (e) {
        e.stopPropagation();
        toggleMenu();
      });
    }

    var menu = el('headerSyncMenu');
    if (menu) {
      menu.addEventListener('click', function (e) {
        if (e.target === menu) {
          toggleMenu(false);
        } else {
          e.stopPropagation();
        }
      });
    }

    var closeModalBtn = el('syncModalCloseBtn');
    if (closeModalBtn) {
      closeModalBtn.addEventListener('click', function () {
        toggleMenu(false);
      });
    }

    var guideTabs = document.querySelectorAll('.sync-guide-tab');
    if (guideTabs && guideTabs.length) {
      guideTabs.forEach(function (tab) {
        tab.addEventListener('click', function () {
          guideTabs.forEach(function (t) { t.classList.remove('is-active'); });
          tab.classList.add('is-active');
          var imgEl = el('syncGuideViewerImg');
          var capEl = el('syncGuideViewerCap');
          if (imgEl && tab.dataset.img) imgEl.src = tab.dataset.img;
          if (capEl && tab.dataset.cap) capEl.textContent = tab.dataset.cap;
        });
      });
    }

    document.addEventListener('click', function () {
      var m = el('headerSyncMenu');
      if (m && !m.hidden) toggleMenu(false);
    });

    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') {
        var m = el('headerSyncMenu');
        if (m && !m.hidden) toggleMenu(false);
      }
    });

    var saveUrlBtn = el('tongSheetUrlSaveBtn');
    if (saveUrlBtn) {
      saveUrlBtn.addEventListener('click', function () {
        var input = el('tongSheetUrlInput');
        var url = input ? input.value.trim() : '';
        if (url && !/^https:\/\/script\.google\.com\/macros\/s\/[a-zA-Z0-9_\-]+\/exec/.test(url)) {
          alert('구글 Apps Script 웹앱 배포 URL 형식이 아닙니다.\nhttps://script.google.com/macros/s/.../exec 주소를 입력해 주세요.');
          return;
        }
        setSheetUrl(url);
        showToast('구글 시트 웹앱 URL이 저장되었습니다.', 'success');
        updateUI();
        if (url) {
          syncNow(true);
        }
      });
    }

    var manualBtn = el('syncManualNowBtn');
    if (manualBtn) {
      manualBtn.addEventListener('click', function () {
        syncNow(true);
        if (typeof GoogleCalendarSync !== 'undefined') {
          if (GoogleCalendarSync.validToken && GoogleCalendarSync.validToken()) {
            GoogleCalendarSync.syncNow(false);
          } else if (GoogleCalendarSync.connect) {
            GoogleCalendarSync.connect(true);
          }
        }
      });
    }

    var calConnectNowBtn = el('syncCalConnectNowBtn');
    if (calConnectNowBtn) {
      calConnectNowBtn.addEventListener('click', function () {
        if (typeof GoogleCalendarSync !== 'undefined') {
          if (GoogleCalendarSync.validToken && GoogleCalendarSync.validToken()) {
            GoogleCalendarSync.syncNow(false);
          } else {
            GoogleCalendarSync.connect(false);
          }
        }
      });
    }

    var switchCalBtn = el('syncSwitchCalAccountBtn');
    if (switchCalBtn) {
      switchCalBtn.addEventListener('click', function () {
        if (typeof GoogleCalendarSync !== 'undefined' && GoogleCalendarSync.switchCalendarAccountOnly) {
          GoogleCalendarSync.switchCalendarAccountOnly();
        }
      });
    }

    var useMainCalBtn = el('syncUseMainCalAccountBtn');
    if (useMainCalBtn) {
      useMainCalBtn.addEventListener('click', function () {
        if (typeof GoogleCalendarSync !== 'undefined' && GoogleCalendarSync.useMainAccountForCalendar) {
          GoogleCalendarSync.useMainAccountForCalendar();
        }
      });
    }

    var pullBtn = el('syncPullFromSheetBtn');
    if (pullBtn) {
      pullBtn.addEventListener('click', function () {
        pullFromSheet(true);
      });
    }

    var logoutBtn = el('syncLogoutBtn');
    if (logoutBtn) {
      logoutBtn.addEventListener('click', logout);
    }

    var topSyncBtn = el('topSyncBtn');
    if (topSyncBtn) {
      topSyncBtn.addEventListener('click', function (e) {
        e.stopPropagation();
        toggleMenu(true);
      });
    }
    var dashOpenSyncBtn = el('dashOpenGoogleSyncBtn');
    if (dashOpenSyncBtn) {
      dashOpenSyncBtn.addEventListener('click', function (e) {
        e.stopPropagation();
        toggleMenu(true);
      });
    }
    var dashPullBtn = el('dashPullFromSheetBtn');
    if (dashPullBtn) {
      dashPullBtn.addEventListener('click', function () { pullFromSheet(true); });
    }
    var dashPushBtn = el('dashPushToSheetBtn');
    if (dashPushBtn) {
      dashPushBtn.addEventListener('click', function () { syncNow(true); });
    }
    async function connectSheetUrlSafely(url) {
      if (url && !/^https:\/\/script\.google\.com\/macros\/s\/[a-zA-Z0-9_\-]+\/exec/.test(url)) {
        alert('구글 Apps Script 웹앱 배포 URL 형식이 아닙니다.\nhttps://script.google.com/macros/s/.../exec 주소를 입력해 주세요.');
        return;
      }
      setSheetUrl(url);
      updateUI();
      if (!url) {
        showToast('시트 URL 연결을 해제했습니다 (로컬 데이터는 안전하게 유지됩니다).', 'info');
        return;
      }
      showToast('구글 시트에 기존 데이터가 있는지 확인하는 중입니다...', 'info');
      try {
        var res = await fetch(url + (url.indexOf('?') >= 0 ? '&' : '?') + 'action=pull_tong_notebook&_t=' + Date.now());
        var data = await res.json();
        var hasSheetData = Boolean(data && data.success && (data.teacherProfile || (Array.isArray(data.students) && data.students.length > 0)));
        if (hasSheetData) {
          var shouldPull = confirm(
            '연결한 구글 시트에 이미 저장된 시간표·명렬 데이터가 있습니다.\n\n' +
            '• [확인]: 시트 데이터를 이 기기로 불러옵니다 (Pull · 권장)\n' +
            '• [취소]: 현재 이 기기의 로컬 데이터를 시트로 올립니다 (Push)'
          );
          if (shouldPull) {
            await pullFromSheet(true);
          } else {
            await syncNow(true);
          }
        } else {
          await syncNow(true);
        }
      } catch (err) {
        await syncNow(true);
      }
    }

    var dashQuickSaveBtn = el('dashQuickSheetSaveBtn');
    if (dashQuickSaveBtn) {
      dashQuickSaveBtn.addEventListener('click', function () {
        var qInput = el('dashQuickSheetUrlInput');
        connectSheetUrlSafely(qInput ? qInput.value.trim() : '');
      });
    }

    updateUI();
    initGIS();

    var existingAuth = getAuth();
    if (existingAuth && existingAuth.email && typeof GoogleCalendarSync !== 'undefined') {
      setTimeout(function () {
        if (GoogleCalendarSync.syncWithAccount) {
          GoogleCalendarSync.syncWithAccount(existingAuth.email, existingAuth.sub, true);
        }
      }, 900);
    }
  }

  // ------------------------------------------------------------------
  // 가져오기(Pull): 시트에 쌓인 기록을 이 기기에 합칩니다.
  //  1) 먼저 이 기기의 변경을 보냅니다. (안 보내고 가져오면 시트의 옛 값이 이 기기의 새 값을 덮어쓸 수 있습니다)
  //  2) 관찰 기록은 없는 것만 더합니다.
  //  3) 출결·칭찬·1인1역·학생메모는 시트에서 가장 나중에 쌓인 줄이 이 기기가 아직 모르는 줄일 때만 반영합니다.
  //     이 시트와 처음 맞춰 보는 항목은 이 기기에 값이 없을 때만 채웁니다. (이 기기의 값을 덮어쓰지 않음)
  // ------------------------------------------------------------------
  function findSyncClass(classes, className, studentId) {
    var sid = String(studentId || '');
    var byName = null;
    var withStudent = [];
    classes.forEach(function (cls) {
      if (className && cls.name === className) byName = cls;
      if ((cls.data.students || []).some(function (st) { return String(st.id) === sid; })) withStudent.push(cls);
    });
    if (byName && withStudent.indexOf(byName) >= 0) return byName;
    if (!withStudent.length) return null;
    var current = withStudent.filter(function (cls) { return cls.index === currentClassIndex; })[0];
    return current || withStudent[0];
  }

  // [8_출결_칭찬_기록] 탭이 예전 6열 양식(학급 열 없음)이면 값이 한 칸씩 밀려 옵니다.
  function normalizeSheetRow(row) {
    var kinds = ['출결', '칭찬', '학급역할', '학생관찰'];
    var out = {
      timestamp: String(row.timestamp || ''), studentId: String(row.studentId || '').trim(), className: String(row.className || '').trim(),
      category: String(row.category || '').trim(), detail: String(row.detail || '').trim(), memo: String(row.memo || '').trim()
    };
    if (kinds.indexOf(out.category) < 0 && kinds.indexOf(out.className) >= 0) {
      out = { timestamp: out.timestamp, studentId: out.studentId, className: '', category: out.className, detail: out.category, memo: out.detail };
    }
    return out;
  }

  function mergeSheetNotes(observations, classes, classNames, state, summary) {
    // 예전 Pull 이 만든 사본(내용 뒤에 과목·반이 붙은 중복)을 정리하고, 반별로 이미 있는 기록의 키를 모읍니다.
    classes.forEach(function (cls) {
      var kept = [];
      cls.noteKeys = {};
      (Array.isArray(cls.data.specialNotes) ? cls.data.specialNotes : []).forEach(function (sn) {
        var key = noteKey(sn.studentId, sn.date, noteCompareText(sn.content, classNames));
        if (cls.noteKeys[key]) { summary.cleaned++; return; }
        cls.noteKeys[key] = true;
        var clean = stripSubjectSuffix(sn.content);
        if (clean && clean !== sn.content) sn.content = clean;
        kept.push(sn);
      });
      cls.data.specialNotes = kept;
    });

    (Array.isArray(observations) ? observations : []).forEach(function (ob) {
      var studentId = String(ob.studentId || '').trim();
      var date = String(ob.date || '').slice(0, 10);
      var raw = String(ob.content || '').trim();
      if (!studentId || !raw) return;
      var suffix = raw.match(SUBJECT_SUFFIX_RE);
      var sheetClass = suffix ? String(suffix[1] || '').trim() : '';
      var content = noteCompareText(raw, classNames);
      var category = String(ob.category || '수업 참여').trim();
      // 시트가 모르는 유형은 '수업관찰' + "[유형] 내용" 으로 보냈으므로 되돌려 놓습니다.
      var tag = category === '수업관찰' ? content.match(/^\[([^\]]+)\]\s*/) : null;
      if (tag && NOTE_CATEGORIES.indexOf(tag[1]) >= 0) {
        category = tag[1];
        content = content.slice(tag[0].length).trim();
      }
      if (!content) return;
      var target = findSyncClass(classes, sheetClass, studentId);
      if (!target) { summary.unmatched++; return; }
      var key = noteKey(studentId, date, content);
      var hash = hashKey('O|' + key);
      if (target.noteKeys[key]) { state.obs[hash] = 1; return; }
      if (state.obs[hash]) return;  // 이 기기에서 받은 뒤 지운 기록 → 되살리지 않습니다.
      var nowIso = new Date().toISOString();
      target.data.specialNotes.push({
        id: generateId(), studentId: studentId, studentName: String(ob.studentName || '').trim(), date: date,
        subject: String(ob.subject || '').trim() || '수업관찰', category: category, content: content,
        createdAt: nowIso, updatedAt: nowIso
      });
      target.noteKeys[key] = true;
      state.obs[hash] = 1;
      summary.notes++;
    });

    // 보내기가 끝난 뒤이므로 이 기기의 기록도 모두 시트에 있습니다.
    classes.forEach(function (cls) {
      Object.keys(cls.noteKeys).forEach(function (key) { state.obs[hashKey('O|' + key)] = 1; });
    });
  }

  // 시트의 한 줄을 이 기기에 반영합니다. fillOnly 이면 이 기기에 값이 없을 때만 채웁니다. 바뀌었으면 true.
  function applySheetRow(group, detail, memo, fillOnly) {
    var d = group.cls.data;
    var sid = group.studentId;
    if (group.kind === 'A') {
      var date = group.date;
      var status = detail.split(' (')[0].trim();
      var current = (d.attendance && d.attendance[date] && d.attendance[date][sid]) || '';
      if (current === '출석') current = '';
      if (status === '출석') {
        if (fillOnly || !current) return false;
        delete d.attendance[date][sid];
        if (d.attendReasons && d.attendReasons[date]) delete d.attendReasons[date][sid];
        if (d.attendMemos && d.attendMemos[date]) delete d.attendMemos[date][sid];
        return true;
      }
      if (SHEET_ATT_STATES.indexOf(status) < 0) return false;
      if (fillOnly && current) return false;
      var reasonMatch = detail.match(/\(([^()]*)\)/);
      var reason = reasonMatch && SHEET_ATT_REASONS.indexOf(reasonMatch[1].trim()) >= 0 ? reasonMatch[1].trim() : '';
      var currentReason = (d.attendReasons && d.attendReasons[date] && d.attendReasons[date][sid]) || '';
      if (current === status && currentReason === reason) return false;
      if (!d.attendance) d.attendance = {};
      if (!d.attendance[date]) d.attendance[date] = {};
      d.attendance[date][sid] = status;
      if (!d.attendReasons) d.attendReasons = {};
      if (!d.attendReasons[date]) d.attendReasons[date] = {};
      if (reason) d.attendReasons[date][sid] = reason; else delete d.attendReasons[date][sid];
      if (memo && !/출결 (기록|정정)$/.test(memo)) {
        if (!d.attendMemos) d.attendMemos = {};
        if (!d.attendMemos[date]) d.attendMemos[date] = {};
        d.attendMemos[date][sid] = memo;
      }
      return true;
    }
    if (group.kind === 'P') {
      var scoreMatch = detail.match(/^\+(\d+)점/);
      if (!scoreMatch) return false;
      var score = Number(scoreMatch[1]);
      if (!d.praise) d.praise = {};
      var currentScore = Number(d.praise[sid]) || 0;
      if (currentScore === score || (fillOnly && currentScore > 0)) return false;
      if (score > 0) d.praise[sid] = score; else delete d.praise[sid];
      return true;
    }
    if (group.kind === 'D') {
      var role = detail.replace(/^1인1역:\s*/, '').trim();
      if (role === '(해제)') role = '';
      if (!d.duties) d.duties = {};
      var currentRole = String(d.duties[sid] || '').trim();
      if (currentRole === role || (fillOnly && currentRole)) return false;
      if (role) d.duties[sid] = role; else delete d.duties[sid];
      return true;
    }
    if (group.kind === 'M') {
      var text = detail === '(메모 삭제)' ? '' : detail;
      var memos = memosOf(group.cls);
      var currentText = String(memos[sid] || '').trim();
      if (currentText === text || (fillOnly && currentText)) return false;
      if (text) memos[sid] = text; else delete memos[sid];
      return true;
    }
    return false;
  }

  function mergeSheetRows(rows, classes, state, summary) {
    var kindOf = { '출결': 'A', '칭찬': 'P', '학급역할': 'D', '학생관찰': 'M' };
    var counter = { A: 'attendance', P: 'praise', D: 'duties', M: 'memos' };
    var groups = {};
    (Array.isArray(rows) ? rows : []).forEach(function (raw) {
      var row = normalizeSheetRow(raw);
      var kind = kindOf[row.category];
      if (!kind || !row.studentId || !row.detail) return;
      var date = row.timestamp.slice(0, 10);
      var cls = findSyncClass(classes, row.className, row.studentId);
      if (!cls) { summary.unmatched++; return; }
      var key = rowsKey(kind, cls.name, row.studentId, kind === 'A' ? date : '');
      if (!groups[key]) groups[key] = { kind: kind, cls: cls, studentId: row.studentId, date: date, entries: [], memo: '' };
      groups[key].entries.push(date + '|' + row.detail);
      groups[key].memo = row.memo;
    });

    Object.keys(groups).forEach(function (key) {
      var group = groups[key];
      var known = state.rows[key] || [];
      var newest = group.entries[group.entries.length - 1];
      // 가장 나중 줄을 이 기기가 아직 모를 때만 반영합니다. (이미 아는 줄이 마지막이면 이 기기가 최신)
      if (known.indexOf(newest) < 0 && applySheetRow(group, entryDetail(newest), group.memo, !known.length)) {
        summary[counter[group.kind]]++;
      }
      state.rows[key] = group.entries.slice();
    });
    // 시트에서 사라진 줄은 잊습니다. (다음에 보낼 때 이 기기의 현재 값을 다시 올립니다)
    Object.keys(state.rows).forEach(function (key) {
      if (!groups[key]) delete state.rows[key];
    });
  }

  function restoreAfterPull(snapshot) {
    try {
      allClassData = JSON.parse(snapshot.classes);
      localStorage.setItem(SYNC_STATE_KEY, snapshot.state);
    } catch (e) { return; }
    loadCurrentClassData();
    saveToLocalStorage();
    renderSeating();
    refreshClassExtras();
    showToast('가져오기 전 상태로 되돌렸습니다.');
  }

  async function pullFromSheet(manual) {
    var sheetUrl = getSheetUrl();
    if (!sheetUrl) {
      if (manual) {
        showToast('먼저 구글 시트 웹앱 URL을 입력하고 저장해 주세요.', 'error');
        toggleMenu(true);
      }
      return;
    }
    isSyncing = true;
    updateUI();
    try {
      await pushToSheet(sheetUrl);

      var res = await fetch(sheetUrl + (sheetUrl.indexOf('?') >= 0 ? '&' : '?') + 'action=pull_tong_notebook&_t=' + Date.now());
      var data = await res.json();
      if (data && data.success) {
        var d = allClassData[currentClassIndex];
        var updatedCount = 0;
        if (data.teacherProfile && applyTeacherProfile(data.teacherProfile)) {
          updatedCount++;
          if (typeof SheetMemoPad !== 'undefined' && SheetMemoPad.reload) SheetMemoPad.reload();
          if (typeof TeacherSchedule !== 'undefined' && TeacherSchedule.reloadFromStorage) TeacherSchedule.reloadFromStorage();
          if (typeof TodayBriefing !== 'undefined' && TodayBriefing.refreshAll) TodayBriefing.refreshAll();
        }
        if (d && Array.isArray(data.students) && data.students.length > 0 && (!students || students.length === 0)) {
          students = data.students.map(function (s) {
            return { id: String(s.id), name: String(s.name), gender: s.gender || '' };
          });
          d.students = students.slice();
          updatedCount += students.length;
        }
        if (d && Array.isArray(data.observations) && data.observations.length > 0) {
          if (!Array.isArray(d.specialNotes)) d.specialNotes = [];
          var existingObs = {};
          d.specialNotes.forEach(function (sn) {
            existingObs[(sn.date || '') + '_' + sn.studentId + '_' + sn.content] = true;
          });
          data.observations.forEach(function (ob) {
            var k = (ob.date || '') + '_' + ob.studentId + '_' + ob.content;
            if (ob.content && !existingObs[k]) {
              existingObs[k] = true;
              d.specialNotes.push({
                id: generateId(),
                studentId: ob.studentId,
                studentName: ob.studentName,
                date: ob.date,
                subject: ob.subject || '수업관찰',
                category: ob.category || '수업 참여',
                content: ob.content,
                createdAt: new Date().toISOString(),
                updatedAt: new Date().toISOString()
              });
              updatedCount++;
            }
          });
        }
        saveCurrentClassData();
        renderSeating();
        if (typeof Attend !== 'undefined' && Attend.render) Attend.render();
        isSyncing = false;
        localStorage.setItem(LAST_SYNC_KEY, String(Date.now()));
        updateUI();
        if (manual) {
          showToast('구글 시트에서 최신 명렬·관찰 데이터를 성공적으로 불러왔습니다 (' + updatedCount + '건 반영).', 'success');
        }
      } else {
        throw new Error('데이터 조회 실패');
      }
    } catch (err) {
      console.warn('Pull error:', err);
      isSyncing = false;
      updateUI();
      if (manual) {
        showToast('시트에서 가져오지 못했습니다. 웹앱 URL과 배포 상태를 확인해 주세요. (이 기기의 자료는 바뀌지 않았습니다)', 'error');
      }
    }
  }

  function collectTeacherProfile() {
    function rj(k, d) {
      try { var v = JSON.parse(localStorage.getItem(k) || 'null'); return v == null ? d : v; } catch (e) { return d; }
    }
    return {
      version: 2,
      updatedAt: Date.now(),
      school: rj('tongNeisSchool_v1', null),
      isNonHomeroom: Boolean(rj('tongIsNonHomeroom_v1', false)),
      homeroom: rj('tongHomeroomClass_v1', null),
      subjectClasses: rj('tongSubjectClasses_v1', []),
      timetable: rj('tongTeacherTimetable_v1', null),
      customClubs: rj('tongCustomClubs_v1', []),
      sheetMemos: rj('tongSheetMemoPad_v1', [])
    };
  }

  function applyTeacherProfile(tp) {
    if (!tp || typeof tp !== 'object') return false;
    try {
      if (tp.school) localStorage.setItem('tongNeisSchool_v1', JSON.stringify(tp.school));
      if (typeof tp.isNonHomeroom === 'boolean') localStorage.setItem('tongIsNonHomeroom_v1', JSON.stringify(tp.isNonHomeroom));
      if (tp.homeroom !== undefined) localStorage.setItem('tongHomeroomClass_v1', JSON.stringify(tp.homeroom));
      if (Array.isArray(tp.subjectClasses)) localStorage.setItem('tongSubjectClasses_v1', JSON.stringify(tp.subjectClasses));
      if (tp.timetable && typeof tp.timetable === 'object') localStorage.setItem('tongTeacherTimetable_v1', JSON.stringify(tp.timetable));
      if (Array.isArray(tp.customClubs)) localStorage.setItem('tongCustomClubs_v1', JSON.stringify(tp.customClubs));
      if (Array.isArray(tp.sheetMemos) && tp.sheetMemos.length) localStorage.setItem('tongSheetMemoPad_v1', JSON.stringify(tp.sheetMemos));
      return true;
    } catch (e) {
      return false;
    }
  }

  function queueProfileSync() {
    if (!getSheetUrl()) return;
    if (autoSyncTimer) clearTimeout(autoSyncTimer);
    autoSyncTimer = setTimeout(function () {
      syncNow(false);
    }, 2500);
  }

  function updateDashSyncHubUI() {
    var url = getSheetUrl();
    var auth = getAuth();
    var badge = el('dashSyncStatusBadge');
    var desc = el('dashSyncStatusDesc');
    var quickInput = el('dashQuickSheetUrlInput');
    if (quickInput && document.activeElement !== quickInput) {
      quickInput.value = url || '';
    }
    if (badge) {
      if (url) {
        badge.className = 'dash-sync-badge connected';
        badge.textContent = auth ? ('연결됨 · ' + auth.email) : '시트 연결됨';
      } else {
        badge.className = 'dash-sync-badge offline';
        badge.textContent = '로컬 모드';
      }
    }
    if (desc) {
      if (url) {
        var lastTs = Number(localStorage.getItem(LAST_SYNC_KEY) || 0);
        var timeTxt = lastTs ? new Date(lastTs).toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' }) : '방금 연결';
        desc.innerHTML = '내 구글 마스터 시트와 연결되어 있습니다 (마지막 동기화: <b>' + timeTxt + '</b>). 다른 기기에서 이어서 쓸 때는 <b>[불러오기 (Pull)]</b>를 누르세요.';
      } else {
        desc.innerHTML = '어느 기기에서든 내 시간표·명렬표·관찰기록을 똑같이 보려면 <b>먼저 내 구글 시트 웹앱 URL을 연결</b>하세요. 연결 즉시 시트에서 기존 시간표를 불러오거나 현재 설정을 저장합니다.';
      }
    }
  }

  return {
    init: init,
    getAuth: getAuth,
    setUnifiedAuthFromUser: setUnifiedAuthFromUser,
    getSheetUrl: getSheetUrl,
    setSheetUrl: setSheetUrl,
    queueRecord: queueRecord,
    queueProfileSync: queueProfileSync,
    collectTeacherProfile: collectTeacherProfile,
    syncNow: syncNow,
    pullFromSheet: pullFromSheet,
    updateUI: updateUI,
    toggleMenu: toggleMenu
  };
})();

window.CloudSheetSync = CloudSheetSync;
window.TongSyncBridge = {
  pushToGoogleSheet: function () { return CloudSheetSync.syncNow(true); },
  pullFromGoogleSheet: function () { return CloudSheetSync.pullFromSheet(true); }
};

// ===== 초기화 =====
document.addEventListener('DOMContentLoaded', function() {
  initAllClassData();
  loadFromLocalStorage();
  initClassDropdown();
  initEventListeners();
  initDateInputs();
  initGuide();
  Timer.init();
  Picker.init();
  Notice.init();
  Praise.init();
  Attend.init();
  Duty.init();
  SpecialNotes.init();
  Ladder.init();
  Records.init();
  CustomFavorites.init();
  ThemeManager.init();
  HelpMenu.init();
  SchoolToday.init();
  UsageStats.init();
  ScheduleCalendar.init();
  GoogleCalendarSync.init();
  CloudSheetSync.init();
  AiAssistant.init();

  initSettingsTabs();
  initModeTabs();
  initWizardSteps();
  initHeroTags();
  initSideRail();
  initSeatActions();

  var dockGuide = document.getElementById('dockGuideBtn');
  if (dockGuide) dockGuide.addEventListener('click', openGuide);
  var dockSettings = document.getElementById('dockSettingsBtn');
  if (dockSettings) dockSettings.addEventListener('click', function () { openModal('settingsModal'); });

  // 소리 설정 버튼 초기 표시
  var soundBtn = document.getElementById('soundToggleBtn');
  if (soundBtn && !Sound.isEnabled()) {
    soundBtn.textContent = '🔇 소리 꺼짐';
    soundBtn.classList.add('muted');
  }

  // 앱을 열면 좌석표가 아니라 독립된 교무수첩 홈을 먼저 보여 줍니다.
  if (students.length > 0) renderSeating();
  showHomeScreen();

  maybeAutoOpenGuide();
});

function initAllClassData() {
  for (var i = 0; i < MAX_CLASSES; i++) {
    if (!allClassData[i]) {
      allClassData[i] = {
        className: (i + 1) + '반',
        students: [],
        currentSeats: {},
        lockedSeats: {},
        config: { totalDesks: 30, columns: 5, seatType: 'pair', layoutIndex: 0, rows: 6 },
        relations: { close: {}, far: {} },
        studentMemos: {},
        specialNotes: [],
        savedRecords: []
      };
    }
    if (!Array.isArray(allClassData[i].specialNotes)) allClassData[i].specialNotes = [];
  }
}

function initClassDropdown() {
  function fill(select) {
    if (!select) return;
    select.innerHTML = '';
    for (var i = 0; i < MAX_CLASSES; i++) {
      var option = document.createElement('option');
      option.value = i;
      var data = allClassData[i];
      option.textContent = data && data.className ? data.className : ((i + 1) + '반');
      if (data && data.students && data.students.length > 0) option.textContent += ' (' + data.students.length + '명)';
      select.appendChild(option);
    }
    select.value = currentClassIndex;
  }
  fill(document.getElementById('currentClass'));
  fill(document.getElementById('avClassSelect'));
}
function updatePeriodSummary() {
  var node = document.getElementById('spPeriodSummary');
  if (!node) return;
  function shortDate(key) {
    var p = String(key || '').split('-');
    return p.length === 3 ? Number(p[1]) + '/' + Number(p[2]) : '-';
  }
  node.textContent = '(' + shortDate(settings.periodStart) + '–' + shortDate(settings.periodEnd) + ')';
}

function initDateInputs() {
  var today = new Date();
  var nextMonth = new Date(today);
  nextMonth.setMonth(nextMonth.getMonth() + 1);
  settings.periodStart = settings.periodStart || formatDateForInput(today);
  settings.periodEnd = settings.periodEnd || formatDateForInput(nextMonth);
  var start = document.getElementById('periodStart');
  var end = document.getElementById('periodEnd');
  if (start) start.value = settings.periodStart;
  if (end) end.value = settings.periodEnd;
  updatePeriodSummary();
}
function initColumnOptions() {
  var container = document.getElementById('columnOptions');
  if (!container) {
    console.error('columnOptions container not found');
    return;
  }
  container.innerHTML = '';
  // 모둠 배치는 가로로 2~4모둠까지만 (그 이상은 교실에 들어가지 않습니다)
  var columnCounts = config.seatType === 'group' ? [2, 3, 4] : [3, 4, 5, 6, 7];
  if (columnCounts.indexOf(config.columns) === -1) {
    config.columns = config.seatType === 'group' ? 3 : 5;
  }
  for (var i = 0; i < columnCounts.length; i++) {
    var cols = columnCounts[i];
    var btn = document.createElement('div');
    btn.className = 'option-btn';
    if (cols === config.columns) {
      btn.className += ' active';
    }
    btn.textContent = cols + (config.seatType === 'group' ? '모둠' : '열');
    btn.dataset.cols = cols;
    (function(c) {
      btn.addEventListener('click', function() {
        selectColumns(c);
      });
    })(cols);
    container.appendChild(btn);
  }
}

// ===== 이벤트 리스너 =====
function initEventListeners() {
  // 헤더
  document.getElementById('currentClass').addEventListener('change', switchClass);
  var avClassSelect = document.getElementById('avClassSelect');
  if (avClassSelect) avClassSelect.addEventListener('change', function () {
    document.getElementById('currentClass').value = this.value;
    switchClass();
  });
  var headerRosterBtn = document.getElementById('headerRosterBtn');
  if (headerRosterBtn) headerRosterBtn.addEventListener('click', openRosterEditor);
  var avRosterBtn = document.getElementById('avRosterBtn');
  if (avRosterBtn) avRosterBtn.addEventListener('click', openRosterEditor);
  document.getElementById('periodStart').addEventListener('change', function() {
    settings.periodStart = this.value;
    saveToLocalStorage();
    updatePeriodSummary();
  });
  document.getElementById('periodEnd').addEventListener('change', function() {
    settings.periodEnd = this.value;
    saveToLocalStorage();
    updatePeriodSummary();
  });
  
  // 액션 버튼 — 추첨은 카운트다운 스플래시를 먼저 보여줍니다
  document.getElementById('shuffleBtn').addEventListener('click', function () {
    if (isLotteryRunning) return;
    if (!students.length) { showToast('명렬표를 먼저 등록해 주세요', 'error'); return; }
    Sound.init();
    runLotterySplash(startLottery);
  });
  document.getElementById('saveBtn').addEventListener('click', openSaveModal);
  document.getElementById('historyBtn').addEventListener('click', openHistoryModal);
  document.getElementById('downloadHistoryBtn').addEventListener('click', downloadHistoryExcel);
  document.getElementById('studentViewBtn').addEventListener('click', function() {
    printSeatingChart('student');
  });
  document.getElementById('teacherViewBtn').addEventListener('click', function() {
    printSeatingChart('teacher');
  });
  var screenTeacherViewBtn = document.getElementById('screenTeacherViewBtn');
  if (screenTeacherViewBtn) screenTeacherViewBtn.addEventListener('click', toggleTeacherScreenView);
  var spBackupBtn = document.getElementById('spBackupBtn');
  if (spBackupBtn) spBackupBtn.addEventListener('click', openBackupModal);
  var sameGenderSeating = document.getElementById('sameGenderSeating');
  if (sameGenderSeating) sameGenderSeating.addEventListener('change', function() {
    config.sameGenderSeating = this.checked;
    saveCurrentClassData();
    showToast(this.checked ? '👫 같은 성별 우선 배치를 켰습니다' : '같은 성별 우선 배치를 껐습니다');
  });
  document.querySelectorAll('input[name="numberOrderMode"]').forEach(function(input) {
    input.addEventListener('change', function() {
      if (!this.checked) return;
      config.numberOrderMode = this.value;
      config.numberOrderSeating = this.value !== 'random';
      syncSameGenderSeatingOption();
      saveCurrentClassData();
      renderSeatPanel();
      var message = this.value === 'row'
        ? '↔ 학번 가로순 배치로 설정했습니다'
        : (this.value === 'column' ? '↕ 학번 세로순 배치로 설정했습니다' : '🎱 무작위 추첨으로 전환했습니다');
      showToast(message);
    });
  });
  syncSameGenderSeatingOption();
  document.getElementById('settingsBtn').addEventListener('click', function() {
    openModal('settingsModal');
  });
  
  // 설정 위자드
  document.getElementById('deskMinus').addEventListener('click', function() {
    adjustDeskCount(-1);
  });
  document.getElementById('deskPlus').addEventListener('click', function() {
    adjustDeskCount(1);
  });
  document.getElementById('totalDesks').addEventListener('change', function() {
    config.totalDesks = parseInt(this.value) || 30;
    updatePreview();
    updateEmptySeatInfo();
  });
  document.getElementById('studentCount').addEventListener('change', updateEmptySeatInfo);
  document.getElementById('applyConfigBtn').addEventListener('click', applyConfig);
  
  document.querySelectorAll('.type-card').forEach(function(card) {
    card.addEventListener('click', function() {
      selectSeatType(this.dataset.type);
    });
  });
  
  // 추첨
  document.getElementById('skipLotteryBtn').addEventListener('click', skipLottery);
  document.querySelectorAll('.speed-btn').forEach(function(btn) {
    btn.addEventListener('click', function() { setLotterySpeed(this.dataset.speed); });
  });
  document.getElementById('soundToggleBtn').addEventListener('click', toggleLotterySound);

  // 창 크기가 바뀌어도 좌석표가 한 화면에 유지되도록
  window.addEventListener('resize', scheduleFit);
  if (window.visualViewport) window.visualViewport.addEventListener('resize', scheduleFit);
  var appHeader = document.querySelector('.app-header');
  if (appHeader && window.ResizeObserver) {
    new ResizeObserver(scheduleFit).observe(appHeader);
  }
  
  // 모달 닫기
  document.querySelectorAll('.close-btn, [data-modal]').forEach(function(btn) {
    btn.addEventListener('click', function() {
      var modal = this.dataset.modal || this.closest('.modal').id;
      if (modal) closeModal(modal);
    });
  });
  
  // 탭
  document.querySelectorAll('.tab-btn').forEach(function(btn) {
    btn.addEventListener('click', function() {
      var tab = this.dataset.tab;
      document.querySelectorAll('.tab-btn').forEach(function(b) {
        b.classList.remove('active');
      });
      document.querySelectorAll('.tab-content').forEach(function(c) {
        c.classList.remove('active');
      });
      this.classList.add('active');
      document.getElementById(tab + 'Tab').classList.add('active');
    });
  });
  
  // 빈자리
  document.getElementById('toggleEmptyLockBtn').addEventListener('click', toggleEmptyLock);
  
  // 학생 설정
  document.getElementById('addCloseBtn').addEventListener('click', function() {
    openStudentPicker('close');
  });
  document.getElementById('addFarBtn').addEventListener('click', function() {
    openStudentPicker('far');
  });
  document.getElementById('closePickerBtn').addEventListener('click', closeStudentPicker);
  document.getElementById('studentSearch').addEventListener('input', filterStudents);
  document.getElementById('saveRelationsBtn').addEventListener('click', saveStudentSettings);
  
  // 메모 자동저장
  renderStudentSummary(selectedStudent);
  document.querySelectorAll('#studentModal .tab-btn').forEach(function(b){ b.classList.toggle('active', b.dataset.tab === 'summary'); });
  document.querySelectorAll('#studentModal .tab-content').forEach(function(c){ c.classList.toggle('active', c.id === 'summaryTab'); });

  var memoInput = document.getElementById('studentMemoInput');
  if (memoInput) {
    memoInput.addEventListener('input', debounce(saveStudentMemo, 500));
  }
  
  // 저장
  document.getElementById('confirmSaveBtn').addEventListener('click', confirmSave);
  
  // 설정 메뉴
  document.getElementById('openRosterBtn').addEventListener('click', openRosterEditor);
  document.getElementById('openExcelUploadBtn').addEventListener('click', openExcelModal);
  document.getElementById('openSectionConfigBtn').addEventListener('click', openSectionConfig);
  document.getElementById('openBackupBtn').addEventListener('click', openBackupModal);
  document.getElementById('confirmResetBtn').addEventListener('click', confirmReset);
  
  // 명렬표
  document.getElementById('saveRosterBtn').addEventListener('click', saveRoster);
  document.getElementById('rosterClassName').addEventListener('input', updateClassName);
  // 명렬표 모달 내 엑셀 버튼
  document.getElementById('downloadSampleBtn').addEventListener('click', downloadSampleExcel);
  document.getElementById('downloadRosterBtn').addEventListener('click', function() {
    downloadRosterExcel(true);
  });
document.getElementById('downloadSampleExcelBtn').addEventListener('click', function() {
  downloadSampleExcel();
  closeModal('settingsModal');
});
  var downloadRosterExcelBtn = document.getElementById('downloadRosterExcelBtn');
  if (downloadRosterExcelBtn) {
    downloadRosterExcelBtn.addEventListener('click', function() {
      downloadRosterExcel(false);
    });
  }
  var uploadExcelBtn = document.getElementById('uploadExcelInRosterBtn');
  var rosterExcelInput = document.getElementById('rosterExcelInput');
  if (uploadExcelBtn && rosterExcelInput) {
    uploadExcelBtn.addEventListener('click', function() { rosterExcelInput.click(); });
    rosterExcelInput.addEventListener('change', function() {
      if (this.files.length > 0) {
        handleExcelFileInRoster(this.files[0]);
      }
    });
  }
  
  // 엑셀 업로드 이벤트
  var dropZone = document.getElementById('excelDropZone');
  var fileInput = document.getElementById('excelFileInput');
  
  if (dropZone && fileInput) {
    dropZone.addEventListener('click', function() {
      fileInput.click();
    });
    
    dropZone.addEventListener('dragover', function(e) {
      e.preventDefault();
      this.classList.add('dragover');
    });
    
    dropZone.addEventListener('dragleave', function() {
      this.classList.remove('dragover');
    });
    
    dropZone.addEventListener('drop', function(e) {
      e.preventDefault();
      this.classList.remove('dragover');
      if (e.dataTransfer.files.length > 0) {
        handleExcelFile(e.dataTransfer.files[0]);
      }
    });
    
    fileInput.addEventListener('change', function() {
      if (this.files.length > 0) {
        handleExcelFile(this.files[0]);
      }
    });
  }
  
  document.getElementById('confirmExcelBtn').addEventListener('click', applyExcelData);
  
  // 백업
  document.getElementById('exportBtn').addEventListener('click', exportData);
  document.getElementById('copyExportBtn').addEventListener('click', copyExportCode);
  document.getElementById('importBtn').addEventListener('click', importData);
}

// ===== 엑셀 모달 =====
function openExcelModal() {
  closeModal('settingsModal');
  parsedExcelData = null;
  var fileInput = document.getElementById('excelFileInput');
  if (fileInput) fileInput.value = '';
  document.getElementById('excelPreview').style.display = 'none';
  document.getElementById('confirmExcelBtn').disabled = true;
  openModal('excelModal');
}
// ===== 샘플 양식 다운로드 =====
function downloadSampleExcel() {
  var sampleData = [
    ['학번', '이름', '성별'],
    ['1', '홍길동', '남'],
    ['2', '김영희', '여'],
    ['3', '이철수', '남'],
    ['4', '박지영', '여'],
    ['5', '최민준', '남'],
    ['6', '정수아', '여'],
    ['7', '강동혁', '남'],
    ['8', '조은지', '여'],
    ['9', '윤재호', '남'],
    ['10', '서하은', '여']
  ];
  var ws = XLSX.utils.aoa_to_sheet(sampleData);
  ws['!cols'] = [{ wch: 8 }, { wch: 12 }, { wch: 8 }];
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '1반');
  var ws2 = XLSX.utils.aoa_to_sheet([
    ['학번', '이름', '성별'],
    ['1', '학생1', '남'],
    ['2', '학생2', '여']
  ]);
  XLSX.utils.book_append_sheet(wb, ws2, '2반');
  XLSX.writeFile(wb, '명렬표_샘플양식.xlsx');
  showToast('샘플 양식 다운로드 완료!');
}

// ===== 현재 명렬표 엑셀 다운로드 =====
function parseRosterTextForExcel(text) {
  return String(text || '').trim().split('\n').filter(function(line) {
    return line.trim();
  }).map(function(line) {
    var parts = line.trim().split(/\s+/);
    var id = parts[0] || '';
    var gender = parts[parts.length - 1];
    var isGender = gender === '남' || gender === '여';
    return {
      id: id,
      name: isGender ? parts.slice(1, -1).join(' ') || '이름없음' : parts.slice(1).join(' ') || '이름없음',
      gender: isGender ? gender : '남'
    };
  });
}

function downloadRosterExcel(useEditor) {
  if (typeof XLSX === 'undefined') {
    showToast('엑셀 기능을 불러오지 못했습니다', 'error');
    return;
  }

  var classIdx = useEditor ? editingClassIndex : currentClassIndex;
  var classData = allClassData[classIdx];
  var className = classData && classData.className ? classData.className : ((classIdx + 1) + '반');
  var studentList = sortStudentsByName(classData && Array.isArray(classData.students) ? classData.students : []);

  if (useEditor) {
    var classNameInput = document.getElementById('rosterClassName');
    var rosterInput = document.getElementById('rosterInput');
    className = classNameInput && classNameInput.value.trim() ? classNameInput.value.trim() : className;
    studentList = parseRosterTextForExcel(rosterInput ? rosterInput.value : '');
  }

  if (!studentList.length) {
    showToast('다운로드할 명렬표를 먼저 입력해 주세요', 'error');
    return;
  }

  var rows = [['학번', '이름', '성별']].concat(studentList.map(function(student) {
    return [String(student.id || ''), student.name || '', student.gender || ''];
  }));
  var ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{ wch: 8 }, { wch: 14 }, { wch: 8 }];
  ws['!autofilter'] = { ref: 'A1:C' + rows.length };

  var wb = XLSX.utils.book_new();
  var sheetName = String(className).replace(/[\\\/?*\[\]:]/g, '_').slice(0, 31) || '명렬표';
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  var safeClassName = String(className).replace(/[\\/:*?"<>|]/g, '_') || '학급';
  XLSX.writeFile(wb, safeClassName + '_명렬표_' + formatDate(new Date()) + '.xlsx');
  showToast(className + ' 명렬표 엑셀 다운로드 완료!');
}

// ===== 명렬표 모달 내 엑셀 처리 =====
function handleExcelFileInRoster(file) {
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    showToast('엑셀 파일(.xlsx, .xls)만 지원합니다', 'error');
    return;
  }
  var reader = new FileReader();
  reader.onload = function(e) {
    try {
      var data = new Uint8Array(e.target.result);
      var workbook = XLSX.read(data, { type: 'array' });
      var parsed = parseExcelWorkbook(workbook);
      if (parsed && Object.keys(parsed).length > 0) {
        // 현재 편집 중인 학급에 첫 번째 시트 데이터 적용
        var firstClassName = Object.keys(parsed)[0];
        var studentList = parsed[firstClassName];
        document.getElementById('rosterClassName').value = firstClassName;
        document.getElementById('rosterInput').value = studentList.map(function(s) {
          return s.id + ' ' + s.name + ' ' + s.gender;
        }).join('\n');
        showToast(studentList.length + '명 불러오기 완료!');
        // 여러 학급이면 전체 적용 안내
        if (Object.keys(parsed).length > 1) {
          if (confirm(Object.keys(parsed).length + '개 학급이 발견되었습니다. 전체 학급을 일괄 적용할까요?')) {
            applyAllExcelClasses(parsed);
          }
        }
      } else {
        showToast('학생 데이터를 찾을 수 없습니다', 'error');
      }
    } catch (err) {
      console.error('Excel parse error:', err);
      showToast('파일 분석 실패', 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}
function applyAllExcelClasses(parsed) {
  var classNames = Object.keys(parsed);
  classNames.forEach(function(className, idx) {
    if (idx >= MAX_CLASSES) return;
    if (!allClassData[idx]) {
      allClassData[idx] = {
        className: className,
        students: [],
        currentSeats: {},
        lockedSeats: {},
        config: JSON.parse(JSON.stringify(config)),
        relations: { close: {}, far: {} },
        studentMemos: {},
        specialNotes: [],
        savedRecords: []
      };
    }
    allClassData[idx].className = className;
    allClassData[idx].students = parsed[className];
    allClassData[idx].currentSeats = {};
    allClassData[idx].lockedSeats = {};
  });
  saveToLocalStorage();
  loadCurrentClassData();
  initClassDropdown();
  renderRosterClassTabs();
  loadRosterForClass(editingClassIndex);
  showToast(classNames.length + '개 학급 적용 완료!');
}

// ===== 엑셀 파일 처리 =====
function handleExcelFile(file) {
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    showToast('엑셀 파일(.xlsx, .xls)만 지원합니다', 'error');
    return;
  }
  
  var reader = new FileReader();
  reader.onload = function(e) {
    try {
      var data = new Uint8Array(e.target.result);
      var workbook = XLSX.read(data, { type: 'array' });
      parsedExcelData = parseExcelWorkbook(workbook);
      
      if (parsedExcelData && Object.keys(parsedExcelData).length > 0) {
        showExcelPreview(parsedExcelData);
        document.getElementById('confirmExcelBtn').disabled = false;
        showToast('파일 분석 완료!');
      } else {
        showToast('학생 데이터를 찾을 수 없습니다', 'error');
      }
    } catch (err) {
      console.error('Excel parse error:', err);
      showToast('파일 분석 실패: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseExcelWorkbook(workbook) {
  var result = {};
  var sheetNames = workbook.SheetNames;
  
  if (sheetNames.length > 1) {
    // 멀티시트
    sheetNames.forEach(function(sheetName, idx) {
      if (idx >= MAX_CLASSES) return;
      var sheet = workbook.Sheets[sheetName];
      var rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      var parsed = parseSheetRows(rows);
      if (parsed.length > 0) {
        result[sheetName] = parsed;
      }
    });
  } else {
    // 단일시트
    var sheet = workbook.Sheets[sheetNames[0]];
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    if (isPivotFormat(rows)) {
      result = parsePivotFormat(rows);
    } else if (hasClassColumn(rows)) {
      result = parseWithClassColumn(rows);
    } else {
      var parsed = parseSheetRows(rows);
      if (parsed.length > 0) {
        result['1반'] = parsed;
      }
    }
  }
  
  Object.keys(result).forEach(function(className) {
    result[className] = sortStudentsByName(result[className]);
  });
  return result;
}

function parseSheetRows(rows) {
  var result = [];
  var header = rows[0] || [];
  var idCol = findCol(header, ['번호', '학번', 'id', 'no']);
  var nameCol = findCol(header, ['이름', '성명', 'name']);
  var genderCol = findCol(header, ['성별', 'gender']);
  
  if (idCol === -1 && nameCol === -1) {
    idCol = 0; nameCol = 1; genderCol = 2;
  }
  
  var startRow = (idCol >= 0 || nameCol >= 0) ? 1 : 0;
  
  for (var i = startRow; i < rows.length; i++) {
    var row = rows[i];
    if (!row || row.length === 0) continue;
    
    var id = String(row[idCol] || (result.length + 1)).trim();
    var name = String(row[nameCol] || '').trim();
    var gender = genderCol >= 0 ? String(row[genderCol] || '').trim() : '';
    
    if (!name || name === 'undefined') continue;
    
    if (gender.includes('여') || gender.toLowerCase() === 'f' || gender === '2') {
      gender = '여';
    } else {
      gender = '남';
    }
    
    result.push({ id: id, name: name, gender: gender });
  }
  
  return result;
}

function findCol(header, keywords) {
  for (var i = 0; i < header.length; i++) {
    var cell = String(header[i] || '').toLowerCase();
    for (var j = 0; j < keywords.length; j++) {
      if (cell.includes(keywords[j].toLowerCase())) return i;
    }
  }
  return -1;
}

function isPivotFormat(rows) {
  if (rows.length < 2) return false;
  var header = rows[0] || [];
  var first = String(header[0] || '').toLowerCase();
  if (first.includes('번') || first.includes('id') || first === '') {
    var cnt = 0;
    for (var i = 1; i < header.length; i++) {
      var cell = String(header[i] || '');
      if (cell.includes('반') || cell.includes('class') || /^\d+$/.test(cell)) cnt++;
    }
    return cnt >= 2;
  }
  return false;
}

function parsePivotFormat(rows) {
  var result = {};
  var header = rows[0] || [];
  
  for (var col = 1; col < header.length; col++) {
    var className = String(header[col] || (col + '반')).trim();
    var list = [];
    
    for (var row = 1; row < rows.length; row++) {
      var id = String(rows[row][0] || row).trim();
      var name = String(rows[row][col] || '').trim();
      if (name && name !== 'undefined') {
        list.push({ id: id, name: name, gender: '남' });
      }
    }
    
    if (list.length > 0) {
      result[className] = list;
    }
  }
  
  return result;
}

function hasClassColumn(rows) {
  var header = rows[0] || [];
  return findCol(header, ['학급', '반', 'class']) >= 0;
}

function parseWithClassColumn(rows) {
  var result = {};
  var header = rows[0] || [];
  var classCol = findCol(header, ['학급', '반', 'class']);
  var idCol = findCol(header, ['번호', '학번', 'id', 'no']);
  var nameCol = findCol(header, ['이름', '성명', 'name']);
  var genderCol = findCol(header, ['성별', 'gender']);
  
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row || row.length === 0) continue;
    
    var className = String(row[classCol] || '1반').trim();
    var id = String(row[idCol] || '').trim();
    var name = String(row[nameCol] || '').trim();
    var gender = genderCol >= 0 ? String(row[genderCol] || '').trim() : '';
    
    if (!name || name === 'undefined') continue;
    
    gender = (gender.includes('여') || gender.toLowerCase() === 'f') ? '여' : '남';
    
    if (!result[className]) result[className] = [];
    result[className].push({
      id: id || String(result[className].length + 1),
      name: name,
      gender: gender
    });
  }
  
  return result;
}

function showExcelPreview(data) {
  var container = document.getElementById('excelPreviewContent');
  container.innerHTML = '';
  
  Object.keys(data).forEach(function(className) {
    var list = data[className];
    var div = document.createElement('div');
    div.className = 'preview-class';
    div.innerHTML = '<strong>' + escapeHtml(className) + ' (' + list.length + '명)</strong>';
    
    var preview = document.createElement('div');
    preview.className = 'preview-students';
    preview.textContent = list.slice(0, 5).map(function(s) {
      return s.id + '번 ' + s.name;
    }).join(', ');
    if (list.length > 5) {
      preview.textContent += ' 외 ' + (list.length - 5) + '명...';
    }
    
    div.appendChild(preview);
    container.appendChild(div);
  });
  
  document.getElementById('excelPreview').style.display = 'block';
}

function applyExcelData() {
  if (!parsedExcelData) return;
  
  var classNames = Object.keys(parsedExcelData);
  classNames.forEach(function(className, idx) {
    if (idx >= MAX_CLASSES) return;
    
    allClassData[idx] = allClassData[idx] || {};
    allClassData[idx].className = className;
    allClassData[idx].students = parsedExcelData[className];
    allClassData[idx].currentSeats = {};
    allClassData[idx].lockedSeats = {};
    allClassData[idx].relations = { close: {}, far: {} };
    allClassData[idx].studentMemos = {};
    allClassData[idx].savedRecords = allClassData[idx].savedRecords || [];
    allClassData[idx].specialNotes = allClassData[idx].specialNotes || [];
    allClassData[idx].config = allClassData[idx].config || JSON.parse(JSON.stringify(config));
  });
  
  currentClassIndex = 0;
  loadCurrentClassData();
  initClassDropdown();
  saveToLocalStorage();
  closeModal('excelModal');
  
  if (students.length > 0) {
    hideSetupWizard();
    renderSeating();
  }
  
  showToast(classNames.length + '개 학급 업로드 완료!');
}

function debounce(func, wait) {
  var timeout;
  return function() {
    var context = this, args = arguments;
    clearTimeout(timeout);
    timeout = setTimeout(function() {
      func.apply(context, args);
    }, wait);
  };
}


// ===== 레이아웃 =====
function selectColumns(cols) {
  config.columns = cols;
  config.layoutIndex = 0;
  document.querySelectorAll('#columnOptions .option-btn').forEach(function(btn, i) {
    btn.classList.toggle('active', [3,4,5,6,7][i] === cols);
  });
  updateLayoutOptions();
  updatePreview();
  updateEmptySeatInfo();
}

function switchClass() {
  saveCurrentClassData();
  currentClassIndex = parseInt(document.getElementById('currentClass').value, 10) || 0;
  loadCurrentClassData();
  initClassDropdown();
  refreshClassExtras();

  // 출결관리에서 학급을 바꿀 때는 명렬표가 비어 있어도 출결 화면을 유지합니다.
  if (appMode === 'attend') {
    var wizard = document.getElementById('setupWizard');
    var stage = document.getElementById('stage');
    var view = document.getElementById('attendView');
    var bar = document.querySelector('.action-bar');
    if (wizard) { wizard.classList.add('hidden'); wizard.style.display = 'none'; }
    setModeScreen(stage, false);
    setModeScreen(bar, false);
    setModeScreen(view, true, 'flex');
    Attend.render();
    return;
  }

  if (students.length === 0) {
    showSetupWizard();
  } else {
    hideSetupWizard();
    renderSeating();
  }
}
function saveCurrentClassData() {
  allClassData[currentClassIndex] = {
    className: allClassData[currentClassIndex] && allClassData[currentClassIndex].className 
      ? allClassData[currentClassIndex].className 
      : ((currentClassIndex + 1) + '반'),
    students: students,
    currentSeats: currentSeats,
    lockedSeats: lockedSeats,
    config: config,
    relations: relations,
    studentMemos: studentMemos,
    specialNotes: (allClassData[currentClassIndex] && allClassData[currentClassIndex].specialNotes) || [],
    savedRecords: savedRecords,
    notice: (allClassData[currentClassIndex] && allClassData[currentClassIndex].notice) || '',
    noticeAt: (allClassData[currentClassIndex] && allClassData[currentClassIndex].noticeAt) || '',
    praise: (allClassData[currentClassIndex] && allClassData[currentClassIndex].praise) || {},
    attendance: (allClassData[currentClassIndex] && allClassData[currentClassIndex].attendance) || {},
    duties: (allClassData[currentClassIndex] && allClassData[currentClassIndex].duties) || {},
    dutyRoles: (allClassData[currentClassIndex] && allClassData[currentClassIndex].dutyRoles) || '',
    attendReasons: (allClassData[currentClassIndex] && allClassData[currentClassIndex].attendReasons) || {},
    attendMemos: (allClassData[currentClassIndex] && allClassData[currentClassIndex].attendMemos) || {}
  };
  saveToLocalStorage();
}

function refreshClassExtras() {
  if (typeof Notice !== 'undefined') Notice.renderPreview();
  if (typeof Praise !== 'undefined') Praise.render();
  if (typeof Attend !== 'undefined') Attend.render();
  if (typeof Duty !== 'undefined') Duty.render();
}

function loadCurrentClassData() {
  var data = allClassData[currentClassIndex];
  if (data) {
    students = sortStudentsByName(data.students || []);
    currentSeats = data.currentSeats || {};
    lockedSeats = data.lockedSeats || {};
    config = data.config || { totalDesks: 30, columns: 5, seatType: 'pair', layoutIndex: 0, rows: 6 };
    relations = data.relations || { close: {}, far: {} };
    studentMemos = data.studentMemos || {};
    savedRecords = data.savedRecords || [];
    syncSameGenderSeatingOption();
  }
}

function syncSameGenderSeatingOption() {
  var input = document.getElementById('sameGenderSeating');
  var orderMode = getNumberOrderMode();
  var useNumberOrder = orderMode !== 'random';
  if (input) {
    input.checked = !!config.sameGenderSeating;
    input.disabled = useNumberOrder;
    if (input.parentElement) {
      input.parentElement.hidden = false;
      input.parentElement.classList.toggle('sp-seat-option-inactive', useNumberOrder);
    }
  }
  document.querySelectorAll('input[name="numberOrderMode"]').forEach(function(orderInput) {
    orderInput.checked = orderInput.value === orderMode;
  });
  var orderSummary = document.getElementById('numberOrderSummary');
  if (orderSummary) {
    orderSummary.textContent = orderMode === 'row' ? '학번 가로순' : (orderMode === 'column' ? '학번 세로순' : '무작위');
  }
  var hint = document.getElementById('numberOrderHint');
  if (hint) {
    hint.textContent = orderMode === 'row'
      ? '앞줄 왼쪽부터 오른쪽으로 채운 뒤 다음 행으로 이동합니다.'
      : (orderMode === 'column'
        ? '왼쪽 열의 앞자리부터 뒷자리로 채운 뒤 다음 열로 이동합니다.'
        : '무작위 추첨은 관계·성별 조건을 반영할 수 있습니다.');
  }
  var mainButton = document.getElementById('shuffleBtn');
  var proxyButton = document.querySelector('.sp-actions [data-proxy="shuffleBtn"]');
  var buttonText = orderMode === 'row'
    ? '↔ 학번 가로순 배치'
    : (orderMode === 'column' ? '↕ 학번 세로순 배치' : '🎱 추첨 시작');
  if (mainButton) mainButton.textContent = buttonText;
  if (proxyButton) proxyButton.textContent = buttonText;
}

function getNumberOrderMode() {
  if (config.numberOrderMode === 'row' || config.numberOrderMode === 'column') return config.numberOrderMode;
  return config.numberOrderSeating ? 'row' : 'random';
}

function usesNumberOrderSeating() {
  return getNumberOrderMode() !== 'random';
}

function orderSeatsForNumberMode(seats) {
  var ordered = seats.slice();
  if (getNumberOrderMode() === 'column') {
    ordered.sort(function(a, b) {
      return a.col - b.col || a.row - b.row;
    });
  }
  return ordered;
}

function getLayouts(cols) {
  return ALL_LAYOUTS[cols] || [[cols]];
}

function getCurrentLayout() {
  var layouts = getLayouts(config.columns);
  var idx = Math.min(config.layoutIndex, layouts.length - 1);
  if (config.seatType === 'single') return Array(config.columns).fill(1);
  if (config.seatType === 'group') {
    var gs = GROUP_SIZES[Math.min(config.layoutIndex, GROUP_SIZES.length - 1)] || 4;
    return Array(config.columns).fill(gs);
  }
  if (config.seatType === 'triple') {
    var tripleLayouts = TRIPLE_LAYOUTS[config.columns] || [[config.columns]];
    return tripleLayouts[Math.min(config.layoutIndex, tripleLayouts.length - 1)];
  }
  return layouts[idx];
}

function getLayoutString(layout) { return layout.join('-'); }
function getTotalSeatsPerRow() { return getCurrentLayout().reduce(function(a, b) { return a + b; }, 0); }
function calculateRows() { return Math.ceil(config.totalDesks / getTotalSeatsPerRow()); }

function updateLayoutOptions() {
  var container = document.getElementById('layoutOptions');
  var step = document.getElementById('layoutStep');
  var title = document.getElementById('layoutStepTitle');
  var layouts = getLayouts(config.columns);

  if (title) title.textContent = config.seatType === 'group' ? '👨‍👩‍👦 모둠 인원' : '🎯 레이아웃';
  var colTitle = document.getElementById('columnStepTitle');
  if (colTitle) colTitle.textContent = config.seatType === 'group' ? '📐 가로 모둠 수' : '📐 열(분단) 수';

  // 모둠 배치일 때는 이 자리에서 모둠 인원을 고릅니다
  if (config.seatType === 'group') {
    step.dataset.available = '1';
    step.style.display = '';
    step.style.visibility = 'visible';
    container.innerHTML = '';
    GROUP_SIZES.forEach(function(size, idx) {
      var btn = document.createElement('div');
      btn.className = 'layout-btn group-size-card' + (config.layoutIndex === idx ? ' active' : '');
      var dots = '';
      for (var i = 0; i < size; i++) dots += '<div class="mini-seat"></div>';
      btn.innerHTML = '<div class="group-size-preview">' + dots + '</div><div class="layout-text">' + size + '명</div>';
      btn.addEventListener('click', function() { selectLayout(idx); });
      container.appendChild(btn);
    });
    return;
  }

  if (config.seatType === 'single') {
    step.dataset.available = '0'; container.innerHTML = '';
    step.style.display = wizardStep === 2 ? '' : 'none'; step.style.visibility = 'hidden'; return;
  }
  if (config.seatType === 'triple') layouts = TRIPLE_LAYOUTS[config.columns] || [[config.columns]];
  if (layouts.length <= 1) {
    step.dataset.available = '0'; container.innerHTML = '';
    step.style.display = wizardStep === 2 ? '' : 'none'; step.style.visibility = 'hidden'; return;
  }
  
  step.dataset.available = '1';
  step.style.display = '';
  container.innerHTML = '';
  
  layouts.forEach(function(layout, idx) {
    var btn = document.createElement('div');
    btn.className = 'layout-btn' + (idx === config.layoutIndex ? ' active' : '');
    btn.addEventListener('click', function() { selectLayout(idx); });
    
    var previewHtml = '<div class="layout-preview">';
    layout.forEach(function(seats) {
      previewHtml += '<div class="layout-group">';
      for (var i = 0; i < seats; i++) previewHtml += '<div class="mini-seat"></div>';
      previewHtml += '</div>';
    });
    btn.innerHTML = previewHtml + '</div><div class="layout-text">' + layout.join('-') + '</div>';
    container.appendChild(btn);
  });
}

function selectLayout(idx) {
  config.layoutIndex = idx;
  document.querySelectorAll('.layout-btn').forEach(function(btn, i) {
    btn.classList.toggle('active', i === idx);
  });
  updatePreview();
  updateEmptySeatInfo();
}

function selectSeatType(type) {
  var wasGroup = config.seatType === 'group';
  config.seatType = type;
  document.querySelectorAll('.type-card').forEach(function(card) {
    card.classList.toggle('active', card.dataset.type === type);
  });
  // 모둠 ↔ 일반 전환 시 선택 가능한 열(모둠) 수가 달라집니다
  if (wasGroup !== (type === 'group')) {
    config.layoutIndex = 0;
    initColumnOptions();
  }
  updateLayoutOptions();
  syncSameGenderSeatingOption();
  updatePreview();
  updateEmptySeatInfo();
}

// ===== 독립 홈 · 설정 위자드 =====
function setModeScreen(element, visible, displayValue) {
  if (!element) return;
  element.hidden = !visible;
  element.classList.toggle('mode-screen-hidden', !visible);
  element.style.display = visible ? (displayValue || '') : 'none';
}

// 교실 좌석 구조 설정 폼(책상 수·분단·학생 수)을 현재 설정값으로 채웁니다.
function prepareSeatConfigForm() {
  document.getElementById('totalDesks').value = config.totalDesks;
  document.getElementById('studentCount').value = students.length || 28;
  initColumnOptions();
  updateLayoutOptions();
  updatePreview();
  updateEmptySeatInfo();
  if (typeof showWizardStep === 'function') showWizardStep(1);
}

// 홈 대시보드 — 헤더의 [홈 대시보드]·홈 메뉴·앱 시작 시 보이는 화면
function showHomeScreen() {
  appMode = 'home';
  document.querySelectorAll('.mode-tab').forEach(function (tab) { tab.classList.remove('active'); });
  var homeTab = document.getElementById('topDashboardBtn');
  if (homeTab) homeTab.classList.add('active');
  var wizard = document.getElementById('setupWizard');
  var stage = document.getElementById('stage');
  var attend = document.getElementById('attendView');
  var bar = document.querySelector('.action-bar');
  if (wizard) {
    wizard.classList.remove('hidden');
    wizard.classList.remove('home-only');
    wizard.classList.remove('seat-setup-only');
    wizard.style.display = 'grid';
  }
  setModeScreen(stage, false);
  setModeScreen(attend, false);
  setModeScreen(bar, false);
  prepareSeatConfigForm();
  window.scrollTo({ top:0, behavior:'auto' });
}

// 교실 좌석 구조 설정 화면.
// [학급 관리(자리)]에서 들어왔을 때(appMode === 'seat')는 대시보드를 빼고 설정 폼만 보여 줍니다.
function showSetupWizard() {
  var wizard = document.getElementById('setupWizard');
  var seatOnly = appMode === 'seat';
  if (wizard) {
    wizard.classList.remove('hidden');
    wizard.classList.remove('home-only');
    wizard.classList.toggle('seat-setup-only', seatOnly);
    wizard.style.display = 'grid';
  }
  setModeScreen(document.getElementById('stage'), false);
  if (appMode === 'seat') setModeScreen(document.querySelector('.action-bar'), false);
  prepareSeatConfigForm();
}

function hideSetupWizard() {
  var wizard = document.getElementById('setupWizard');
  if (wizard) {
    wizard.classList.remove('home-only');
    wizard.classList.remove('seat-setup-only');
    wizard.classList.add('hidden');
    wizard.style.display = 'none';
  }
  // 대시보드에서 바로 자리 화면으로 넘어온 경우(명렬표 등록 등) 탭 표시도 자리 화면으로 맞춥니다.
  if (appMode === 'home') {
    appMode = 'seat';
    document.querySelectorAll('.mode-tab').forEach(function (tab) {
      tab.classList.toggle('active', tab.dataset.mode === 'seat');
    });
  }
  setModeScreen(document.getElementById('stage'), true, 'flex');
  // 상단 원본 액션은 좌측 메뉴와 우측 좌석 패널의 프록시 연결용으로만 숨겨 둡니다.
  // 별도 버튼 행을 표시하지 않아 좌석표의 세로 공간을 확보합니다.
  if (appMode === 'seat') setModeScreen(document.querySelector('.action-bar'), false);
}

function adjustDeskCount(delta) {
  var input = document.getElementById('totalDesks');
  var val = Math.max(10, Math.min(50, (parseInt(input.value) || 30) + delta));
  input.value = val;
  config.totalDesks = val;
  updatePreview();
  updateEmptySeatInfo();
}

function updateEmptySeatInfo() {
  var deskCount = parseInt(document.getElementById('totalDesks').value) || 30;
  var studentCount = parseInt(document.getElementById('studentCount').value) || 0;
  config.totalDesks = deskCount;
  
  var layout = getCurrentLayout();
  var seatsPerRow = layout.reduce(function(a, b) { return a + b; }, 0);
  var rows = Math.ceil(deskCount / seatsPerRow);
  var actualSeats = seatsPerRow * rows;
  var emptyCount = actualSeats - studentCount;
  
  var deskDisplay = document.getElementById('deskCountDisplay');
  var infoEl = document.getElementById('emptySeatInfo');
  
  if (deskDisplay) deskDisplay.textContent = deskCount;
  if (infoEl) {
    if (emptyCount > 0) {
      infoEl.innerHTML = '→ 빈자리 <strong>' + emptyCount + '</strong>개';
      infoEl.style.background = '#e3f2fd';
      infoEl.style.color = '#1976d2';
    } else if (emptyCount < 0) {
      infoEl.innerHTML = '⚠️ 책상 부족 <strong>' + Math.abs(emptyCount) + '</strong>개';
      infoEl.style.background = '#ffebee';
      infoEl.style.color = '#c62828';
    } else {
      infoEl.innerHTML = '→ 모든 자리 배정';
      infoEl.style.background = '#e8f5e9';
      infoEl.style.color = '#2e7d32';
    }
  }
}

function updatePreview() {
  var grid = document.getElementById('previewGrid');
  if (!grid) return;
  
  config.rows = calculateRows();
  var layout = getCurrentLayout();
  grid.innerHTML = '';
  
  for (var r = 0; r < Math.min(config.rows, 8); r++) {
    var rowDiv = document.createElement('div');
    rowDiv.className = 'preview-row';
    layout.forEach(function(seats) {
      var group = document.createElement('div');
      group.className = 'preview-group';
      for (var s = 0; s < seats; s++) {
        var seat = document.createElement('div');
        seat.className = 'preview-seat';
        group.appendChild(seat);
      }
      rowDiv.appendChild(group);
    });
    grid.appendChild(rowDiv);
  }
  
  var previewLayout = document.getElementById('previewLayout');
  if (previewLayout) previewLayout.textContent = getLayoutString(layout);
}

function applyConfig() {
  var studentCount = parseInt(document.getElementById('studentCount').value) || 28;
  config.totalDesks = parseInt(document.getElementById('totalDesks').value) || 30;
  config.rows = calculateRows();
  
  if (students.length === 0) {
    students = [];
    for (var i = 1; i <= studentCount; i++) {
      students.push({ id: String(i), name: '학생' + i, gender: i % 2 === 0 ? '여' : '남' });
    }
  } else if (studentCount > students.length) {
    for (var j = students.length + 1; j <= studentCount; j++) {
      students.push({ id: String(j), name: '학생' + j, gender: j % 2 === 0 ? '여' : '남' });
    }
  } else if (studentCount < students.length) {
    students = students.slice(0, studentCount);
  }
  
  currentSeats = {};
  lockedSeats = {};
  initializeEmptySeats();
  saveCurrentClassData();
  hideSetupWizard();
  renderSeating();
  showToast('설정 완료!');
}

function initializeEmptySeats() {
  var layout = getCurrentLayout();
  for (var r = 0; r < config.rows; r++) {
    var colIndex = 0;
    layout.forEach(function(seatsInGroup) {
      for (var s = 0; s < seatsInGroup; s++) {
        var seatId = r + '-' + colIndex;
        if (!currentSeats[seatId]) currentSeats[seatId] = null;
        colIndex++;
      }
    });
  }
}

// ===== 자리 배치 렌더링 =====
function renderSeating() {
  var area = document.getElementById('seatingArea');
  area.innerHTML = '';
  var layout = getCurrentLayout();
  
  for (var r = 0; r < config.rows; r++) {
    var rowDiv = document.createElement('div');
    rowDiv.className = 'seat-row';
    var colIndex = 0;
    
    layout.forEach(function(seatsInGroup, groupIndex) {
      var group = document.createElement('div');
      group.className = 'seat-group';
      if (config.seatType === 'group') {
        group.className += ' cluster cluster-' + seatsInGroup;
        group.setAttribute('data-group-label', (r * layout.length + groupIndex + 1) + '모둠');
      }
      
      for (var s = 0; s < seatsInGroup; s++) {
        var seatId = r + '-' + colIndex;
        var seatData = currentSeats[seatId];
        var isLocked = lockedSeats[seatId];
        
        var seatCard = document.createElement('div');
        seatCard.className = 'seat-card';
        seatCard.dataset.seatId = seatId;
        seatCard.dataset.row = r;
        seatCard.dataset.col = colIndex;
        seatCard.dataset.group = groupIndex;
        seatCard.setAttribute('draggable', 'true');
        
        if (isLocked) seatCard.classList.add('locked');
        
        if (!seatData || !seatData.studentId) {
          seatCard.classList.add('empty');
          if (isLotteryRunning) seatCard.classList.add('waiting');
          seatCard.innerHTML = '<span class="empty-text">' + (isLotteryRunning ? '?' : '빈자리') + '</span>';
          (function(sid, gi, row) {
            seatCard.addEventListener('click', function() {
              if (!isLotteryRunning) openEmptySeatModal(sid, gi, row);
            });
          })(seatId, groupIndex, r);
        } else {
          var student = students.find(function(st) { return st.id === seatData.studentId; });
          if (student) {
            if (student.gender === '남') seatCard.classList.add('male');
            else if (student.gender === '여') seatCard.classList.add('female');
          }
          
          var pts = (typeof Praise !== 'undefined') ? Praise.get(seatData.studentId) : 0;
          var praiseGrowth = (typeof Praise !== 'undefined') ? Praise.stage(pts) : null;
          var seatWhere = config.seatType === 'group'
            ? (r * layout.length + groupIndex + 1) + '모둠'
            : (groupIndex + 1) + '분단 ' + (r + 1) + '행';
          seatCard.innerHTML = '<span class="student-id">' + seatData.studentId + '번</span>' +
            '<span class="student-name">' + seatData.name + '</span>' +
            '<span class="seat-number">' + seatWhere + '</span>' +
            (praiseGrowth ? '<span class="praise-badge stage-' + praiseGrowth.key + '" title="칭찬 ' + praiseGrowth.score + '점 · ' + praiseGrowth.label + '" aria-label="칭찬 ' + praiseGrowth.score + '점, ' + praiseGrowth.label + '"><span>' + praiseGrowth.icon + '</span><b>' + praiseGrowth.score + '</b></span>' : '') +
            ((typeof Attend !== 'undefined' && Attend.mark(seatData.studentId))
              ? '<span class="attend-badge s' + Attend.get(seatData.studentId) + '">' + Attend.mark(seatData.studentId) + '</span>' : '') +
            '<button type="button" class="lock-btn">' + (isLocked ? '🔒' : '🔓') + '</button>';
          
          (function(studentId, sId, gi, row) {
            seatCard.addEventListener('click', function(e) {
              if (!e.target.classList.contains('lock-btn') && !isLotteryRunning) {
                openStudentSettings(studentId);
              }
            });
            seatCard.addEventListener('dblclick', function(e) {
              e.preventDefault();
              if (!isLotteryRunning) openAssignSeatModal(studentId, sId, gi, row);
            });
          })(seatData.studentId, seatId, groupIndex, r);
          
          var lockBtn = seatCard.querySelector('.lock-btn');
          (function(sid, sd) {
            lockBtn.addEventListener('click', function(e) {
              e.stopPropagation();
              toggleLock(sid, sd);
            });
          })(seatId, seatData);
        }
        
        group.appendChild(seatCard);
        colIndex++;
      }
      rowDiv.appendChild(group);
    });
    area.appendChild(rowDiv);
  }
  
  applyTeacherScreenOrder();
  initDragAndDrop();
  fitSeatingToViewport();
  if (typeof renderSeatPanel === 'function') renderSeatPanel();
}

function applyTeacherScreenOrder() {
  if (!isTeacherScreenView) return;
  var area = document.getElementById('seatingArea');
  if (!area) return;
  Array.prototype.slice.call(area.children).reverse().forEach(function(row) {
    Array.prototype.slice.call(row.children).reverse().forEach(function(group) {
      Array.prototype.slice.call(group.children).reverse().forEach(function(seat) {
        group.appendChild(seat);
      });
      row.appendChild(group);
    });
    area.appendChild(row);
  });
}

function syncTeacherScreenView() {
  var main = document.getElementById('mainArea');
  var desk = main && main.querySelector('.desk');
  var classroom = main && main.querySelector('.classroom');
  var button = document.getElementById('screenTeacherViewBtn');
  if (main) main.classList.toggle('teacher-screen-view', isTeacherScreenView);
  if (main && desk && classroom) {
    if (isTeacherScreenView) main.insertBefore(desk, classroom.nextSibling);
    else main.insertBefore(desk, classroom);
  }
  if (button) {
    button.setAttribute('aria-pressed', isTeacherScreenView ? 'true' : 'false');
    button.textContent = isTeacherScreenView ? '🧑‍🎓 학생 시점으로' : '👩‍🏫 교탁 시점';
    button.title = isTeacherScreenView ? '학생이 칠판을 보는 방향으로 되돌립니다' : '교탁에서 학생을 마주 보는 방향으로 뒤집습니다';
  }
}

function toggleTeacherScreenView() {
  isTeacherScreenView = !isTeacherScreenView;
  syncTeacherScreenView();
  renderSeating();
  showToast(isTeacherScreenView ? '교탁 시점으로 뒤집었습니다' : '학생 시점으로 되돌렸습니다');
}

function toggleLock(seatId, seatData) {
  if (lockedSeats[seatId]) {
    delete lockedSeats[seatId];
    showToast('고정 해제');
  } else {
    lockedSeats[seatId] = seatData ? seatData.studentId : 'empty';
    showToast('자리 고정');
  }
  saveCurrentClassData();
  renderSeating();
}

function openEmptySeatModal(seatId, group, row) {
  selectedEmptySeatId = seatId;
  var isLocked = lockedSeats[seatId];
  var actualTotalSeats = getTotalSeatsPerRow() * config.rows;
  var maxEmptySeats = actualTotalSeats - students.length;
  var currentLockedEmpty = Object.keys(lockedSeats).filter(function(k) {
    return lockedSeats[k] === 'empty';
  }).length;
  
  document.getElementById('emptySeatInfo2').textContent = (group + 1) + '분단 ' + (row + 1) + '행 빈자리';
  
  var lockBtn = document.getElementById('toggleEmptyLockBtn');
  var lockText = document.getElementById('emptyLockBtnText');
  
  if (isLocked) {
    lockText.textContent = '🔓 고정 해제';
    lockBtn.disabled = false;
  } else if (currentLockedEmpty >= maxEmptySeats) {
    lockText.textContent = '⚠️ 고정 한도 초과';
    lockBtn.disabled = true;
  } else {
    lockText.textContent = '🔒 이 자리 고정 (' + currentLockedEmpty + '/' + maxEmptySeats + ')';
    lockBtn.disabled = false;
  }
  
  openModal('emptySeatModal');
}

function toggleEmptyLock() {
  if (!selectedEmptySeatId) return;
  
  if (lockedSeats[selectedEmptySeatId]) {
    delete lockedSeats[selectedEmptySeatId];
    showToast('고정 해제');
  } else {
    var actualTotalSeats = getTotalSeatsPerRow() * config.rows;
    var maxEmptySeats = actualTotalSeats - students.length;
    var currentLockedEmpty = Object.keys(lockedSeats).filter(function(k) {
      return lockedSeats[k] === 'empty';
    }).length;
    
    if (currentLockedEmpty >= maxEmptySeats) {
      showToast('빈자리 고정 한도 초과', 'error');
      closeModal('emptySeatModal');
      return;
    }
    lockedSeats[selectedEmptySeatId] = 'empty';
    showToast('빈자리 고정');
  }
  
  saveCurrentClassData();
  closeModal('emptySeatModal');
  renderSeating();
}

function initDragAndDrop() {
  document.querySelectorAll('.seat-card').forEach(function(seat) {
    seat.addEventListener('dragstart', handleDragStart);
    seat.addEventListener('dragend', handleDragEnd);
    seat.addEventListener('dragover', handleDragOver);
    seat.addEventListener('dragleave', handleDragLeave);
    seat.addEventListener('drop', handleDrop);
  });
}

function handleDragStart(e) {
  if (this.classList.contains('locked')) { e.preventDefault(); return; }
  draggedSeat = this;
  this.classList.add('dragging');
}
function handleDragEnd() { this.classList.remove('dragging'); draggedSeat = null; }
function handleDragOver(e) {
  e.preventDefault();
  if (!this.classList.contains('locked') && this !== draggedSeat) {
    this.classList.add('drag-over');
  }
}
function handleDragLeave() { this.classList.remove('drag-over'); }
function handleDrop(e) {
  e.preventDefault();
  this.classList.remove('drag-over');
  if (this.classList.contains('locked') || this === draggedSeat || !draggedSeat) return;
  
  var seatId1 = draggedSeat.dataset.seatId;
  var seatId2 = this.dataset.seatId;
  var temp = currentSeats[seatId1];
  currentSeats[seatId1] = currentSeats[seatId2];
  currentSeats[seatId2] = temp;
  
  saveCurrentClassData();
  renderSeating();
  showToast('자리 교환');
}

// ===== 추첨 =====
function startLottery() {
  if (isLotteryRunning) return;
  
  var lockedStudentIds = [];
  Object.keys(lockedSeats).forEach(function(key) {
    if (lockedSeats[key] !== 'empty') lockedStudentIds.push(lockedSeats[key]);
  });
  
  var availableStudents = students.filter(function(s) {
    return lockedStudentIds.indexOf(s.id) === -1;
  });
  
  if (availableStudents.length === 0) {
    showToast('배치할 학생이 없습니다');
    return;
  }
  
  var layout = getCurrentLayout();
  var allSeats = [];
  
  for (var r = 0; r < config.rows; r++) {
    var colIndex = 0;
    layout.forEach(function(seatsInGroup, groupIndex) {
      for (var s = 0; s < seatsInGroup; s++) {
        var seatId = r + '-' + colIndex;
        if (!lockedSeats[seatId]) {
          allSeats.push({ id: seatId, row: r, col: colIndex, group: groupIndex + 1 });
        }
        colIndex++;
      }
    });
  }
  
  var relationResult = usesNumberOrderSeating()
    ? { students: sortStudentsByNumber(availableStudents), seats: orderSeatsForNumberMode(allSeats) }
    : assignWithRelations(availableStudents, allSeats, layout);
  var shuffledStudents = relationResult.students;
  var shuffledSeats = relationResult.seats;
  
  var assignments = [];
  shuffledSeats.forEach(function(seat, index) {
    assignments.push({
      seat: seat,
      student: index < shuffledStudents.length ? shuffledStudents[index] : null
    });
  });
  
  allSeats.forEach(function(seat) { currentSeats[seat.id] = null; });
  
  isLotteryRunning = true;
  document.body.classList.add('lottery-active');   // 추첨 패널에 좌석이 가리지 않도록 여백 확보
  document.getElementById('lotteryPanel').classList.add('show');
  document.getElementById('currentBall').classList.remove('show');
  document.getElementById('progressFill').style.width = '0%';
  
  renderSeating();
  createBouncingBalls(shuffledStudents);
  
  var studentAssignments = assignments.filter(function(a) { return a.student; });
  var total = studentAssignments.length;
  var done = 0;
  var currentIndex = 0;

  // 소리는 이 클릭(사용자 제스처) 시점에 준비되므로 지연 없이 즉시 울립니다.
  Sound.init();
  Sound.startRoll();

  updateLotteryHud(0, total);
  var orderMode = getNumberOrderMode();
  setHeadline(orderMode === 'row' ? '↔ 학번 가로순 배치' : (orderMode === 'column' ? '↕ 학번 세로순 배치' : '🎉 자리 추첨'), false);

  function processNext() {
    if (currentIndex >= assignments.length || !isLotteryRunning) {
      finishLottery();
      return;
    }

    var assignment = assignments[currentIndex];
    if (!assignment.student) {
      currentIndex++;
      processNext();
      return;
    }

    var remaining = total - done;
    var isFinal = remaining === 1;
    var isClimax = remaining <= 3;

    if (isClimax) {
      setHeadline(isFinal ? '🥁 마지막 한 명!' : '🔥 이제 ' + remaining + '명 남았습니다', true);
    }

    dropBall(assignment.student, assignment.seat, done, total, isFinal);
    currentSeats[assignment.seat.id] = {
      studentId: assignment.student.id,
      name: assignment.student.name
    };

    done++;
    currentIndex++;
    updateLotteryHud(done, total);
    Sound.setTension(done / total, isClimax);

    setTimeout(processNext, stepInterval(done, total));
  }

  setTimeout(processNext, 700);
}

// 처음엔 뜸을 들이고, 중반엔 빠르게, 마지막 3명은 다시 느리게 — 긴장의 리듬
function stepInterval(done, total) {
  var remaining = total - done;
  var base = Math.max(260, Math.min(820, LOTTERY_BASE_MS / Math.max(total, 1)));
  var factor = 1;
  if (done <= 2) factor = 1.7;
  else if (remaining === 1) factor = 2.8;
  else if (remaining <= 3) factor = 2.2;
  else if (remaining <= 6) factor = 1.35;
  return base * factor * SPEED_MULTIPLIERS[lotterySpeed];
}

function setHeadline(text, isFinal) {
  var el = document.getElementById('lotteryHeadline');
  if (!el) return;
  el.textContent = text;
  el.classList.toggle('final', !!isFinal);
}

function updateLotteryHud(done, total) {
  var pct = total ? (done / total) * 100 : 0;
  document.getElementById('progressFill').style.width = pct + '%';
  document.getElementById('lotteryProgress').textContent = done + ' / ' + total;
  var remainEl = document.getElementById('lotteryRemain');
  if (remainEl) remainEl.textContent = Math.max(0, total - done);
}

function setLotterySpeed(speed) {
  if (!SPEED_MULTIPLIERS[speed]) return;
  lotterySpeed = speed;
  document.querySelectorAll('.speed-btn').forEach(function(b) {
    b.classList.toggle('active', b.dataset.speed === speed);
  });
}

function toggleLotterySound() {
  var on = Sound.toggle();
  var btn = document.getElementById('soundToggleBtn');
  if (btn) {
    btn.textContent = on ? '🔊 소리 켜짐' : '🔇 소리 꺼짐';
    btn.classList.toggle('muted', !on);
  }
  if (on && isLotteryRunning) Sound.startRoll();
}

function shuffleArray(arr) {
  for (var i = arr.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = arr[i];
    arr[i] = arr[j];
    arr[j] = temp;
  }
}

function createBouncingBalls(studentList) {
  var ballsArea = document.getElementById('ballsArea');
  ballsArea.innerHTML = '';
  bouncingBalls = [];
  
  studentList.forEach(function(student) {
    var ball = document.createElement('div');
    ball.className = 'bouncing-ball';
    ball.id = 'ball-' + student.id;
    ball.style.background = student.gender === '여' ? '#ffd54f' : '#64b5f6';
    ball.textContent = student.id;
    ball.style.left = Math.random() * 140 + 10 + 'px';
    ball.style.top = Math.random() * 35 + 10 + 'px';
    ballsArea.appendChild(ball);
    
    bouncingBalls.push({
      element: ball,
      student: student,
      x: parseFloat(ball.style.left),
      y: parseFloat(ball.style.top),
      vx: (Math.random() - 0.5) * 4,
      vy: (Math.random() - 0.5) * 4
    });
  });
  
  animateBalls();
}

function animateBalls() {
  if (!isLotteryRunning) return;
  
  bouncingBalls.forEach(function(ball) {
    if (!ball.element.classList.contains('falling')) {
      ball.x += ball.vx;
      ball.y += ball.vy;
      
      if (ball.x <= 0 || ball.x >= 140) ball.vx *= -0.9;
      if (ball.y <= 0 || ball.y >= 45) ball.vy *= -0.9;
      
      ball.x = Math.max(0, Math.min(140, ball.x));
      ball.y = Math.max(0, Math.min(45, ball.y));
      
      ball.vx += (Math.random() - 0.5) * 0.5;
      ball.vy += (Math.random() - 0.5) * 0.5;
      
      ball.element.style.left = ball.x + 'px';
      ball.element.style.top = ball.y + 'px';
    }
  });
  
  lotteryAnimationId = requestAnimationFrame(animateBalls);
}

function dropBall(student, seat, current, total, isFinal) {
  var ballEl = document.getElementById('ball-' + student.id);
  var seatCard = document.querySelector('[data-seat-id="' + seat.id + '"]');

  // 추첨기에서 뽑힌 공이 실제 좌석까지 날아갑니다
  if (ballEl) {
    if (seatCard) flyBallToSeat(ballEl, seatCard);
    ballEl.remove();
    bouncingBalls = bouncingBalls.filter(function(b) {
      return b.student.id !== student.id;
    });
  }

  var machine = document.querySelector('.lottery-machine');
  if (machine) {
    machine.classList.remove('shake');
    void machine.offsetWidth;
    machine.classList.add('shake');
  }

  showCurrentBall(student, seat);
  var seatLabel = document.getElementById('lotterySeatLabel');
  if (seatLabel) {
    seatLabel.textContent = '→ ' + seat.group + '분단 ' + (seat.row + 1) + '행';
    seatLabel.classList.add('show');
  }
  Sound.playPick(current, total, isFinal);

  setTimeout(function() {
    if (!seatCard) return;
    seatCard.classList.remove('empty', 'waiting');
    seatCard.classList.add(isFinal ? 'final-reveal' : 'just-assigned');
    if (student.gender === '남') seatCard.classList.add('male');
    else if (student.gender === '여') seatCard.classList.add('female');

    seatCard.innerHTML = '<span class="student-id">' + student.id + '번</span>' +
      '<span class="student-name">' + student.name + '</span>' +
      '<span class="seat-number">' + seat.group + '분단 ' + (seat.row + 1) + '행</span>';

    Confetti.burstAt(seatCard, isFinal ? 46 : 14);

    setTimeout(function() {
      seatCard.classList.remove('just-assigned', 'final-reveal');
    }, isFinal ? 1400 : 900);
  }, 380);
}

function flyBallToSeat(ballEl, seatCard) {
  var from = ballEl.getBoundingClientRect();
  var to = seatCard.getBoundingClientRect();
  if (!from.width || !to.width) return;

  var flyer = document.createElement('div');
  flyer.className = 'flying-ball';
  flyer.textContent = ballEl.textContent;
  flyer.style.left = from.left + 'px';
  flyer.style.top = from.top + 'px';
  flyer.style.width = from.width + 'px';
  flyer.style.height = from.height + 'px';
  flyer.style.background = ballEl.style.background;
  document.body.appendChild(flyer);

  var dx = (to.left + to.width / 2) - (from.left + from.width / 2);
  var dy = (to.top + to.height / 2) - (from.top + from.height / 2);

  requestAnimationFrame(function() {
    flyer.style.transform = 'translate(' + dx + 'px, ' + dy + 'px) scale(1.9)';
    flyer.style.opacity = '0';
  });
  setTimeout(function() { flyer.remove(); }, 560);
}

function showCurrentBall(student, seat) {
  var currentBall = document.getElementById('currentBall');
  currentBall.style.background = student.gender === '여' 
    ? 'linear-gradient(145deg, #ffd54f, #ffca28)' 
    : 'linear-gradient(145deg, #64b5f6, #42a5f5)';
  currentBall.querySelector('.ball-number').textContent = student.id + '번';
  currentBall.querySelector('.ball-name').textContent = student.name;
  currentBall.classList.remove('show');
  void currentBall.offsetWidth;
  currentBall.classList.add('show');
}

function skipLottery() {
  if (!isLotteryRunning) return;
  
  var layout = getCurrentLayout();
  var lockedStudentIds = [];
  Object.keys(lockedSeats).forEach(function(key) {
    if (lockedSeats[key] !== 'empty') lockedStudentIds.push(lockedSeats[key]);
  });
  
  var availableStudents = students.filter(function(s) {
    return lockedStudentIds.indexOf(s.id) === -1;
  });
  
  var allSeats = [];
  for (var r = 0; r < config.rows; r++) {
    var colIndex = 0;
    layout.forEach(function(seatsInGroup, groupIndex) {
      for (var s = 0; s < seatsInGroup; s++) {
        var seatId = r + '-' + colIndex;
        if (!lockedSeats[seatId] && (!currentSeats[seatId] || !currentSeats[seatId].studentId)) {
          allSeats.push({ id: seatId, row: r, col: colIndex, group: groupIndex + 1 });
        }
        colIndex++;
      }
    });
  }
  
  var placedIds = [];
  Object.keys(currentSeats).forEach(function(key) {
    if (currentSeats[key] && currentSeats[key].studentId) {
      placedIds.push(currentSeats[key].studentId);
    }
  });
  
  var remainingStudents = availableStudents.filter(function(s) {
    return placedIds.indexOf(s.id) === -1;
  });
  
  if (usesNumberOrderSeating()) {
    remainingStudents = sortStudentsByNumber(remainingStudents);
    orderSeatsForNumberMode(allSeats).forEach(function(seat, idx) {
      if (idx < remainingStudents.length) {
        currentSeats[seat.id] = {
          studentId: remainingStudents[idx].id,
          name: remainingStudents[idx].name
        };
      }
    });
  } else if (config.sameGenderSeating) {
    var genderResult = assignSameGenderGroups(remainingStudents, allSeats, layout, true);
    genderResult.seats.forEach(function(seat, idx) {
      var student = genderResult.students[idx];
      if (student) currentSeats[seat.id] = { studentId: student.id, name: student.name };
    });
  } else {
    shuffleArray(remainingStudents);
    allSeats.forEach(function(seat, idx) {
      if (idx < remainingStudents.length) {
        currentSeats[seat.id] = {
          studentId: remainingStudents[idx].id,
          name: remainingStudents[idx].name
        };
      }
    });
  }
  
  finishLottery();
}

function finishLottery() {
  var wasRunning = isLotteryRunning;
  isLotteryRunning = false;
  if (lotteryAnimationId) cancelAnimationFrame(lotteryAnimationId);

  document.body.classList.remove('lottery-active');
  Sound.stopRoll();
  if (wasRunning) {
    Sound.playFanfare();
    Confetti.celebrate();
  }

  setHeadline('🎉 자리 추첨', false);
  var seatLabel = document.getElementById('lotterySeatLabel');
  if (seatLabel) { seatLabel.textContent = ''; seatLabel.classList.remove('show'); }

  document.getElementById('lotteryPanel').classList.remove('show');
  document.getElementById('ballsArea').innerHTML = '';
  bouncingBalls = [];
  saveCurrentClassData();
  renderSeating();
  showToast('🎉 배치 완료!');
}

// ===== 같은 성별 우선 배치 =====
function getLotteryGender(student) {
  if (student && student.gender === '남') return '남';
  if (student && student.gender === '여') return '여';
  return '기타';
}

function assignSameGenderGroups(availableStudents, allSeats, layout, includeCurrentAnchors) {
  var pools = { '남': [], '여': [], '기타': [] };
  availableStudents.forEach(function(student) {
    pools[getLotteryGender(student)].push(student);
  });
  shuffleArray(pools['남']);
  shuffleArray(pools['여']);
  shuffleArray(pools['기타']);

  var availableSeatIds = {};
  var bucketsByKey = {};
  allSeats.forEach(function(seat) {
    availableSeatIds[seat.id] = true;
    var key = seat.row + '|' + seat.group;
    if (!bucketsByKey[key]) bucketsByKey[key] = { key: key, row: seat.row, group: seat.group, seats: [], anchorGenders: [] };
    bucketsByKey[key].seats.push(seat);
  });

  function addAnchor(seatId, studentId) {
    if (!studentId || studentId === 'empty' || availableSeatIds[seatId]) return;
    var parts = String(seatId).split('-');
    var row = parseInt(parts[0], 10);
    var col = parseInt(parts[1], 10);
    if (isNaN(row) || isNaN(col)) return;
    var group = getGroupIndexByCol(col, layout) + 1;
    var bucket = bucketsByKey[row + '|' + group];
    if (!bucket) return;
    var student = students.find(function(item) { return String(item.id) === String(studentId); });
    var gender = getLotteryGender(student);
    if (gender !== '기타' && bucket.anchorGenders.indexOf(gender) === -1) bucket.anchorGenders.push(gender);
  }

  Object.keys(lockedSeats).forEach(function(seatId) {
    addAnchor(seatId, lockedSeats[seatId]);
  });
  if (includeCurrentAnchors) {
    Object.keys(currentSeats).forEach(function(seatId) {
      var seatData = currentSeats[seatId];
      if (seatData && seatData.studentId) addAnchor(seatId, seatData.studentId);
    });
  }

  var buckets = Object.keys(bucketsByKey).map(function(key) { return bucketsByKey[key]; });
  shuffleArray(buckets);
  buckets.sort(function(a, b) {
    var aForced = a.anchorGenders.length === 1 ? 1 : 0;
    var bForced = b.anchorGenders.length === 1 ? 1 : 0;
    return bForced - aForced;
  });

  var assignments = {};
  var deferredSeats = [];

  function chooseGender(bucket) {
    if (bucket.anchorGenders.length === 1) return bucket.anchorGenders[0];
    var capacity = bucket.seats.length;
    var maleCount = pools['남'].length;
    var femaleCount = pools['여'].length;
    if (!maleCount && !femaleCount) return null;
    if (!maleCount) return '여';
    if (!femaleCount) return '남';

    var needed = Math.min(capacity, maleCount + femaleCount);
    var maleCanFill = maleCount >= needed;
    var femaleCanFill = femaleCount >= needed;
    if (maleCanFill && femaleCanFill) return Math.random() < 0.5 ? '남' : '여';
    if (maleCanFill) return '남';
    if (femaleCanFill) return '여';
    if (maleCount === femaleCount) return Math.random() < 0.5 ? '남' : '여';
    return maleCount > femaleCount ? '남' : '여';
  }

  function takeStudentsFromPool(gender, count) {
    var pool = pools[gender] || [];
    var picked = [];

    // 같은 성별로 설정한 가까이 배치 관계는 같은 짝·모둠 안에서 함께 반영합니다.
    while (picked.length + 1 < count) {
      var pairIndexes = null;
      for (var i = 0; i < pool.length && !pairIndexes; i++) {
        var targets = (relations.close && relations.close[pool[i].id]) || [];
        for (var j = i + 1; j < pool.length; j++) {
          if (targets.indexOf(pool[j].id) !== -1 || (((relations.close && relations.close[pool[j].id]) || []).indexOf(pool[i].id) !== -1)) {
            pairIndexes = [i, j];
            break;
          }
        }
      }
      if (!pairIndexes) break;
      var second = pool.splice(pairIndexes[1], 1)[0];
      var first = pool.splice(pairIndexes[0], 1)[0];
      picked.push(first, second);
    }
    while (picked.length < count && pool.length) picked.push(pool.pop());
    return picked;
  }

  buckets.forEach(function(bucket) {
    shuffleArray(bucket.seats);
    var gender = chooseGender(bucket);
    var pool = gender ? pools[gender] : [];
    var fillCount = Math.min(bucket.seats.length, pool.length);
    var picked = gender ? takeStudentsFromPool(gender, fillCount) : [];
    for (var i = 0; i < picked.length; i++) assignments[bucket.seats[i].id] = picked[i];
    for (var j = picked.length; j < bucket.seats.length; j++) deferredSeats.push(bucket.seats[j]);
  });

  var leftovers = pools['남'].concat(pools['여'], pools['기타']);
  shuffleArray(leftovers);
  shuffleArray(deferredSeats);
  deferredSeats.forEach(function(seat, index) {
    if (leftovers[index]) assignments[seat.id] = leftovers[index];
  });

  var resultStudents = [];
  var resultSeats = [];
  allSeats.forEach(function(seat) {
    if (assignments[seat.id]) {
      resultStudents.push(assignments[seat.id]);
      resultSeats.push(seat);
    }
  });
  return { students: resultStudents, seats: resultSeats };
}

// ===== 관계 설정 배치 알고리즘 =====
function assignWithRelations(availableStudents, allSeats, layout) {
  if (config.sameGenderSeating) return assignSameGenderGroups(availableStudents, allSeats, layout, false);
  var closePairs = [];
  var farPairs = [];
  
  Object.keys(relations.close).forEach(function(studentId) {
    (relations.close[studentId] || []).forEach(function(targetId) {
      var exists = closePairs.some(function(p) {
        return (p.a === studentId && p.b === targetId) || (p.a === targetId && p.b === studentId);
      });
      if (!exists && availableStudents.some(function(s) { return s.id === studentId; }) && 
          availableStudents.some(function(s) { return s.id === targetId; })) {
        closePairs.push({ a: studentId, b: targetId });
      }
    });
  });
  
  Object.keys(relations.far).forEach(function(studentId) {
    (relations.far[studentId] || []).forEach(function(targetId) {
      var exists = farPairs.some(function(p) {
        return (p.a === studentId && p.b === targetId) || (p.a === targetId && p.b === studentId);
      });
      if (!exists && availableStudents.some(function(s) { return s.id === studentId; }) && 
          availableStudents.some(function(s) { return s.id === targetId; })) {
        farPairs.push({ a: studentId, b: targetId });
      }
    });
  });
  
  allSeats.forEach(function(seat) {
    seat.groupIndex = getGroupIndexByCol(seat.col, layout);
  });
  
  var seatMap = {};
  allSeats.forEach(function(seat) {
    seatMap[seat.row + '-' + seat.col] = seat;
  });
  
  var directions = [[-1,-1],[-1,0],[-1,1],[0,-1],[0,1],[1,-1],[1,0],[1,1]];
  var adjacentPairs = [];
  var pairSet = {};
  
  allSeats.forEach(function(seat) {
    directions.forEach(function(dir) {
      var neighborKey = (seat.row + dir[0]) + '-' + (seat.col + dir[1]);
      var neighborSeat = seatMap[neighborKey];
      if (neighborSeat) {
        var key1 = seat.id + '|' + neighborSeat.id;
        var key2 = neighborSeat.id + '|' + seat.id;
        if (!pairSet[key1] && !pairSet[key2]) {
          pairSet[key1] = true;
          adjacentPairs.push({ seat1: seat, seat2: neighborSeat });
        }
      }
    });
  });
  
  var assignments = {};
  var placedStudentIds = [];
  var usedSeatIds = [];
  
  shuffleArray(adjacentPairs);
  closePairs.forEach(function(pair) {
    if (placedStudentIds.indexOf(pair.a) !== -1 || placedStudentIds.indexOf(pair.b) !== -1) return;
    for (var i = 0; i < adjacentPairs.length; i++) {
      var seatPair = adjacentPairs[i];
      if (usedSeatIds.indexOf(seatPair.seat1.id) === -1 && usedSeatIds.indexOf(seatPair.seat2.id) === -1) {
        var studentA = availableStudents.find(function(s) { return s.id === pair.a; });
        var studentB = availableStudents.find(function(s) { return s.id === pair.b; });
        if (studentA && studentB) {
          assignments[seatPair.seat1.id] = studentA;
          assignments[seatPair.seat2.id] = studentB;
          placedStudentIds.push(pair.a, pair.b);
          usedSeatIds.push(seatPair.seat1.id, seatPair.seat2.id);
          break;
        }
      }
    }
  });
  
  var remainingStudents = availableStudents.filter(function(s) {
    return placedStudentIds.indexOf(s.id) === -1;
  });
  var remainingSeats = allSeats.filter(function(seat) {
    return usedSeatIds.indexOf(seat.id) === -1;
  });
  
  shuffleArray(remainingStudents);
  shuffleArray(remainingSeats);
  
  function violatesFarConstraint(studentId, seat) {
    var conflicts = farPairs.filter(function(p) {
      return p.a === studentId || p.b === studentId;
    });
    for (var i = 0; i < conflicts.length; i++) {
      var otherId = conflicts[i].a === studentId ? conflicts[i].b : conflicts[i].a;
      var otherSeatId = null;
      Object.keys(assignments).forEach(function(sId) {
        if (assignments[sId] && assignments[sId].id === otherId) otherSeatId = sId;
      });
      if (otherSeatId) {
        var otherSeat = allSeats.find(function(s) { return s.id === otherSeatId; });
        if (otherSeat && (Math.abs(seat.row - otherSeat.row) < 2 || Math.abs(seat.col - otherSeat.col) < 2)) {
          return true;
        }
      }
    }
    return false;
  }
  
  var unplacedStudents = [];
  remainingStudents.forEach(function(student) {
    var placed = false;
    for (var i = 0; i < remainingSeats.length; i++) {
      var seat = remainingSeats[i];
      if (usedSeatIds.indexOf(seat.id) === -1 && !violatesFarConstraint(student.id, seat)) {
        assignments[seat.id] = student;
        placedStudentIds.push(student.id);
        usedSeatIds.push(seat.id);
        placed = true;
        break;
      }
    }
    if (!placed) unplacedStudents.push(student);
  });
  
  if (unplacedStudents.length > 0) {
    var emptySeats = allSeats.filter(function(seat) {
      return usedSeatIds.indexOf(seat.id) === -1;
    });
    unplacedStudents.forEach(function(student, idx) {
      if (emptySeats[idx]) {
        assignments[emptySeats[idx].id] = student;
        placedStudentIds.push(student.id);
        usedSeatIds.push(emptySeats[idx].id);
      }
    });
  }
  
  var resultStudents = [];
  var resultSeats = [];
  allSeats.forEach(function(seat) {
    if (assignments[seat.id]) {
      resultStudents.push(assignments[seat.id]);
      resultSeats.push(seat);
    }
  });
  
  return { students: resultStudents, seats: resultSeats };
}

function getGroupIndexByCol(col, layout) {
  var cumulative = 0;
  for (var g = 0; g < layout.length; g++) {
    cumulative += layout[g];
    if (col < cumulative) return g;
  }
  return layout.length - 1;
}

// ===== 학생 설정 =====
function openStudentSettings(studentId) {
  selectedStudent = students.find(function(s) { return s.id === studentId; });
  if (!selectedStudent) return;
  
  var layout = getCurrentLayout();
  var studentSeatId = null;
  var studentGroup = null;
  var studentRow = null;
  
  Object.keys(currentSeats).forEach(function(seatId) {
    if (currentSeats[seatId] && currentSeats[seatId].studentId === studentId) {
      studentSeatId = seatId;
      var parts = seatId.split('-');
      studentRow = parseInt(parts[0]);
      var col = parseInt(parts[1]);
      studentGroup = getGroupIndexByCol(col, layout);
    }
  });
  
  document.getElementById('selectedStudentName').textContent = selectedStudent.id + '번 ' + selectedStudent.name;
  
  var seatDisplay = document.getElementById('currentSeatDisplay');
  var lockBtn = document.getElementById('modalLockBtn');
  var moveBtn = document.getElementById('modalMoveBtn');
  var praiseBtn = document.getElementById('studentPraiseBtn');
  var praiseMinusBtn = document.getElementById('studentPraiseMinusBtn');
  if (praiseBtn && praiseMinusBtn) {
    function refreshPraiseButtons() {
      var score = Praise.get(studentId);
      var growth = Praise.stage(score);
      praiseBtn.textContent = growth.icon + ' 칭찬 +1 · ' + growth.label + ' ' + score + '점';
      praiseMinusBtn.disabled = score <= 0;
    }
    refreshPraiseButtons();
    praiseBtn.onclick = function() {
      Praise.add(studentId, 1);
      refreshPraiseButtons();
      renderStudentSummary(selectedStudent);
      showToast(selectedStudent.name + ' 학생에게 칭찬 포인트를 주었습니다');
    };
    praiseMinusBtn.onclick = function() {
      if (Praise.get(studentId) <= 0) return;
      Praise.add(studentId, -1);
      refreshPraiseButtons();
      renderStudentSummary(selectedStudent);
      showToast(selectedStudent.name + ' 학생의 칭찬 포인트를 1점 차감했습니다');
    };
  }
    if (studentSeatId) {
    seatDisplay.textContent = (studentGroup + 1) + '분단 ' + (studentRow + 1) + '행';
    var isLocked = lockedSeats[studentSeatId];
    lockBtn.innerHTML = isLocked ? '🔓 고정 해제' : '🔒 자리 고정';
    lockBtn.style.display = '';
    moveBtn.style.display = '';
    
    lockBtn.onclick = function() {
      var seatData = currentSeats[studentSeatId];
      toggleLock(studentSeatId, seatData);
      closeModal('studentModal');
    };
    
    moveBtn.onclick = function() {
      closeModal('studentModal');
      setMode('seat');
      assignSeatStudent = {
        id: selectedStudent.id,
        name: selectedStudent.name,
        currentSeatId: studentSeatId
      };
      window.startSeatPicking();
    };
  } else {
    seatDisplay.textContent = '미배치';
    lockBtn.style.display = 'none';
    moveBtn.style.display = 'none';
  }
  
  // 메모·오늘 출결 로드
  renderStudentSummary(selectedStudent);
  renderStudentAttendanceQuick(selectedStudent);
  document.querySelectorAll('#studentModal .tab-btn').forEach(function(b){ b.classList.toggle('active', b.dataset.tab === 'summary'); });
  document.querySelectorAll('#studentModal .tab-content').forEach(function(c){ c.classList.toggle('active', c.id === 'summaryTab'); });

  var memoInput = document.getElementById('studentMemoInput');
  if (memoInput) {
    memoInput.value = studentMemos[studentId] || '';
  }
  
  tempRelations = {
    close: (relations.close[studentId] || []).slice(),
    far: (relations.far[studentId] || []).slice()
  };
  
  renderRelationChips();
  renderPairHistory();
  
  SpecialNotes.openForStudent(selectedStudent);
  openModal('studentModal');
}

function renderStudentAttendanceQuick(student) {
  var panel = document.getElementById('studentAttendanceQuick');
  if (!panel || !student || typeof Attend === 'undefined') return;

  var studentId = student.id;
  var dateKey = formatDate(new Date());
  var state = Attend.get(studentId, dateKey);
  var dateLabel = document.getElementById('studentAttendanceDate');
  var badge = document.getElementById('studentAttendanceBadge');
  var reasonSelect = document.getElementById('studentAttendanceReason');
  var memoInput = document.getElementById('studentAttendanceMemo');
  var isPresent = state === '출석';

  panel.dataset.studentId = String(studentId);
  panel.setAttribute('aria-label', student.name + ' 학생 오늘 출결');
  if (dateLabel) dateLabel.textContent = dateKey + ' · 오늘';
  if (badge) {
    badge.textContent = state;
    badge.dataset.state = state;
  }

  panel.querySelectorAll('[data-attendance-state]').forEach(function (button) {
    var buttonState = button.dataset.attendanceState;
    var active = buttonState === state;
    button.classList.toggle('active', active);
    button.setAttribute('aria-pressed', active ? 'true' : 'false');
    button.onclick = function () {
      if (memoInput && memoInput._attendanceSaveTimer) {
        clearTimeout(memoInput._attendanceSaveTimer);
        memoInput._attendanceSaveTimer = null;
        if (buttonState !== '출석') Attend.setMemo(studentId, memoInput.value, dateKey);
      }
      Attend.setState(studentId, buttonState === '출석' ? '' : buttonState, dateKey);
      renderStudentAttendanceQuick(student);
      renderStudentSummary(student);
      showToast(
        student.name + ' 학생을 ' +
        (buttonState === '출석' ? '출석으로 되돌렸습니다' : buttonState + '으로 기록했습니다')
      );
    };
  });

  if (reasonSelect) {
    reasonSelect.value = Attend.getReason(studentId, dateKey);
    reasonSelect.disabled = isPresent;
    reasonSelect.onchange = function () {
      Attend.setReason(studentId, this.value, dateKey);
      renderStudentSummary(student);
    };
  }

  if (memoInput) {
    if (memoInput._attendanceSaveTimer) clearTimeout(memoInput._attendanceSaveTimer);
    memoInput._attendanceSaveTimer = null;
    memoInput.value = Attend.getMemo(studentId, dateKey);
    memoInput.disabled = isPresent;

    var persistMemo = function () {
      if (memoInput._attendanceSaveTimer) clearTimeout(memoInput._attendanceSaveTimer);
      memoInput._attendanceSaveTimer = null;
      if (Attend.get(studentId, dateKey) === '출석') return;
      Attend.setMemo(studentId, memoInput.value, dateKey);
    };

    memoInput.oninput = function () {
      if (memoInput._attendanceSaveTimer) clearTimeout(memoInput._attendanceSaveTimer);
      memoInput._attendanceSaveTimer = setTimeout(persistMemo, 400);
    };
    memoInput.onchange = persistMemo;
    memoInput.onblur = persistMemo;
  }
}

function saveStudentMemo() {
  if (!selectedStudent) return;
  renderStudentSummary(selectedStudent);
  document.querySelectorAll('#studentModal .tab-btn').forEach(function(b){ b.classList.toggle('active', b.dataset.tab === 'summary'); });
  document.querySelectorAll('#studentModal .tab-content').forEach(function(c){ c.classList.toggle('active', c.id === 'summaryTab'); });

  var memoInput = document.getElementById('studentMemoInput');
  var memo = memoInput ? memoInput.value.trim() : '';
  if (memo) {
    studentMemos[selectedStudent.id] = memo;
  } else {
    delete studentMemos[selectedStudent.id];
  }
  saveCurrentClassData();
}

function saveStudentSettings() {
  saveStudentMemo();
  relations.close[selectedStudent.id] = tempRelations.close;
  relations.far[selectedStudent.id] = tempRelations.far;
  saveCurrentClassData();
  closeModal('studentModal');
  showToast('저장됨');
}

function openAssignSeatModal(studentId, currentSeatId, group, row) {
  var student = students.find(function(s) { return s.id === studentId; });
  if (!student) return;
  
  assignSeatStudent = {
    id: studentId,
    name: student.name,
    currentSeatId: currentSeatId
  };
  
  var info = (group + 1) + '분단 ' + (row + 1) + '행';
  var isLocked = lockedSeats[currentSeatId];
  
  var modalHtml = '<div class="modal show" id="assignSeatModal" style="z-index:1200;">' +
    '<div class="modal-content small">' +
    '<div class="modal-header"><h3>📍 ' + student.id + '번 ' + escapeHtml(student.name) + '</h3>' +
    '<button type="button" class="close-btn" onclick="document.getElementById(\'assignSeatModal\').remove();">&times;</button></div>' +
    '<div class="modal-body"><p style="margin-bottom:16px;color:#666;">현재 위치: <strong>' + info + '</strong></p>' +
    '<div style="display:flex;flex-direction:column;gap:12px;">' +
    '<button type="button" class="btn" onclick="toggleLockFromModal(\'' + currentSeatId + '\');" style="width:100%;justify-content:center;">' +
    (isLocked ? '🔓 고정 해제' : '🔒 이 자리 고정') + '</button>' +
    '<button type="button" class="btn primary" onclick="startSeatPicking();" style="width:100%;justify-content:center;">🎯 다른 자리로 이동</button>' +
    '</div></div></div></div>';
  
  document.body.insertAdjacentHTML('beforeend', modalHtml);
}

window.toggleLockFromModal = function(seatId) {
  var seatData = currentSeats[seatId];
  toggleLock(seatId, seatData);
  var modal = document.getElementById('assignSeatModal');
  if (modal) modal.remove();
};

window.startSeatPicking = function() {
  var modal = document.getElementById('assignSeatModal');
  if (modal) modal.remove();
  if (!assignSeatStudent) return;
  
  showToast('이동할 자리를 클릭하세요', 'info');
  
  document.querySelectorAll('.seat-card').forEach(function(card) {
    if (card.dataset.seatId !== assignSeatStudent.currentSeatId) {
      card.style.cursor = 'crosshair';
      card.classList.add('selectable-seat');
    }
  });
  
  var clickHandler = function(e) {
    var seatCard = e.target.closest('.seat-card');
    if (!seatCard || !assignSeatStudent) return;
    
    var targetSeatId = seatCard.dataset.seatId;
    if (targetSeatId === assignSeatStudent.currentSeatId) {
      showToast('같은 자리입니다', 'error');
      return;
    }
    if (lockedSeats[targetSeatId]) {
      showToast('고정된 자리입니다', 'error');
      return;
    }
    
    var temp = currentSeats[assignSeatStudent.currentSeatId];
    currentSeats[assignSeatStudent.currentSeatId] = currentSeats[targetSeatId];
    currentSeats[targetSeatId] = temp;
    
    saveCurrentClassData();
    renderSeating();
    showToast(assignSeatStudent.name + ' 자리 이동 완료!');
    
    assignSeatStudent = null;
    document.removeEventListener('click', clickHandler);
    cleanupSeatPicking();
  };
  
  var escHandler = function(e) {
    if (e.key === 'Escape') {
      assignSeatStudent = null;
      document.removeEventListener('click', clickHandler);
      document.removeEventListener('keydown', escHandler);
      cleanupSeatPicking();
      showToast('취소됨');
    }
  };
  
  setTimeout(function() {
    document.addEventListener('click', clickHandler);
    document.addEventListener('keydown', escHandler);
  }, 100);
};

function cleanupSeatPicking() {
  document.querySelectorAll('.seat-card').forEach(function(card) {
    card.style.cursor = '';
    card.classList.remove('selectable-seat');
  });
}

// ===== 관계 설정 UI =====
function renderRelationChips() {
  var closeList = document.getElementById('closeStudentsList');
  var farList = document.getElementById('farStudentsList');
  
  closeList.innerHTML = tempRelations.close.length === 0 ? '<span class="hint">없음</span>' : '';
  farList.innerHTML = tempRelations.far.length === 0 ? '<span class="hint">없음</span>' : '';
  
  tempRelations.close.forEach(function(id) {
    var s = students.find(function(st) { return st.id === id; });
    var chip = document.createElement('div');
    chip.className = 'chip close';
    chip.innerHTML = '<span>' + (s ? s.name : id) + '</span><button type="button" class="remove">&times;</button>';
    chip.querySelector('.remove').addEventListener('click', function() {
      tempRelations.close = tempRelations.close.filter(function(i) { return i !== id; });
      renderRelationChips();
    });
    closeList.appendChild(chip);
  });
  
  tempRelations.far.forEach(function(id) {
    var s = students.find(function(st) { return st.id === id; });
    var chip = document.createElement('div');
    chip.className = 'chip far';
    chip.innerHTML = '<span>' + (s ? s.name : id) + '</span><button type="button" class="remove">&times;</button>';
    chip.querySelector('.remove').addEventListener('click', function() {
      tempRelations.far = tempRelations.far.filter(function(i) { return i !== id; });
      renderRelationChips();
    });
    farList.appendChild(chip);
  });
}

function openStudentPicker(mode) {
  pickerMode = mode;
  renderStudentPicker();
  document.getElementById('studentPicker').classList.add('show');
}

function closeStudentPicker() {
  document.getElementById('studentPicker').classList.remove('show');
  document.getElementById('studentSearch').value = '';
}

function renderStudentPicker() {
  var list = document.getElementById('studentPickerList');
  var searchTerm = document.getElementById('studentSearch').value.toLowerCase();
  var excluded = [selectedStudent.id].concat(tempRelations.close).concat(tempRelations.far);
  
  var filtered = students.filter(function(s) {
    return excluded.indexOf(s.id) === -1 && 
           (s.name.toLowerCase().indexOf(searchTerm) !== -1 || s.id.indexOf(searchTerm) !== -1);
  });
  
  list.innerHTML = filtered.length === 0 ? '<p class="hint" style="padding:20px;">검색 결과 없음</p>' : '';
  
  filtered.forEach(function(s) {
    var item = document.createElement('div');
    item.className = 'picker-item';
    item.innerHTML = '<span>' + s.id + '번</span><span>' + s.name + '</span>';
    item.addEventListener('click', function() {
      tempRelations[pickerMode].push(s.id);
      renderRelationChips();
      closeStudentPicker();
    });
    list.appendChild(item);
  });
}

function filterStudents() {
  renderStudentPicker();
}

// ===== 짝꿍 이력 =====
function getSeatHistory(studentId) {
  var history = [];
  savedRecords.forEach(function(record) {
    var studentSeatId = null;
    Object.keys(record.seats).forEach(function(seatId) {
      var seat = record.seats[seatId];
      if (seat && seat.studentId === studentId) studentSeatId = seatId;
    });
    if (!studentSeatId) return;
    
    var parts = studentSeatId.split('-');
    var row = parseInt(parts[0]);
    var col = parseInt(parts[1]);
    var recordConfig = record.config || config;
    var layout = getLayoutFromConfig(recordConfig);
    var myGroup = getGroupIndexByCol(col, layout);
    
    var neighbors = [];
    [col - 1, col + 1].forEach(function(neighborCol) {
      if (neighborCol < 0) return;
      var neighborGroup = getGroupIndexByCol(neighborCol, layout);
      if (neighborGroup !== myGroup) return;
      var neighborSeatId = row + '-' + neighborCol;
      var neighborSeat = record.seats[neighborSeatId];
      if (neighborSeat && neighborSeat.studentId && neighborSeat.studentId !== studentId) {
        var neighborStudent = students.find(function(s) { return s.id === neighborSeat.studentId; });
        neighbors.push({
          id: neighborSeat.studentId,
          name: neighborSeat.name || (neighborStudent ? neighborStudent.name : '알수없음')
        });
      }
    });
    
    history.push({
      date: record.date,
      memo: record.memo,
      group: myGroup + 1,
      row: row + 1,
      neighbors: neighbors
    });
  });
  return history;
}

function getLayoutFromConfig(cfg) {
  if (cfg.seatType === 'single') return Array(cfg.columns).fill(1);
  if (cfg.seatType === 'group') {
    var gsz = GROUP_SIZES[Math.min(cfg.layoutIndex || 0, GROUP_SIZES.length - 1)] || 4;
    return Array(cfg.columns).fill(gsz);
  }
  if (cfg.seatType === 'triple') {
    var tripleLayouts = TRIPLE_LAYOUTS[cfg.columns] || [[cfg.columns]];
    return tripleLayouts[Math.min(cfg.layoutIndex || 0, tripleLayouts.length - 1)];
  }
  var layouts = ALL_LAYOUTS[cfg.columns] || [[cfg.columns]];
  return layouts[Math.min(cfg.layoutIndex || 0, layouts.length - 1)];
}

function renderPairHistory() {
  var list = document.getElementById('pairHistoryList');
  if (!selectedStudent) {
    list.innerHTML = '<p class="hint">학생을 선택하세요</p>';
    return;
  }
  
  var history = getSeatHistory(selectedStudent.id);
  if (history.length === 0) {
    list.innerHTML = '<div class="history-item"><p>📭 저장된 기록 없음</p></div>';
    return;
  }
  
  list.innerHTML = '';
  history.forEach(function(h) {
    var item = document.createElement('div');
    item.className = 'history-item';
    var neighborsText = h.neighbors.length > 0 
      ? h.neighbors.map(function(n) { return n.name; }).join(', ')
      : '짝꿍 없음';
    item.innerHTML = '<div><strong>' + formatDateTime(new Date(h.date)) + '</strong> - ' + escapeHtml(h.memo) + '</div>' +
      '<div style="font-size:0.8rem;color:#666;">📍 ' + h.group + '분단 ' + h.row + '행 | 👥 ' + neighborsText + '</div>';
    list.appendChild(item);
  });
}

// ===== 명렬표 =====
function openRosterEditor() {
  closeModal('settingsModal');
  editingClassIndex = currentClassIndex;
  renderRosterClassTabs();
  loadRosterForClass(editingClassIndex);
  openModal('rosterModal');
}

function renderRosterClassTabs() {
  var container = document.getElementById('rosterClassTabs');
  container.innerHTML = '';
  
  for (var i = 0; i < MAX_CLASSES; i++) {
    var tab = document.createElement('button');
    tab.type = 'button';
    tab.className = 'class-tab' + (i === editingClassIndex ? ' active' : '');
    var data = allClassData[i];
    if (data && data.students && data.students.length > 0) {
      tab.classList.add('has-data');
    }
    tab.textContent = data && data.className ? data.className : ((i + 1) + '반');
    
    (function(idx) {
      tab.addEventListener('click', function() {
        saveRosterForClass(editingClassIndex);
        editingClassIndex = idx;
        document.querySelectorAll('.class-tab').forEach(function(t, ti) {
          t.classList.toggle('active', ti === idx);
        });
        loadRosterForClass(idx);
      });
    })(i);
    
    container.appendChild(tab);
  }
}

function loadRosterForClass(classIdx) {
  var data = allClassData[classIdx];
  document.getElementById('rosterClassName').value = data && data.className ? data.className : ((classIdx + 1) + '반');
  var studentList = sortStudentsByName(data && data.students ? data.students : []);
  document.getElementById('rosterInput').value = studentList.map(function(s) {
    return s.id + ' ' + s.name + ' ' + (s.gender || '남');
  }).join('\n');
}

function saveRosterForClass(classIdx) {
  var className = document.getElementById('rosterClassName').value.trim() || ((classIdx + 1) + '반');
  var lines = document.getElementById('rosterInput').value.trim().split('\n').filter(function(l) { return l.trim(); });
  
  var newStudents = lines.map(function(line) {
    var parts = line.trim().split(/\s+/);
    var id = parts[0] || '';
    var gender = parts[parts.length - 1];
    var isGender = gender === '남' || gender === '여';
    return {
      id: id,
      name: isGender ? parts.slice(1, -1).join(' ') || '이름없음' : parts.slice(1).join(' ') || '이름없음',
      gender: isGender ? gender : '남'
    };
  });
  newStudents = sortStudentsByName(newStudents);
  
  if (!allClassData[classIdx]) {
    allClassData[classIdx] = {
      className: className,
      students: [],
      currentSeats: {},
      lockedSeats: {},
      config: JSON.parse(JSON.stringify(config)),
      relations: { close: {}, far: {} },
      studentMemos: {},
      specialNotes: [],
      savedRecords: []
    };
  }
  
  allClassData[classIdx].className = className;
  allClassData[classIdx].students = newStudents;
}

function updateClassName() {
  var className = document.getElementById('rosterClassName').value.trim();
  if (className && allClassData[editingClassIndex]) {
    allClassData[editingClassIndex].className = className;
  }
}

function saveRoster() {
  saveRosterForClass(editingClassIndex);
  saveToLocalStorage();
  loadCurrentClassData();
  initClassDropdown();
  refreshClassExtras();
  closeModal('rosterModal');
  
  if (editingClassIndex === currentClassIndex) {
    currentSeats = {};
    lockedSeats = {};
    renderSeating();
  }
  
  showToast('명렬표 저장 완료');
}

// ===== 저장/기록 =====
function openSaveModal() {
  var now = new Date();
  document.getElementById('saveDateTime').textContent = formatDateTime(now);
  document.getElementById('saveStudentCount').textContent = Object.keys(currentSeats).filter(function(k) {
    return currentSeats[k] && currentSeats[k].studentId;
  }).length;
  document.getElementById('saveMemo').value = '';
  openModal('saveModal');
}

function confirmSave() {
  var memo = document.getElementById('saveMemo').value.trim();
  var now = new Date();
  
  savedRecords.unshift({
    id: generateId(),
    date: now.toISOString(),
    memo: memo || formatDate(now) + ' 배치',
    seats: JSON.parse(JSON.stringify(currentSeats)),
    config: JSON.parse(JSON.stringify(config)),
    lockedSeats: JSON.parse(JSON.stringify(lockedSeats)),
    studentCount: Object.keys(currentSeats).filter(function(k) {
      return currentSeats[k] && currentSeats[k].studentId;
    }).length
  });
  
  saveCurrentClassData();
  closeModal('saveModal');
  showToast('저장 완료');
}

function openHistoryModal() {
  renderHistoryList();
  openModal('historyModal');
}

function renderHistoryList() {
  var list = document.getElementById('historyList');
  
  if (savedRecords.length === 0) {
    list.innerHTML = '<div class="history-item"><p>저장된 기록 없음</p></div>';
    return;
  }
  
  list.innerHTML = '';
  savedRecords.forEach(function(r) {
    var record = document.createElement('div');
    record.className = 'record';
    record.innerHTML = '<div class="record-info">' +
      '<span class="record-date">' + formatDateTime(new Date(r.date)) + '</span>' +
      '<span class="record-memo">' + escapeHtml(r.memo) + '</span>' +
      '<span class="record-count">👥 ' + r.studentCount + '명</span></div>' +
      '<div class="record-actions">' +
      '<button type="button" class="restore-btn" title="이 배치 불러오기">📂 <span>불러오기</span></button>' +
      '<button type="button" class="delete-btn" title="삭제">🗑️ <span>삭제</span></button></div>';
    
    record.querySelector('.restore-btn').addEventListener('click', function() {
      restoreRecord(r.id);
    });
    record.querySelector('.delete-btn').addEventListener('click', function() {
      deleteRecord(r.id);
    });
    
    list.appendChild(record);
  });
}

function restoreRecord(id) {
  var r = savedRecords.find(function(rec) { return rec.id === id; });
  if (r && confirm('이 기록으로 복원하시겠습니까?')) {
    currentSeats = JSON.parse(JSON.stringify(r.seats));
    if (r.config) config = JSON.parse(JSON.stringify(r.config));
    if (r.lockedSeats) lockedSeats = JSON.parse(JSON.stringify(r.lockedSeats));
    saveCurrentClassData();
    renderSeating();
    closeModal('historyModal');
    showToast('저장한 자리배치를 불러왔습니다');
  }
}

function deleteRecord(id) {
  if (confirm('이 기록을 삭제하시겠습니까?')) {
    savedRecords = savedRecords.filter(function(r) { return r.id !== id; });
    saveCurrentClassData();
    renderHistoryList();
    showToast('삭제됨');
  }
}

// ===== 이력 엑셀 다운로드 =====
function downloadHistoryExcel() {
  if (savedRecords.length === 0) {
    showToast('저장된 배치 기록이 없습니다', 'error');
    return;
  }
  
  var className = allClassData[currentClassIndex] && allClassData[currentClassIndex].className 
    ? allClassData[currentClassIndex].className 
    : ((currentClassIndex + 1) + '반');
  
  var data = [['학번', '이름', '성별', '메모']];
  
  savedRecords.forEach(function(r) {
    data[0].push(formatDate(new Date(r.date)) + ' (' + r.memo + ')');
  });
  
  students.forEach(function(student) {
    var row = [student.id, student.name, student.gender, studentMemos[student.id] || ''];
    
    savedRecords.forEach(function(record) {
      var seatInfo = '';
      Object.keys(record.seats).forEach(function(seatId) {
        var seat = record.seats[seatId];
        if (seat && seat.studentId === student.id) {
          var parts = seatId.split('-');
          var rowNum = parseInt(parts[0]) + 1;
          var col = parseInt(parts[1]);
          var layout = getLayoutFromConfig(record.config || config);
          var group = getGroupIndexByCol(col, layout) + 1;
          seatInfo = group + '분단 ' + rowNum + '행';
          
          var neighbors = [];
          [col - 1, col + 1].forEach(function(neighborCol) {
            if (neighborCol < 0) return;
            var neighborSeatId = parts[0] + '-' + neighborCol;
            var neighborSeat = record.seats[neighborSeatId];
            if (neighborSeat && neighborSeat.studentId && neighborSeat.studentId !== student.id) {
              neighbors.push(neighborSeat.name);
            }
          });
          if (neighbors.length > 0) {
            seatInfo += ' (짝: ' + neighbors.join(', ') + ')';
          }
        }
      });
      row.push(seatInfo);
    });
    
    data.push(row);
  });
  
  var ws = XLSX.utils.aoa_to_sheet(data);
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '자리이력');
  XLSX.writeFile(wb, className + '_자리이력_' + formatDate(new Date()) + '.xlsx');
  
  showToast('엑셀 다운로드 완료!');
}

// ===== 출력 =====
function printSeatingChart(mode) {
  var printArea = document.getElementById('printArea');
  var isTeacher = mode === 'teacher';
  var layout = getCurrentLayout();
  var numGroups = layout.length;
  var isGroupPrint = config.seatType === 'group';
  printArea.classList.toggle('print-group-mode', isGroupPrint);
  if (isGroupPrint) printArea.style.setProperty('--print-group-count', String(numGroups));
  else printArea.style.removeProperty('--print-group-count');
  var className = allClassData[currentClassIndex] && allClassData[currentClassIndex].className 
    ? allClassData[currentClassIndex].className 
    : ((currentClassIndex + 1) + '반');
  
  var html = '<div class="print-header"><h1>' + className + ' 자리배치도</h1>' +
    '<p>' + settings.periodStart + ' ~ ' + settings.periodEnd + '</p></div>';
  
  if (isTeacher) {
    html += '<div class="print-seating">';
    for (var r = config.rows - 1; r >= 0; r--) {
      html += '<div class="print-row">';
      for (var gi = numGroups - 1; gi >= 0; gi--) {
        var teacherGroupNo = r * numGroups + gi + 1;
        html += '<div class="print-group' + (isGroupPrint ? ' print-cluster print-cluster-' + layout[gi] : '') + '"' +
          (isGroupPrint ? ' data-group-label="' + teacherGroupNo + '모둠"' : '') + '>';
        var startCol = 0;
        for (var g = 0; g < gi; g++) startCol += layout[g];
        for (var s = layout[gi] - 1; s >= 0; s--) {
          var actualCol = startCol + s;
          var seatId = r + '-' + actualCol;
          html += generatePrintSeat(seatId);
        }
        html += '</div>';
      }
      html += '</div>';
    }
    html += '</div><div class="print-desk">교 탁</div>';
  } else {
    html += '<div class="print-desk">교 탁</div><div class="print-seating">';
    for (var r2 = 0; r2 < config.rows; r2++) {
      html += '<div class="print-row">';
      for (var gi2 = 0; gi2 < numGroups; gi2++) {
        var studentGroupNo = r2 * numGroups + gi2 + 1;
        html += '<div class="print-group' + (isGroupPrint ? ' print-cluster print-cluster-' + layout[gi2] : '') + '"' +
          (isGroupPrint ? ' data-group-label="' + studentGroupNo + '모둠"' : '') + '>';
        var startCol2 = 0;
        for (var g2 = 0; g2 < gi2; g2++) startCol2 += layout[g2];
        for (var s2 = 0; s2 < layout[gi2]; s2++) {
          var actualCol2 = startCol2 + s2;
          var seatId2 = r2 + '-' + actualCol2;
          html += generatePrintSeat(seatId2);
        }
        html += '</div>';
      }
      html += '</div>';
    }
    html += '</div>';
  }
  
  html += '<div class="print-legend">' +
    '<div class="print-legend-item"><div class="print-legend-box male"></div><span>남</span></div>' +
    '<div class="print-legend-item"><div class="print-legend-box female"></div><span>여</span></div></div>' +
    '<div class="print-footer"><span>' + formatDateTime(new Date()) + '</span><span>' + (isTeacher ? '교탁용' : '학생용') + '</span></div>';
  
  printArea.innerHTML = html;
  setTimeout(function() { window.print(); }, 100);
}

function generatePrintSeat(seatId) {
  var seatData = currentSeats[seatId];
  var isLocked = lockedSeats[seatId];
  var student = seatData ? students.find(function(st) { return st.id === seatData.studentId; }) : null;
  
  var seatClass = 'print-seat';
  if (isLocked) seatClass += ' locked';
  if (student) {
    if (student.gender === '남') seatClass += ' male';
    else if (student.gender === '여') seatClass += ' female';
  }
  if (!seatData || !seatData.studentId) seatClass += ' empty';
  
  var html = '<div class="' + seatClass + '">';
  if (seatData && seatData.studentId) {
    html += '<span class="print-seat-id">' + seatData.studentId + '번</span>';
    html += '<span class="print-seat-name">' + seatData.name + '</span>';
  }
  html += '</div>';
  
  return html;
}

// ===== 설정/백업 =====
function openSectionConfig() {
  closeModal('settingsModal');
  showSetupWizard();
}

function openBackupModal() {
  closeModal('settingsModal');
  document.getElementById('exportCode').value = '';
  document.getElementById('importCode').value = '';
  openModal('backupModal');
}

function exportData() {
  saveCurrentClassData();
  var obj = {
    version: APP_VERSION,
    exportDate: new Date().toISOString(),
    settings: settings,
    allClassData: allClassData,
    currentClassIndex: currentClassIndex,
    customFavorites: CustomFavorites.getData(),
    calendarEvents: ScheduleCalendar.getData(),
    calendarProperties: ScheduleCalendar.getProperties()
  };
  document.getElementById('exportCode').value = 'SEAT_V7_' + btoa(unescape(encodeURIComponent(JSON.stringify(obj))));
  showToast('코드 생성 완료');
}

function copyExportCode() {
  var code = document.getElementById('exportCode').value;
  if (!code) {
    showToast('먼저 코드 생성', 'error');
    return;
  }
  navigator.clipboard.writeText(code).then(function() {
    showToast('복사됨');
  }).catch(function() {
    document.getElementById('exportCode').select();
    document.execCommand('copy');
    showToast('복사됨');
  });
}

function importData() {
  var code = document.getElementById('importCode').value.trim();
  if (!code) {
    showToast('코드 입력', 'error');
    return;
  }
  
  try {
    if (code.indexOf('SEAT_V') !== 0) throw new Error('유효하지 않음');
    var obj = JSON.parse(decodeURIComponent(escape(atob(code.replace(/^SEAT_V\d+_/, '')))));
    
    if (confirm('데이터를 복원하시겠습니까?')) {
      settings = obj.settings || settings;
      allClassData = obj.allClassData || {};
      currentClassIndex = obj.currentClassIndex || 0;
      CustomFavorites.setData(obj.customFavorites || []);
      ScheduleCalendar.setData(obj.calendarEvents || []);
      ScheduleCalendar.setProperties(obj.calendarProperties || []);
      initAllClassData();
      loadCurrentClassData();
      saveToLocalStorage();
      initClassDropdown();
      closeModal('backupModal');
      renderSeating();
      showToast('저장한 자리배치를 불러왔습니다');
    }
  } catch (e) {
    showToast('복원 실패', 'error');
  }
}

function confirmReset() {
  var input = prompt('"초기화" 입력:');
  if (input === '초기화') {
    localStorage.removeItem(STORAGE_KEY);
    localStorage.removeItem('seatCustomFavorites_v1');
    localStorage.removeItem('seatCalendar_v1');
    localStorage.removeItem('seatCalendarProperties_v1');
    localStorage.removeItem('seatCalendarProperties_v2');
    localStorage.removeItem('tongSchoolProfile_v1');
    localStorage.removeItem('tongSchoolTodayCache_v1');
    localStorage.removeItem('tongGoogleCalendarConfig_v1');
    localStorage.removeItem('tongGoogleCalendarSync_v1');
    location.reload();
  }
}

// ===== 유틸리티 =====
// 헤더 탭에서 화면 가득 여는 창들 — 열려 있는 동안에는 뒤의 홈 대시보드를 숨기고 해당 탭을 강조합니다.
var HEADER_PAGE_MODALS = {
  teacherScheduleModal: 'topScheduleBtn',
  calendarModal: 'calendarBtn',
  sheetMemoModal: 'topSheetMemoBtn'
};

function syncHeaderPageState() {
  var openId = '';
  Object.keys(HEADER_PAGE_MODALS).forEach(function (id) {
    var modal = document.getElementById(id);
    if (modal && modal.classList.contains('show')) openId = id;
  });
  document.body.classList.toggle('header-page-open', Boolean(openId));
  Object.keys(HEADER_PAGE_MODALS).forEach(function (id) {
    var tab = document.getElementById(HEADER_PAGE_MODALS[id]);
    if (tab) tab.classList.toggle('active', id === openId);
  });
  // 창을 닫으면 원래 보고 있던 화면의 탭 강조가 돌아옵니다.
  var baseTab = appMode === 'home'
    ? document.getElementById('topDashboardBtn')
    : document.querySelector('.mode-tab[data-mode="' + appMode + '"]');
  if (baseTab) baseTab.classList.toggle('active', !openId);
}

function openModal(id) {
  document.getElementById(id).classList.add('show');
  if (HEADER_PAGE_MODALS[id]) syncHeaderPageState();
}

function closeModal(id) {
  document.getElementById(id).classList.remove('show');
  if (HEADER_PAGE_MODALS[id]) syncHeaderPageState();
}

function saveToLocalStorage() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({
    settings: settings,
    allClassData: allClassData,
    currentClassIndex: currentClassIndex
  }));
  if (typeof CloudSheetSync !== 'undefined' && CloudSheetSync.getAuth() && CloudSheetSync.getSheetUrl()) {
    if (saveToLocalStorage._timer) clearTimeout(saveToLocalStorage._timer);
    saveToLocalStorage._timer = setTimeout(function () {
      CloudSheetSync.syncNow(false);
    }, 4000);
  }
}

function loadFromLocalStorage() {
  var saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    try {
      var data = JSON.parse(saved);
      settings = data.settings || settings;
      allClassData = data.allClassData || {};
      currentClassIndex = data.currentClassIndex || 0;
      initAllClassData();
      loadCurrentClassData();
    } catch (e) {
      console.error('Load failed:', e);
    }
  }
}

function showToast(msg, type, actionLabel, action, duration) {
  var toast = document.getElementById('toast');
  if (!toast) return;
  clearTimeout(toastHideTimer);
  toast.textContent = '';
  var text = document.createElement('span');
  text.textContent = msg;
  toast.appendChild(text);
  if (actionLabel && typeof action === 'function') {
    var button = document.createElement('button');
    button.type = 'button';
    button.className = 'toast-action';
    button.textContent = actionLabel;
    button.addEventListener('click', function() {
      clearTimeout(toastHideTimer);
      toast.classList.remove('show');
      action();
    });
    toast.appendChild(button);
  }
  toast.className = 'toast show ' + (type || '');
  toastHideTimer = setTimeout(function() {
    toast.classList.remove('show');
  }, duration || (actionLabel ? 5200 : 2500));
}

function formatDate(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

function formatDateTime(d) {
  return formatDate(d) + ' ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
}

function formatDateForInput(d) {
  return d.toISOString().split('T')[0];
}

function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

function escapeHtml(str) {
  var div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

// ===================================================================
// 사운드 — Web Audio API로 즉석 합성합니다.
// 음원 파일이나 iframe을 불러오지 않으므로 '추첨 시작'과 동시에 소리가 납니다.
// ===================================================================
var Sound = (function() {
  var ctx = null;
  var master = null;
  var roll = null;
  var enabled = true;

  try {
    enabled = localStorage.getItem(SOUND_PREF_KEY) !== 'off';
  } catch (e) { /* 저장소 접근 불가 시 기본값 사용 */ }

  function init() {
    if (!enabled) return null;
    try {
      if (!ctx) {
        var AC = window.AudioContext || window.webkitAudioContext;
        if (!AC) { enabled = false; return null; }
        ctx = new AC();
        master = ctx.createGain();
        master.gain.value = 0.5;
        master.connect(ctx.destination);
      }
      if (ctx.state === 'suspended' && ctx.resume) ctx.resume();
      return ctx;
    } catch (e) {
      enabled = false;
      return null;
    }
  }

  function noiseBuffer(seconds) {
    var len = Math.floor(ctx.sampleRate * seconds);
    var buf = ctx.createBuffer(1, len, ctx.sampleRate);
    var data = buf.getChannelData(0);
    for (var i = 0; i < len; i++) data[i] = Math.random() * 2 - 1;
    return buf;
  }

  // 두구두구… 드럼롤
  function startRoll() {
    if (!init() || roll) return;
    try {
      var src = ctx.createBufferSource();
      src.buffer = noiseBuffer(1);
      src.loop = true;

      var band = ctx.createBiquadFilter();
      band.type = 'bandpass';
      band.frequency.value = 160;
      band.Q.value = 1.1;

      var gain = ctx.createGain();
      gain.gain.value = 0.10;

      // 트레몰로(떨림)로 드럼롤 느낌을 만듭니다
      var lfo = ctx.createOscillator();
      lfo.type = 'square';
      lfo.frequency.value = 15;
      var lfoGain = ctx.createGain();
      lfoGain.gain.value = 0.075;
      lfo.connect(lfoGain);
      lfoGain.connect(gain.gain);

      src.connect(band);
      band.connect(gain);
      gain.connect(master);
      src.start();
      lfo.start();

      roll = { src: src, gain: gain, lfo: lfo };
    } catch (e) { roll = null; }
  }

  function setTension(progress, climax) {
    if (!roll || !ctx) return;
    try {
      var t = ctx.currentTime;
      roll.lfo.frequency.linearRampToValueAtTime(14 + progress * 20 + (climax ? 12 : 0), t + 0.25);
      roll.gain.gain.linearRampToValueAtTime(0.09 + progress * 0.07, t + 0.25);
    } catch (e) { /* 무시 */ }
  }

  function stopRoll() {
    if (!roll || !ctx) return;
    try {
      roll.gain.gain.linearRampToValueAtTime(0.0001, ctx.currentTime + 0.15);
      roll.src.stop(ctx.currentTime + 0.2);
      roll.lfo.stop(ctx.currentTime + 0.2);
    } catch (e) { /* 무시 */ }
    roll = null;
  }

  function tone(freq, start, dur, type, peak, sweepTo) {
    var osc = ctx.createOscillator();
    var g = ctx.createGain();
    osc.type = type || 'sine';
    osc.frequency.setValueAtTime(freq, start);
    if (sweepTo) osc.frequency.exponentialRampToValueAtTime(sweepTo, start + dur);
    g.gain.setValueAtTime(0.0001, start);
    g.gain.exponentialRampToValueAtTime(peak, start + 0.012);
    g.gain.exponentialRampToValueAtTime(0.0001, start + dur);
    osc.connect(g);
    g.connect(master);
    osc.start(start);
    osc.stop(start + dur + 0.02);
  }

  // 한 명 확정될 때마다 음이 한 칸씩 올라갑니다 (5음 음계)
  var PENTA = [0, 2, 4, 7, 9];
  function playPick(index, total, isFinal) {
    if (!init()) return;
    try {
      var t = ctx.currentTime;
      if (isFinal) {
        tone(523.25, t, 0.5, 'triangle', 0.30, 1046.5);
        tone(784.0, t + 0.06, 0.5, 'sine', 0.22);
        return;
      }
      var step = PENTA[index % PENTA.length] + 12 * Math.floor((index % (PENTA.length * 2)) / PENTA.length);
      var freq = 587.33 * Math.pow(2, step / 12);
      tone(freq, t, 0.16, 'triangle', 0.24);
      tone(freq * 2, t + 0.01, 0.10, 'sine', 0.10);
    } catch (e) { /* 무시 */ }
  }

  function playFanfare() {
    if (!init()) return;
    try {
      var t = ctx.currentTime + 0.05;
      var notes = [523.25, 659.25, 783.99, 1046.5];
      notes.forEach(function(f, i) {
        tone(f, t + i * 0.11, 0.42, 'triangle', 0.26);
      });
      tone(1318.5, t + 0.46, 0.75, 'sine', 0.22);

      // 박수 느낌의 노이즈 스웰
      var src = ctx.createBufferSource();
      src.buffer = noiseBuffer(1.4);
      var hp = ctx.createBiquadFilter();
      hp.type = 'highpass';
      hp.frequency.value = 1400;
      var g = ctx.createGain();
      g.gain.setValueAtTime(0.0001, t);
      g.gain.exponentialRampToValueAtTime(0.13, t + 0.22);
      g.gain.exponentialRampToValueAtTime(0.0001, t + 1.3);
      src.connect(hp); hp.connect(g); g.connect(master);
      src.start(t);
      src.stop(t + 1.4);
    } catch (e) { /* 무시 */ }
  }

  // 룰렛이 돌아가는 소리 — 점점 낮고 느려지는 느낌
  function playSpin(i, total) {
    if (!init()) return;
    try {
      var t = ctx.currentTime;
      var p = i / Math.max(1, total);
      tone(1400 - p * 500, t, 0.05, 'square', 0.09);
    } catch (e) { /* 무시 */ }
  }

  // 룰렛 당첨 — 팡파레보다 짧고 경쾌하게
  function playWin() {
    if (!init()) return;
    try {
      var t = ctx.currentTime;
      [523.25, 659.25, 783.99, 1046.5, 1318.5].forEach(function (f, i) {
        tone(f, t + i * 0.07, 0.3, 'triangle', 0.26);
      });
    } catch (e) { /* 무시 */ }
  }

  // 타이머 마지막 10초 똑딱임
  function playTick() {
    if (!init()) return;
    try { tone(1100, ctx.currentTime, 0.06, 'square', 0.10); } catch (e) { /* 무시 */ }
  }

  // 타이머 종료 알람
  function playAlarm() {
    if (!init()) return;
    try {
      var t = ctx.currentTime;
      for (var i = 0; i < 3; i++) {
        tone(880, t + i * 0.32, 0.16, 'square', 0.26);
        tone(1320, t + i * 0.32 + 0.14, 0.16, 'square', 0.22);
      }
    } catch (e) { /* 무시 */ }
  }

  // 추첨 스플래시 카운트다운 (n=0 이면 출발 신호)
  function playCountdown(n) {
    if (!init()) return;
    try {
      var t = ctx.currentTime;
      if (n === 0) {
        tone(523.25, t, 0.5, 'triangle', 0.30, 1567.98);
        tone(1046.5, t + 0.05, 0.45, 'sine', 0.22);
      } else {
        tone(392 + (3 - n) * 90, t, 0.22, 'triangle', 0.26);
      }
    } catch (e) { /* 무시 */ }
  }

  function toggle() {
    enabled = !enabled;
    try { localStorage.setItem(SOUND_PREF_KEY, enabled ? 'on' : 'off'); } catch (e) { /* 무시 */ }
    if (!enabled) stopRoll();
    return enabled;
  }

  function isEnabled() { return enabled; }

  return {
    init: init, startRoll: startRoll, stopRoll: stopRoll, setTension: setTension,
    playPick: playPick, playFanfare: playFanfare, playTick: playTick,
    playSpin: playSpin, playWin: playWin,
    playAlarm: playAlarm, playCountdown: playCountdown,
    toggle: toggle, isEnabled: isEnabled
  };
})();

// ===================================================================
// 색종이(컨페티) — 외부 라이브러리 없이 canvas로 직접 그립니다
// ===================================================================
var Confetti = (function() {
  var canvas = null, ctx2d = null, particles = [], rafId = null;
  var COLORS = ['#667eea', '#764ba2', '#f093fb', '#ffd54f', '#4ade80', '#38bdf8', '#fb7185'];
  var reduced = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  function ensure() {
    if (!canvas) {
      canvas = document.getElementById('confettiCanvas');
      if (!canvas) return null;
      ctx2d = canvas.getContext('2d');
      window.addEventListener('resize', resize);
    }
    resize();
    return ctx2d;
  }

  function resize() {
    if (!canvas) return;
    var dpr = window.devicePixelRatio || 1;
    canvas.width = window.innerWidth * dpr;
    canvas.height = window.innerHeight * dpr;
    ctx2d.setTransform(dpr, 0, 0, dpr, 0, 0);
  }

  function spawn(x, y, count, power) {
    for (var i = 0; i < count; i++) {
      var angle = Math.random() * Math.PI * 2;
      var speed = (0.5 + Math.random()) * power;
      particles.push({
        x: x, y: y,
        vx: Math.cos(angle) * speed,
        vy: Math.sin(angle) * speed - power * 0.6,
        size: 4 + Math.random() * 5,
        color: COLORS[(Math.random() * COLORS.length) | 0],
        rot: Math.random() * Math.PI,
        vr: (Math.random() - 0.5) * 0.28,
        life: 1
      });
    }
    start();
  }

  function start() {
    if (rafId || !canvas) return;
    canvas.classList.add('show');
    rafId = requestAnimationFrame(tick);
  }

  function tick() {
    ctx2d.clearRect(0, 0, canvas.width, canvas.height);
    for (var i = particles.length - 1; i >= 0; i--) {
      var p = particles[i];
      p.vy += 0.22;          // 중력
      p.vx *= 0.992;
      p.x += p.vx;
      p.y += p.vy;
      p.rot += p.vr;
      p.life -= 0.008;
      if (p.life <= 0 || p.y > window.innerHeight + 40) {
        particles.splice(i, 1);
        continue;
      }
      ctx2d.save();
      ctx2d.translate(p.x, p.y);
      ctx2d.rotate(p.rot);
      ctx2d.globalAlpha = Math.max(0, Math.min(1, p.life));
      ctx2d.fillStyle = p.color;
      ctx2d.fillRect(-p.size / 2, -p.size / 2, p.size, p.size * 0.62);
      ctx2d.restore();
    }
    if (particles.length) {
      rafId = requestAnimationFrame(tick);
    } else {
      ctx2d.clearRect(0, 0, canvas.width, canvas.height);
      canvas.classList.remove('show');
      rafId = null;
    }
  }

  function burstAt(el, count) {
    if (reduced || !ensure() || !el) return;
    var r = el.getBoundingClientRect();
    spawn(r.left + r.width / 2, r.top + r.height / 2, count, 5);
  }

  function celebrate() {
    if (reduced || !ensure()) return;
    var w = window.innerWidth;
    spawn(w * 0.2, window.innerHeight * 0.34, 60, 9);
    spawn(w * 0.5, window.innerHeight * 0.26, 80, 10);
    spawn(w * 0.8, window.innerHeight * 0.34, 60, 9);
  }

  return { burstAt: burstAt, celebrate: celebrate };
})();

// ===================================================================
// 사용 설명서 캐러셀
// ===================================================================
var GUIDE_SLIDES = [
  // ===== 1. 시작 · 1~7번 전체 구조 =====
  {
    section: 'start',
    img: 'guide/v13_00_dashboard.png',
    title: '처음 시작 1단계 — 1·2번 초기 설정과 명렬표 · 우측 하단 버튼',
    spots: [
      { num: 1, target: [11.7, 1.3, 10.7, 4.1], card: [4, 18], title: '1. Google 시트 연동', text: '마스터 시트 URL과 Google 계정·캘린더를 한 번에 연결합니다.', when: '가장 먼저 1회 설정' },
      { num: 2, target: [22.6, 1.3, 9.9, 4.1], card: [33, 18], title: '2. 학교·수업시간표', text: 'NEIS 학교 검색으로 내 수업 시간표와 학사일정을 자동 불러옵니다.', when: '학기 초·시간표 변경 시' },
      { num: 3, target: [80.6, 1.3, 9.3, 4.2], card: [66, 18], title: '학급 명렬표 등록', text: '직접 입력·엑셀 업로드·마스터 시트 연동으로 학생 명단을 올립니다.', when: '자리·출결 사용 전 필수' },
      { num: 4, target: [93.8, 89.2, 4.9, 9.2], card: [64, 68], title: '우측 하단 [오늘] 버튼', text: '오늘 브리핑을 빠르게 열고, 내 사진으로 프로필 아이콘을 바꿀 수 있습니다.', when: '나만의 교무수첩 꾸미기' }
    ],
    callout: { type: 'good', text: '1번(시트 연동) → 2번(시간표) → 명렬표 등록 순서로 1회만 설정해 두면 됩니다.' }
  },
  {
    section: 'start',
    img: 'guide/v13_00_dashboard.png',
    title: '매일 쓰는 교무 업무 — 상단 3~7번 핵심 메뉴 한눈에 보기',
    spots: [
      { num: 3, target: [32.8, 1.3, 9.7, 4.1], card: [4, 22], title: '3. 학급 관리(자리)', text: '자리 배치, 추첨, 교탁 시점, 자리 클릭 출결·관찰을 실행합니다.', when: '수업 중·자리 교체 시' },
      { num: 4, target: [42.8, 1.3, 9.3, 4.1], card: [34, 22], title: '4. 수업·출결 기록', text: '출석 체크, 칭찬·세특 관찰 기록, 월간 출결 엑셀을 관리합니다.', when: '매일 조회·종례·수업 후' },
      { num: 5, target: [52.4, 1.3, 9.0, 4.1], card: [64, 22], title: '5. 일정·학사달력', text: 'NEIS 학사일정, 학생 상담 기록, Google 캘린더를 동기화합니다.', when: '주간·월간 일정 관리' },
      { num: 6, target: [61.7, 1.3, 18.1, 4.1], card: [64, 52], title: '6. 교실 도구 · 7. 시트 모음', text: '타이머·칭찬·1인1역·알림장·학습앱·웹툰과 마스터 시트를 엽니다.', when: '수업 활동·자료 확인 시' }
    ],
    callout: { type: 'good', text: '필요한 번호만 눌러 바로 이동할 수 있습니다. 다음 장부터 각 메뉴 상세 사용법을 안내합니다.' }
  },
  {
    section: 'start',
    img: 'guide/v11_01_sync_modal.png',
    title: '1번 메뉴 · Google 시트 연동 — 중앙 워크플로우 창에서 3단계로 연결',
    desc: '<ul>' +
      '<li><strong>왼쪽 3단계 따라하기</strong>① 구글 드라이브 사본 복사(권한 오류 시 표준 엑셀 다운로드) → ② Apps Script 새 배포(/exec 복사) → ③ 오른쪽 칸에 URL 저장<span class="gd-when">활용: PC·태블릿 간 데이터 동기화</span></li>' +
      '<li><strong>화면으로 보는 단계별 안내</strong>하단 1~4번 탭을 누르면 사본 복사부터 웹앱 URL 복사까지 그림으로 바로 확인할 수 있습니다.</li>' +
      '<li><strong>오른쪽 통합 로그인</strong>[Google 통합 로그인]을 한 번 누르면 교무수첩 계정과 Google 캘린더가 동시에 연결됩니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '시트 사본을 만든 직후, 학생 자료를 입력하기 전에 먼저 Apps Script URL을 저장해 두면 이후 기록이 안전하게 쌓입니다.' }
  },
  {
    section: 'start',
    img: 'guide/tong_sync_step1_numbered.png',
    title: 'STEP 1 · 시트 1:1 탭 매핑 & 기존 데이터 보호 원칙',
    desc: '<ul>' +
      '<li><strong>① 기존 데이터 안전 보호</strong>현재 기기에 입력한 명렬표·출결·좌석 기록은 지워지지 않고 시트와 안전하게 병합됩니다.</li>' +
      '<li><strong>② 시트 탭 1:1 자동 저장</strong>수업관찰 → [6_수업관찰_기록] / AI 세특 → [7_AI세특_작성] / 출결·칭찬·1인1역 → [8_출결_칭찬_기록]<span class="gd-when">활용: 나이스(NEIS-Help)와 자동 연계</span></li>' +
      '<li><strong>③ 태블릿 ↔ 노트북 이어쓰기</strong>두 기기에 같은 웹앱 URL을 저장한 뒤 [시트로 보내기]와 [시트에서 가져오기]를 사용합니다.</li>' +
      '</ul>',
    callout: { type: 'warn', text: '두 기기에서 함께 쓸 때는 반 이름과 명렬표(학번)를 동일하게 맞춰야 학생 기록이 정확히 연결됩니다.' }
  },
  {
    section: 'start',
    img: 'guide/v11_02_schedule_modal.png',
    title: '2번 메뉴 · 학교·수업시간표 — NEIS API로 학교와 내 시간표 완성',
    desc: '<ul>' +
      '<li><strong>우리 학교 검색</strong>학교명을 검색해 고르면 학년·학급과 이번 주 NEIS 시간표가 자동으로 불러와집니다.</li>' +
      '<li><strong>담당 수업 추가</strong>학년·학급·과목을 골라 [+ 추가]를 누르면 요일·교시별 내 교무 시간표가 완성됩니다.<span class="gd-when">활용: 담임·교과전담·비담임 모두 지원</span></li>' +
      '<li><strong>출석부·달력 원클릭 반영</strong>선택한 학급을 출석부 반 목록에 넣고, NEIS 학사일정도 달력에 바로 채울 수 있습니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '교과전담 선생님은 담임 학급을 [담임 없음]으로 두고 담당하는 여러 반의 수업만 골라 추가하면 됩니다.' }
  },
  {
    section: 'start',
    img: 'guide/v11_03_roster_modal.png',
    title: '학급 명렬표 등록 — 직접 입력 · 엑셀/시트 업로드 · NEIS-Help 자동 연동',
    desc: '<ul>' +
      '<li><strong>방법 1. 직접 입력</strong>소수 학생 수정이나 빠른 테스트 시 [학급 명렬표 등록]에서 <code>번호 이름 성별</code> 한 줄씩 입력합니다.<span class="gd-when">활용: 전출입 1~2명 수정 시</span></li>' +
      '<li><strong>방법 2. 엑셀 / 시트 업로드</strong>우측 패널 [명렬표 업로드]를 통해 나이스에서 내려받은 학급 엑셀/시트 파일을 그대로 올릴 수 있습니다.<span class="gd-when">활용: 20~35명 전체 반 최초 등록</span></li>' +
      '<li><strong>방법 3. Google 시트·NEIS-Help 연동</strong>마스터 시트 [2_학생명렬표]에 올려두면 NEIS-Help와 T-ong이 같은 명렬표를 자동으로 공유합니다.<span class="gd-when">활용: 생기부·교무수첩 동시 사용자 추천</span></li>' +
      '</ul>',
    callout: { type: 'good', text: '성별(남/여)을 함께 넣어두면 자리배치에서 같은 성별 우선·남녀 짝꿍 옵션을 바로 쓸 수 있습니다.' }
  },

  // ===== 2. 3번 메뉴: 자리 배치 (앱의 출발점) =====
  {
    section: 'seat',
    img: 'guide/v13_04_seating_main.png',
    title: '3번 메뉴 · 자리 배치 ① — 배치 전 준비와 추첨 (1~4단계)',
    spots: [
      { num: 1, target: [84.0, 15.1, 13.6, 6.8], card: [48, 13], title: '배치기간 설정', text: '이번 자리 배치의 시작일과 종료일을 지정해 이력으로 남깁니다.', when: '새 달 자리 바꿀 때' },
      { num: 2, target: [84.0, 23.0, 13.6, 5.0], card: [48, 29], title: '명렬표 업로드', text: '엑셀/시트 파일이나 텍스트로 반 명렬표를 불러옵니다.', when: '학생 명단 최초·수정 시' },
      { num: 3, target: [84.0, 29.3, 13.6, 5.0], card: [48, 45], title: '좌석 설정', text: '분단·모둠(4/6인)·학번순(가로/세로)·성별 우선 조건을 고릅니다.', when: '교실 책상 구조 맞출 때' },
      { num: 4, target: [84.0, 35.6, 13.6, 5.0], card: [48, 62], title: '추첨 시작', text: '카운트다운과 함께 공정하게 자리를 추첨해 배정합니다.', when: '학생들과 함께 추첨할 때' }
    ],
    callout: { type: 'good', text: '자리 배치는 T-ong의 출발점입니다. 오른쪽 위에서 아래 순서대로 누르면 배치가 끝납니다.' }
  },
  {
    section: 'seat',
    img: 'guide/v13_04_seating_main.png',
    title: '3번 메뉴 · 자리 배치 ② — 저장 · 교탁 시점 · 출력 · 좌측 메뉴',
    spots: [
      { num: 1, target: [84.0, 41.3, 13.6, 10.7], card: [48, 24], title: '저장 · 배치 불러오기', text: '현재 자리를 저장하고, 이전 달 배치로 언제든 되돌립니다.', when: '직전 짝꿍 중복 확인 시' },
      { num: 2, target: [84.0, 52.8, 13.6, 5.0], card: [48, 44], title: '교탁 시점 (180도 반전)', text: '인쇄 없이 선생님 눈높이 기준으로 좌석표를 뒤집어 봅니다.', when: '수업 중 교탁에서 이름 확인할 때' },
      { num: 3, target: [84.0, 58.5, 13.6, 10.7], card: [48, 63], title: '데이터 백업 · 출력', text: '교실 게시용/교탁용 좌석표 인쇄 및 백업 파일을 저장합니다.', when: '학급 게시·PC 이전 시' },
      { num: 4, target: [1.3, 8.6, 3.8, 80.4], card: [9, 28], title: '좌측 빠른 아이콘 메뉴', text: '학생 메모·교우관계(붙이기/떼기)·배치 이력·AI 비서를 엽니다.', when: '배려 학생 조건 지정 시' }
    ],
    callout: { type: 'good', text: '배치 후에는 학생 카드를 클릭해 출결·칭찬·관찰 기록을 바로 남길 수 있습니다.' }
  },
  {
    section: 'seat',
    img: 'guide/v11_05_seating_arranged.png',
    title: '좌석 배치 방식 — 상황별 추천 조건 (학번순 · 성별 우선 · 모둠 · 수동 교환)',
    desc: '<ul>' +
      '<li><strong>학번 가로순 · 세로순</strong>교사가 학생 이름과 얼굴을 가장 빨리 외워야 하는 시기에 쓰기 좋습니다.<span class="gd-when">활용: 새 학기 첫 주 · 교과 전담 첫 수업</span></li>' +
      '<li><strong>같은 성별 우선 · 남녀 짝꿍 배치</strong>학급 분위기나 활동 성격에 맞춰 짝꿍 성별 규칙을 자동으로 반영해 추첨합니다.<span class="gd-when">활용: 정기 자리 교체 · 짝 활동</span></li>' +
      '<li><strong>4인·6인 모둠 배치 & 드래그 맞교환</strong>책상을 모둠형으로 묶어 배정하고, 배치 후에도 두 학생을 끌어 자리를 맞바꿀 수 있습니다.<span class="gd-when">활용: 토의·협동학습 · 미세 조정</span></li>' +
      '</ul>',
    callout: { type: 'good', text: '빈 통로나 사용하지 않는 책상은 좌석 설정에서 클릭해 비워두면 실제 교실 구조와 똑같이 맞출 수 있습니다.' }
  },
  {
    section: 'seat',
    img: 'guide/v11_06_seating_teacher_view.png',
    title: '교탁 시점 — 인쇄하지 않고 교탁 앞 내 시야로 화면을 뒤집습니다',
    desc: '<ul>' +
      '<li><strong>교탁 시점 원클릭 전환</strong>우측 [7 교탁 시점] 버튼을 누르면 교탁이 아래로 내려오고 좌석표 전체가 180도 반전됩니다.<span class="gd-when">활용: 수업 중 교탁에서 학생 파악</span></li>' +
      '<li><strong>학생 시점과 즉시 전환</strong>학생들에게 프로젝터로 보여줄 때는 한 번 더 눌러 칠판 기준 정방향으로 되돌립니다.</li>' +
      '<li><strong>반전 상태에서도 기록 가능</strong>교탁 시점 그대로 학생 카드를 눌러 출결·칭찬·관찰 메모를 남길 수 있습니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '종이 좌석표를 매달 출력해 교탁에 붙이지 않아도 태블릿/PC에서 바로 학생 이름을 확인할 수 있습니다.' }
  },
  {
    section: 'seat',
    img: 'guide/v67-seat-attendance.png',
    title: '자리에서 학생 클릭 — 수업 중 출결·칭찬·1인1역·관찰을 즉시 기록',
    desc: '<ul>' +
      '<li><strong>좌석 카드 클릭 팝업</strong>앉아 있는 학생 카드를 누르면 해당 학생의 종합 기록 창이 열립니다.<span class="gd-when">활용: 수업 시작 1분 출석 체크</span></li>' +
      '<li><strong>출결·칭찬 원클릭 등록</strong>결석·지각·조퇴를 체크하면 4번 [수업·출결 기록] 탭과 즉시 동기화됩니다.</li>' +
      '<li><strong>수업 관찰 메모 누적</strong>발표·질문·모둠 기여도 메모를 남기면 Google 마스터 시트 [6_수업관찰_기록]으로 이어집니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '빈자리를 보며 학생을 바로 클릭하면 출석부와 좌석표를 번갈아 찾을 필요가 없습니다.' }
  },
  {
    section: 'seat',
    img: 'guide/07-relations.png',
    title: '교우 관계 · 자리 고정 · 학생 메모 — 배려 학생을 미리 설정하고 추첨',
    desc: '<ul>' +
      '<li><strong>붙일 학생 · 떼어놓을 학생</strong>수업 집중과 교우 관계를 위해 함께 앉히거나 떨어뜨릴 학생 쌍을 미리 등록합니다.<span class="gd-when">활용: 생활지도·갈등 예방</span></li>' +
      '<li><strong>특정 자리 고정(자물쇠)</strong>시력·청력이나 건강상 앞자리·창가가 필요한 학생은 자물쇠로 고정한 뒤 나머지만 추첨합니다.<span class="gd-when">활용: 신체적 배려·학습 지원</span></li>' +
      '<li><strong>학생별 메모</strong>상담 특이사항이나 관찰 포인트를 적어두면 좌석과 출결 카드에서 함께 확인됩니다.</li>' +
      '</ul>',
    callout: { type: 'warn', text: '고정석과 떼어놓기 조건이 너무 많으면 배치 조합이 줄어드므로 꼭 필요한 학생만 지정하세요.' }
  },
  {
    section: 'seat',
    img: 'guide/v67-history-load.png',
    title: '배치 저장 · 불러오기 · 인쇄 · 누가기록 엑셀 다운로드',
    desc: '<ul>' +
      '<li><strong>배치 저장 & 이전 배치 복원</strong>[5 저장]으로 날짜별 배치를 기록해 두면 [6 배치 불러오기]로 과거 자리로 바로 복원합니다.<span class="gd-when">활용: 직전 짝꿍 중복 방지</span></li>' +
      '<li><strong>교실 게시용 / 교탁용 출력</strong>[8 출력]에서 학생용(칠판 기준)과 교사용(교탁 기준) 인쇄 화면을 선택할 수 있습니다.</li>' +
      '<li><strong>누가기록 엑셀 내려받기</strong>좌측 메뉴에서 학생별 자리 이력·칭찬·메모를 엑셀 파일로 한 번에 저장합니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '새 자리를 배정할 때마다 [저장]을 눌러두면 학기 말 학생 활동 이력 정리에 큰 도움이 됩니다.' }
  },

  // ===== 3. 4번 메뉴: 수업·출결 기록 =====
  {
    section: 'attend',
    img: 'guide/v11_07_attend_main.png',
    title: '4번 메뉴 · 수업·출결 기록 — 반별 출결 현황과 일괄 입력을 한 화면에서',
    desc: '<ul>' +
      '<li><strong>오늘 전원 출석 & 변동만 선택</strong>[오늘 전원 출석]을 누른 뒤 결석·지각·조퇴 학생만 드롭다운에서 상태와 사유를 선택합니다.<span class="gd-when">활용: 매일 아침 조·종례 30초 마감</span></li>' +
      '<li><strong>상단 실시간 인원 집계</strong>출석·결석·지각·조퇴·결과 인원수가 상단에 자동 계산됩니다.</li>' +
      '<li><strong>구글 시트 실시간 연동</strong>시트가 연결되어 있으면 변경 내용이 마스터 시트 [8_출결_칭찬_기록]에 자동 저장됩니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '다른 날짜의 출결을 수정할 때는 상단 날짜 좌우 화살표(< >)를 눌러 이전 일자로 이동하세요.' }
  },
  {
    section: 'attend',
    img: 'guide/v67-student-record.png',
    title: '출결 학생 클릭 팝업 — 칭찬·수업 관찰·AI 세특 초안까지 연결',
    desc: '<ul>' +
      '<li><strong>학생 이름 클릭</strong>출결 카드의 학생 이름을 누르면 출결·칭찬·1인1역·관찰 메모를 한 번에 입력하는 팝업이 뜹니다.<span class="gd-when">활용: 수업 직후 1줄 관찰 기록</span></li>' +
      '<li><strong>세특·행특 누가기록 연계</strong>여기에 적은 관찰 키워드는 마스터 시트 [6_수업관찰_기록]과 [7_AI세특_작성]으로 바로 이어집니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '수업이 끝난 직후 기억이 생생할 때 키워드 1~2줄만 남겨두면 학기 말 세특 작성이 훨씬 편해집니다.' }
  },
  {
    section: 'attend',
    img: 'guide/20A-attend-excel.png',
    title: '출결 엑셀 · 구글시트 복사 — 월간 출결 통계와 백업 관리',
    desc: '<ul>' +
      '<li><strong>출결 엑셀 다운로드</strong>선택한 기간의 일자별 출결 현황과 학생별 합계 시트를 엑셀(.xlsx)로 내려받습니다.<span class="gd-when">활용: 월말 나이스 출결 대조</span></li>' +
      '<li><strong>구글시트로 복사</strong>표 데이터를 복사해 기존에 쓰던 학급 관리 시트에 그대로 붙여넣을 수 있습니다.</li>' +
      '<li><strong>다운로드 후 초기화</strong>학기 말 백업 엑셀을 먼저 저장한 뒤 새 학기 출결을 깔끔하게 시작합니다.</li>' +
      '</ul>',
    callout: { type: 'warn', text: '초기화 전에는 반드시 [출결 엑셀]이나 [데이터 백업]으로 사본을 먼저 보관하세요.' }
  },

  // ===== 4. 5번 메뉴: 일정·학사달력 =====
  {
    section: 'calendar',
    img: 'guide/v13_08_calendar_main.png',
    title: '5번 메뉴 · 일정·학사달력 — 핵심 4개 버튼으로 일정 완벽 관리',
    spots: [
      { num: 1, target: [41.9, 2.4, 6.6, 3.8], card: [8, 20], title: '+ 일정 추가', text: '학사·수업·학급·상담·개인 등 속성별로 일정을 등록합니다.', when: '시험·상담 예약 입력 시' },
      { num: 2, target: [49.1, 2.3, 11.6, 4.0], card: [36, 20], title: 'NEIS 학사일정 가져오기', text: '2번 메뉴에서 고른 우리 학교 공식 학사일정을 한 번에 채웁니다.', when: '학기 초 1클릭 일괄 등록' },
      { num: 3, target: [70.1, 2.1, 9.8, 4.3], card: [66, 20], title: 'Google 동기화', text: '내 Google 기본 캘린더와 일정을 양방향으로 동기화합니다.', when: '스마트폰과 일정 공유 시' },
      { num: 4, target: [86.7, 2.3, 10.2, 3.9], card: [66, 52], title: '가져오기·내보내기', text: '학생 상담 기록만 따로 엑셀로 내려받거나 시트로 복사합니다.', when: '상담 누가기록 정리 시' }
    ],
    callout: { type: 'good', text: '일정 카드는 마우스로 끌어 다른 날짜로 바로 옮길 수 있고, 휠로 달을 넘깁니다.' }
  },
  {
    section: 'calendar',
    img: 'guide/v67-calendar-delete.png',
    title: '일정 수정·삭제 및 학생 상담 기록 관리',
    desc: '<ul>' +
      '<li><strong>일정 카드 클릭 수정</strong>달력에 표시된 일정을 클릭하면 제목·날짜·분류·메모를 즉시 수정하거나 삭제할 수 있습니다.</li>' +
      '<li><strong>상담 기록 전용 필드</strong>분류를 [상담]으로 고르면 학생 이름, 상담 방법(대면·전화), 후속 조치를 함께 기록합니다.<span class="gd-when">활용: 학부모·학생 상담 누가기록</span></li>' +
      '</ul>',
    callout: { type: 'good', text: '상담 기록은 기본적으로 교무수첩 내부에만 저장되며, Google 캘린더 연동 시 포함 여부를 직접 켜고 끌 수 있습니다.' }
  },
  {
    section: 'calendar',
    img: 'guide/v67-google-calendar-sync.png',
    title: 'Google 캘린더 양방향 동기화 — 통합 로그인으로 스마트폰까지 연결',
    desc: '<ul>' +
      '<li><strong>1번 메뉴와 통합 로그인</strong>[1 Google 시트 연동]에서 통합 로그인을 하면 Google 기본 캘린더와 자동으로 양방향 동기화됩니다.<span class="gd-when">활용: 스마트폰 캘린더와 일정 공유</span></li>' +
      '<li><strong>캘린더 전용 계정 변경</strong>필요 시 교무수첩 계정은 유지한 채 캘린더만 다른 Google 계정으로 연결할 수도 있습니다.</li>' +
      '</ul>',
    callout: { type: 'warn', text: '브라우저를 새로고침하면 보안상 캘린더 토큰이 초기화되므로, 새로고침 후에는 달력 하단의 연결/동기화 버튼을 한 번 눌러주세요.' }
  },

  // ===== 5. 6번 메뉴: 교실 도구 =====
  {
    section: 'classroom',
    img: 'guide/v11_09_classroom_menu.png',
    title: '6번 메뉴 · 교실 도구 — 발표·학급운영·교과활동·게임·웹툰·학습앱',
    desc: '<ul>' +
      '<li><strong>발표·뽑기 & 수업·학급 운영</strong>룰렛, 사다리타기, 랜덤게임, 간편 타이머, 칭찬 포인트, 1인 1역, 알림장을 바로 엽니다.<span class="gd-when">활용: 수업 활동·학급 자치</span></li>' +
      '<li><strong>교과·사고력 & 학습용·활용 앱</strong>SET, Stock Arena v3(모의주식투자), 기하 벡터, 인공지능 수학, 타로·관상 앱을 실행합니다.</li>' +
      '<li><strong>웹툰 연재 (수학·교양)</strong>인공지능 수학, 경제수학 등 교과 흥미를 높이는 웹툰 연재 링크가 함께 정리되어 있습니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '각 카테고리를 클릭해 접고 펼 수 있어 수업 중 필요한 도구만 깔끔하게 찾아 실행할 수 있습니다.' }
  },
  {
    section: 'classroom',
    img: 'guide/v11_10_praise_modal.png',
    title: '교실 도구 · 칭찬 포인트 — 성장형 점수 적립과 시트 연동',
    desc: '<ul>' +
      '<li><strong>+1 / +5 즉시 부여</strong>명렬표와 연결된 학생 목록에서 발표·선행·과제 완수 포인트를 바로 부여합니다.<span class="gd-when">활용: 학급 긍정 행동 강화</span></li>' +
      '<li><strong>시트 실시간 동기화</strong>기록된 점수는 Google 마스터 시트 [8_출결_칭찬_기록] 탭에 자동 반영됩니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '좌석표에서 학생을 클릭해 부여한 칭찬 점수와 합산되어 관리됩니다.' }
  },
  {
    section: 'classroom',
    img: 'guide/v11_11_duty_modal.png',
    title: '교실 도구 · 1인 1역 — 직접 배정 · 무작위 배정 · 엑셀 연동',
    desc: '<ul>' +
      '<li><strong>엑셀 양식 & 무작위 배정</strong>학급 역할 목록을 올려 무작위로 배정하거나 직접배정 시트로 한 번에 불러옵니다.<span class="gd-when">활용: 학기 초·월별 역할 교대</span></li>' +
      '<li><strong>좌석표·행특 연동</strong>배정된 역할은 좌석표 학생 카드에 표시되어 학급 전체가 함께 확인합니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '1인 1역 수행 기록은 학기 말 창체·행특 작성 시 구체적인 근거 자료로 활용할 수 있습니다.' }
  },
  {
    section: 'classroom',
    img: 'guide/v11_12_notice_modal.png',
    title: '교실 도구 · 오늘의 알림장 & 간편 타이머',
    desc: '<ul>' +
      '<li><strong>오늘의 알림장</strong>준비물·과제·공지를 입력해 칠판 프로젝터에 크게 띄우고 날짜별로 자동 저장합니다.<span class="gd-when">활용: 아침 조회·종례 안내</span></li>' +
      '<li><strong>간편 타이머</strong>1·3·5·10분 프리셋이나 직접 입력한 시간으로 전체화면 카운트다운과 종료 알림음을 실행합니다.<span class="gd-when">활용: 모둠 토의·문제 풀이</span></li>' +
      '</ul>',
    callout: { type: 'good', text: '타이머와 알림장 모두 교실 뒷자리에서도 선명하게 보이도록 큰 글씨와 전체화면 모드를 지원합니다.' }
  },

  // ===== 6. 7번 메뉴: 구글 시트 · 추천 사이트 & 설정·백업 =====
  {
    section: 'favorites',
    img: 'guide/v11_14_fav_sheets.png',
    title: '7번 메뉴 · 구글 시트 모음 & 추천 사이트 — 나만의 링크 관리',
    desc: '<ul>' +
      '<li><strong>7 구글 시트 모음</strong>연결된 마스터 시트(10개 탭)와 시트 메모장에 등록한 주요 스프레드시트를 원클릭으로 엽니다.<span class="gd-when">활용: 시트 원본 데이터 직접 확인</span></li>' +
      '<li><strong>추천 사이트 & 직접 추가</strong>수학관련 사이트, 수업·클래스 사이트를 모아두고 [+ 사이트 추가]로 내 즐겨찾기를 덧붙입니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '학교 홈페이지, 학년 공유 드라이브 등 매일 여는 주소를 등록해 두면 브라우저 북마크처럼 쓸 수 있습니다.' }
  },
  {
    section: 'settings',
    img: 'guide/v11_15_help_menu.png',
    title: '설정·도움말 — 9가지 화면 테마 · 올인원 매뉴얼 · 우측 하단 아이콘',
    desc: '<ul>' +
      '<li><strong>9가지 감성 화면 테마</strong>초록 식물도감, T-ong 클래식, 별빛 야간 등 9종 테마 중 교실 분위기에 맞게 고릅니다.<span class="gd-when">활용: 눈이 편한 테마 선택</span></li>' +
      '<li><strong>올인원 동기화 매뉴얼</strong>T-ong과 NEIS-Help(생기부·대입상담)를 함께 쓰는 2트랙 상세 매뉴얼을 새 창으로 엽니다.</li>' +
      '<li><strong>데이터 저장 원칙</strong>교무 자료는 T-ong 서버에 저장하지 않으며 선생님 브라우저와 본인이 연결한 Google 시트에만 보관됩니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '우측 하단 [오늘] 아이콘을 눌러 내 사진으로 커스터마이징해 보세요.' }
  },
  {
    section: 'settings',
    img: 'guide/v67-data-backup.png',
    title: '데이터 백업 & 복원 — PC 교체 · 브라우저 초기화 전 필수 보관',
    desc: '<ul>' +
      '<li><strong>백업 코드 생성 및 보관</strong>홈 ▾ → [데이터 백업]에서 전체 자료를 백업 코드/파일로 내려받아 개인 드라이브에 저장합니다.<span class="gd-when">활용: 학교 PC ↔ 집 PC 이전 시</span></li>' +
      '<li><strong>Google 시트와 병행</strong>출결·관찰·칭찬·1인1역은 Google 시트로 실시간 동기화하고, 자리배치·시간표·설정은 백업으로 옮깁니다.</li>' +
      '</ul>',
    callout: { type: 'danger', text: '백업 파일에는 학생 명렬과 학급 기록이 포함되어 있으므로 공용 PC나 단체 메신저에 올리지 말고 개인 드라이브에 안전하게 보관하세요.' }
  }
];

var GUIDE_SECTIONS = [
  { key: 'start', label: '시작·1~2번' },
  { key: 'seat', label: '3. 자리배치' },
  { key: 'attend', label: '4. 수업·출결' },
  { key: 'calendar', label: '5. 일정·달력' },
  { key: 'classroom', label: '6. 교실도구' },
  { key: 'favorites', label: '7. 시트·사이트' },
  { key: 'settings', label: '설정·백업' }
];

function guideSectionKey(slide) {
  if (slide.section) return slide.section;
  var img = slide.img || '';
  if (/20(?:A)?-/.test(img)) return 'attend';
  if (/0B-|0E-|15-|16-|17-|21A-/.test(img)) return 'classroom';
  if (/0C-|02-|03-|10-|12-|18-/.test(img)) return 'settings';
  if (/00-|0A-|01-/.test(img)) return 'start';
  return 'seat';
}

function guideSectionLabel(slide) {
  var key = guideSectionKey(slide);
  for (var i = 0; i < GUIDE_SECTIONS.length; i++) {
    if (GUIDE_SECTIONS[i].key === key) return GUIDE_SECTIONS[i].label;
  }
  return '사용법';
}
GUIDE_SLIDES = GUIDE_SECTIONS.reduce(function (ordered, section) {
  return ordered.concat(GUIDE_SLIDES.filter(function (slide) {
    return guideSectionKey(slide) === section.key;
  }));
}, []);
function initGuide() {
  var track = document.getElementById('guideTrack');
  var dots = document.getElementById('guideDots');
  var sections = document.getElementById('guideSections');
  if (!track || !dots) return;

  if (sections) {
    GUIDE_SECTIONS.forEach(function (section) {
      var button = document.createElement('button');
      button.type = 'button';
      button.className = 'guide-section-btn';
      button.dataset.section = section.key;
      button.textContent = section.label;
      button.addEventListener('click', function () {
        for (var i = 0; i < GUIDE_SLIDES.length; i++) {
          if (guideSectionKey(GUIDE_SLIDES[i]) === section.key) { goToSlide(i); break; }
        }
      });
      sections.appendChild(button);
    });
  }

  GUIDE_SLIDES.forEach(function(slide, i) {
    var el = document.createElement('div');
    el.className = 'guide-slide is-stage';
    var hasSpots = Array.isArray(slide.spots) && slide.spots.length > 0;
    var spotsHtml = '';
    var cardsHtml = '';
    var linesHtml = '';
    var dimHtml = hasSpots ? '<div class="guide-stage-dim"></div>' : '';

    if (hasSpots) {
      spotsHtml = slide.spots.map(function (sp, idx) {
        var t = sp.target;
        return '<div class="guide-spot" style="left:' + t[0] + '%;top:' + t[1] + '%;width:' + t[2] + '%;height:' + t[3] + '%;">' +
          '<span class="guide-spot-num">' + (sp.num || (idx + 1)) + '</span></div>';
      }).join('');
      cardsHtml = slide.spots.map(function (sp, idx) {
        var c = sp.card;
        var whenHtml = sp.when ? '<span class="guide-call-when">' + escapeHtml(sp.when) + '</span>' : '';
        return '<div class="guide-call-card" style="left:' + c[0] + '%;top:' + c[1] + '%;">' +
          '<div class="guide-call-head"><span class="guide-call-num">' + (sp.num || (idx + 1)) + '</span><strong>' + escapeHtml(sp.title) + '</strong></div>' +
          '<p>' + sp.text + '</p>' + whenHtml +
        '</div>';
      }).join('');
      linesHtml = slide.spots.map(function (sp) {
        var t = sp.target;
        var c = sp.card;
        var sx = Math.round((c[0] + 13) * 10);
        var sy = Math.round((c[1] + (c[1] > t[1] ? 1 : 9)) * 10);
        var tx = Math.round((t[0] + t[2] / 2) * 10);
        var ty = Math.round((t[1] + (c[1] > t[1] ? t[3] : t[3] / 2)) * 10);
        if (c[0] + 28 < t[0]) { sx = Math.round((c[0] + 26) * 10); sy = Math.round((c[1] + 7) * 10); tx = Math.round(t[0] * 10); ty = Math.round((t[1] + t[3] / 2) * 10); }
        else if (c[0] > t[0] + t[2]) { sx = Math.round(c[0] * 10); sy = Math.round((c[1] + 7) * 10); tx = Math.round((t[0] + t[2]) * 10); ty = Math.round((t[1] + t[3] / 2) * 10); }
        return '<line x1="' + sx + '" y1="' + sy + '" x2="' + tx + '" y2="' + ty + '" stroke="#ef4444" stroke-width="3.2" stroke-linecap="round" marker-end="url(#guideArrowHead)" />';
      }).join('');
    } else if (slide.desc) {
      cardsHtml = '<div class="guide-stage-float-desc">' + slide.desc + '</div>';
    }

    var tipBadge = slide.callout ? '<span class="guide-stage-tip">' + escapeHtml(slide.callout.text) + '</span>' : '';
    el.innerHTML =
      '<div class="guide-stage-banner">' +
        '<div><span class="guide-stage-sub">' + guideSectionLabel(slide) + ' · STEP ' + (i + 1) + ' / ' + GUIDE_SLIDES.length + '</span><h4>' + slide.title + '</h4></div>' +
        tipBadge +
      '</div>' +
      '<div class="guide-stage">' +
        '<img class="guide-stage-img" src="' + slide.img + '?v=' + encodeURIComponent(APP_VERSION) + '" alt="' + escapeHtml(slide.title) + '" loading="lazy">' +
        dimHtml +
        spotsHtml +
        (hasSpots ? '<svg class="guide-stage-svg" viewBox="0 0 1000 1000" preserveAspectRatio="none"><defs><marker id="guideArrowHead" markerWidth="6" markerHeight="6" refX="5" refY="3" orient="auto"><path d="M0,0 L6,3 L0,6 Z" fill="#ef4444" /></marker></defs>' + linesHtml + '</svg>' : '') +
        cardsHtml +
      '</div>';
    track.appendChild(el);

    var dot = document.createElement('button');
    dot.type = 'button';
    dot.className = 'guide-dot' + (i === 0 ? ' active' : '');
    dot.setAttribute('aria-label', (i + 1) + '번째 안내');
    dot.addEventListener('click', function() { goToSlide(i); });
    dots.appendChild(dot);
  });

  document.getElementById('guidePrev').addEventListener('click', function() { goToSlide(guideIndex - 1); });
  document.getElementById('guideNext').addEventListener('click', function() { goToSlide(guideIndex + 1); });
  document.getElementById('guideStartBtn').addEventListener('click', function() {
    if (guideIndex === GUIDE_SLIDES.length - 1) closeGuide();
    else goToSlide(guideIndex + 1);
  });
  document.getElementById('guideBtn').addEventListener('click', openGuide);
  var introGuideBtn = document.getElementById('introGuideBtn');
  if (introGuideBtn) introGuideBtn.addEventListener('click', openGuide);

  document.getElementById('guideDontShow').addEventListener('change', function() {
    try {
      if (this.checked) localStorage.setItem(GUIDE_SEEN_KEY, '1');
      else localStorage.removeItem(GUIDE_SEEN_KEY);
    } catch (e) { /* 무시 */ }
  });

  // 키보드 좌우 이동
  document.addEventListener('keydown', function(e) {
    var modal = document.getElementById('guideModal');
    if (!modal.classList.contains('show')) return;
    if (e.key === 'ArrowRight') goToSlide(guideIndex + 1);
    else if (e.key === 'ArrowLeft') goToSlide(guideIndex - 1);
    else if (e.key === 'Escape') closeGuide();
  });

  // 모바일 스와이프
  var startX = null;
  track.addEventListener('touchstart', function(e) { startX = e.touches[0].clientX; }, { passive: true });
  track.addEventListener('touchend', function(e) {
    if (startX === null) return;
    var dx = e.changedTouches[0].clientX - startX;
    if (Math.abs(dx) > 55) goToSlide(guideIndex + (dx < 0 ? 1 : -1));
    startX = null;
  }, { passive: true });

  goToSlide(0);
}

function goToSlide(i) {
  guideIndex = Math.max(0, Math.min(GUIDE_SLIDES.length - 1, i));
  document.getElementById('guideTrack').style.transform = 'translateX(-' + (guideIndex * 100) + '%)';
  document.querySelectorAll('.guide-dot').forEach(function(d, idx) {
    d.classList.toggle('active', idx === guideIndex);
  });
  var sectionKey = guideSectionKey(GUIDE_SLIDES[guideIndex]);
  document.querySelectorAll('.guide-section-btn').forEach(function (button) {
    button.classList.toggle('active', button.dataset.section === sectionKey);
  });
  document.getElementById('guidePrev').disabled = guideIndex === 0;
  document.getElementById('guideNext').disabled = guideIndex === GUIDE_SLIDES.length - 1;
  document.getElementById('guideStartBtn').textContent =
    guideIndex === GUIDE_SLIDES.length - 1 ? '시작하기 🚀' : '다음 →';
  var status = document.getElementById('guideStatus');
  if (status) status.textContent = (guideIndex + 1) + ' / ' + GUIDE_SLIDES.length;
  var activeDot = document.querySelector('.guide-dot.active');
  if (activeDot && activeDot.scrollIntoView) activeDot.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  var vp = document.querySelector('.guide-viewport');
  if (vp) vp.scrollTop = 0;
  var slide = document.querySelectorAll('.guide-slide')[guideIndex];
  if (slide) slide.scrollTop = 0;
}

function openGuide() {
  goToSlide(0);
  openModal('guideModal');
}

function openGuideSection(sectionKey) {
  openGuide();
  for (var i = 0; i < GUIDE_SLIDES.length; i++) {
    if (guideSectionKey(GUIDE_SLIDES[i]) === sectionKey) { goToSlide(i); break; }
  }
}

function closeGuide() {
  closeModal('guideModal');
  try { localStorage.setItem(GUIDE_SEEN_KEY, '1'); } catch (e) { /* 무시 */ }
}

// ===================================================================
// 페이지 스크롤 허용 — 좌석표를 강제로 축소하지 않습니다.
// 호출 지점은 그대로 두고, 이전 버전에서 남은 transform만 정리합니다.
// ===================================================================
function fitSeatingToViewport() {
  var stage = document.getElementById('stage');
  var header = document.querySelector('.app-header');
  if (stage && header) {
    // Header controls can wrap to two rows. Use their real height, not a fixed offset.
    stage.style.setProperty('--seat-header-height', Math.ceil(header.getBoundingClientRect().height) + 'px');
    var viewport = window.visualViewport;
    var availableHeight = viewport && Math.abs(viewport.scale - 1) < 0.01 ? viewport.height : window.innerHeight;
    stage.style.setProperty('--seat-viewport-height', Math.floor(availableHeight) + 'px');
  }
  var area = document.getElementById('seatingArea');
  if (!area) return;
  area.style.transform = 'none';
  area.style.marginBottom = '0';
}

var fitTimer = null;
function scheduleFit() {
  clearTimeout(fitTimer);
  fitTimer = setTimeout(fitSeatingToViewport, 60);
}

// ===================================================================
// 사용자 즐겨찾기 — 일반·교과서·진로진학 사이트를 브라우저에 저장
// ===================================================================
var CustomFavorites = (function () {
  var KEY = 'seatCustomFavorites_v1';
  var items = [];
  function el(id) { return document.getElementById(id); }
  function kindOf(item) { var kind=item&&item.kind; return ['math','classroom','textbook','admission'].indexOf(kind)>=0?kind:'general'; }
  function kindTitle(kind) { return {math:'📐 내 수학관련 사이트',classroom:'🏫 내 수업 및 클래스 사이트',textbook:'📚 내 교과서 사이트',admission:'🎓 내 진로진학 사이트'}[kind]||'🔖 나만의 즐겨찾기'; }
  function kindIcon(kind) { return {math:'📐 ',classroom:'🏫 ',textbook:'📖 ',admission:'🎓 '}[kind]||'🌐 '; }
  function kindEmpty(kind) { return {math:'추가한 수학관련 사이트가 없습니다.',classroom:'추가한 수업 및 클래스 사이트가 없습니다.',textbook:'추가한 교과서 사이트가 없습니다.',admission:'추가한 진로진학 사이트가 없습니다.'}[kind]||'아직 추가한 사이트가 없습니다.'; }
  function kindAdded(kind) { return {math:'수학관련 사이트를 추가했습니다',classroom:'수업 및 클래스 사이트를 추가했습니다',textbook:'교과서 사이트를 추가했습니다',admission:'진로진학 사이트를 추가했습니다'}[kind]||'나만의 즐겨찾기에 추가했습니다'; }
  function load() {
    try {
      var raw = JSON.parse(localStorage.getItem(KEY) || '[]');
      items = Array.isArray(raw) ? raw.filter(function (x) { return x && x.name && x.url; }).map(function (x) {
        return { id:x.id || generateId(), name:String(x.name), url:String(x.url), kind:kindOf(x) };
      }) : [];
    } catch (e) { items = []; }
  }
  function save() {
    try { localStorage.setItem(KEY, JSON.stringify(items)); }
    catch (e) { showToast('즐겨찾기를 저장할 수 없습니다', 'error'); }
  }
  function normalizeUrl(value) {
    var url = String(value || '').trim();
    if (!/^https?:\/\//i.test(url)) url = 'https://' + url;
    try {
      var parsed = new URL(url);
      if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') return '';
      return parsed.href;
    } catch (e) { return ''; }
  }
  function renderInto(box, kind) {
    if (!box) return;
    var list = items.filter(function (item) { return kindOf(item) === kind; });
    if (!list.length) {
      box.innerHTML = '<p class="custom-fav-empty">' + kindEmpty(kind) + '</p>';
      return;
    }
    box.innerHTML = list.slice().sort(function (a,b) { return b.name.length - a.name.length || a.name.localeCompare(b.name); }).map(function (item) {
      return '<div class="custom-fav-item"><a href="' + escapeHtml(item.url) + '" target="_blank" rel="noopener noreferrer">' + kindIcon(kind) + escapeHtml(item.name) + '</a>' +
        '<button type="button" data-custom-fav-edit="' + item.id + '" aria-label="' + escapeHtml(item.name) + ' 수정">✎</button></div>';
    }).join('');
    box.querySelectorAll('[data-custom-fav-edit]').forEach(function (button) {
      button.addEventListener('click', function (e) { e.preventDefault(); e.stopPropagation(); openEditor(this.dataset.customFavEdit, kind); });
    });
  }
  function render() {
    renderInto(el('customFavList'), 'general');
    renderInto(el('customFavModalList'), 'general');
    renderInto(el('customMathList'), 'math');
    renderInto(el('customMathModalList'), 'math');
    renderInto(el('customClassList'), 'classroom');
    renderInto(el('customClassModalList'), 'classroom');
    renderInto(el('customTextbookList'), 'textbook');
    renderInto(el('customTextbookModalList'), 'textbook');
    renderInto(el('customAdmissionList'), 'admission');
    renderInto(el('customAdmissionModalList'), 'admission');
  }
  function openEditor(id, kind) {
    var modal = el('favoriteEditModal'); if (!modal) return;
    var item = items.find(function (x) { return x.id === id; });
    var selectedKind = item ? kindOf(item) : kindOf({kind:kind});
    el('customFavId').value = item ? item.id : '';
    el('customFavKind').value = selectedKind;
    el('customFavName').value = item ? item.name : '';
    el('customFavUrl').value = item ? item.url : '';
    var title = el('customFavModalTitle'); if (title) title.textContent = kindTitle(selectedKind);
    el('customFavDeleteBtn').hidden = !item;
    openModal('favoriteEditModal');
    setTimeout(function () { el('customFavName').focus(); }, 30);
  }
  function submit(e) {
    e.preventDefault();
    var name = el('customFavName').value.trim();
    var url = normalizeUrl(el('customFavUrl').value);
    var kind = kindOf({kind:el('customFavKind').value});
    if (!name || !url) { showToast('사이트 이름과 올바른 주소를 입력해 주세요', 'error'); return; }
    var id = el('customFavId').value;
    var item = { id: id || generateId(), name: name, url: url, kind: kind };
    var index = items.findIndex(function (x) { return x.id === id; });
    if (index >= 0) items[index] = item; else items.push(item);
    save(); render(); closeModal('favoriteEditModal'); showToast(index >= 0 ? '사이트를 수정했습니다' : kindAdded(kind));
  }
  function remove() {
    var id = el('customFavId').value;
    var item = items.find(function (x) { return x.id === id; });
    if (!item || !confirm('“' + item.name + '” 사이트를 삭제할까요?')) return;
    items = items.filter(function (x) { return x.id !== id; });
    save(); render(); closeModal('favoriteEditModal'); showToast('사이트를 삭제했습니다');
  }
  function sortBuiltInLinks() {
    document.querySelectorAll('.fav-category-grid').forEach(function (grid) {
      Array.prototype.slice.call(grid.children).sort(function (a,b) {
        var at=a.textContent.trim(),bt=b.textContent.trim();
        return bt.length-at.length||at.localeCompare(bt,'ko');
      }).forEach(function (link) { grid.appendChild(link); });
    });
  }
  function init() {
    load(); render(); sortBuiltInLinks();
    var add = el('customFavAddBtn'); if (add) add.addEventListener('click', function (e) { e.stopPropagation(); openEditor('', 'general'); });
    var addModal = el('customFavModalAddBtn'); if (addModal) addModal.addEventListener('click', function () { openEditor('', 'general'); });
    var addMath = el('customMathAddBtn'); if (addMath) addMath.addEventListener('click', function (e) { e.stopPropagation(); openEditor('', 'math'); });
    var addMathModal = el('customMathModalAddBtn'); if (addMathModal) addMathModal.addEventListener('click', function () { openEditor('', 'math'); });
    var addClass = el('customClassAddBtn'); if (addClass) addClass.addEventListener('click', function (e) { e.stopPropagation(); openEditor('', 'classroom'); });
    var addClassModal = el('customClassModalAddBtn'); if (addClassModal) addClassModal.addEventListener('click', function () { openEditor('', 'classroom'); });
    var addTextbook = el('customTextbookAddBtn'); if (addTextbook) addTextbook.addEventListener('click', function (e) { e.stopPropagation(); openEditor('', 'textbook'); });
    var addTextbookModal = el('customTextbookModalAddBtn'); if (addTextbookModal) addTextbookModal.addEventListener('click', function () { openEditor('', 'textbook'); });
    var addAdmission = el('customAdmissionAddBtn'); if (addAdmission) addAdmission.addEventListener('click', function (e) { e.stopPropagation(); openEditor('', 'admission'); });
    var addAdmissionModal = el('customAdmissionModalAddBtn'); if (addAdmissionModal) addAdmissionModal.addEventListener('click', function () { openEditor('', 'admission'); });
    var form = el('customFavForm'); if (form) form.addEventListener('submit', submit);
    var del = el('customFavDeleteBtn'); if (del) del.addEventListener('click', remove);
  }
  function getData() { return items.slice(); }
  function setData(value) { items = Array.isArray(value) ? value.map(function (x) { x.kind = kindOf(x); return x; }) : []; save(); render(); }
  return { init:init, render:render, getData:getData, setData:setData };
})();

// ===================================================================
// 상단 도움말
// ===================================================================
var ThemeManager = (function () {
  var KEY = 'tongTheme_v1';
  var THEMES = {
    classic: { label:'클래식', color:'#667eea' },
    romance: { label:'순정만화', color:'#d97d9f' },
    cute: { label:'몽글 스티커', color:'#f09b4d' },
    waterblue: { label:'물빛 산책', color:'#78aabd' },
    forestletter: { label:'초록 숲편지', color:'#819c6c' },
    night: { label:'별빛 야간', color:'#242a47' },
    campus: { label:'맑은 교정', color:'#6f9eb5' },
    botanical: { label:'초록 식물도감', color:'#708763' },
    teacherdesk: { label:'포근한 교사 책상', color:'#a47759' }
  };

  function selected() {
    var value = document.documentElement.getAttribute('data-theme') || 'botanical';
    return THEMES[value] ? value : 'botanical';
  }
  function close() {
    var menu = document.getElementById('themeMenu');
    var button = document.getElementById('themeBtn');
    if (menu) menu.classList.remove('open');
    if (button) button.setAttribute('aria-expanded', 'false');
  }
  function render(theme) {
    document.querySelectorAll('.theme-choice').forEach(function (button) {
      button.setAttribute('aria-pressed', String(button.dataset.theme === theme));
    });
    var trigger = document.getElementById('themeBtn');
    if (trigger) trigger.textContent = '🎨 ' + THEMES[theme].label + ' ▾';
    var curLabel = document.getElementById('currentThemeName');
    if (curLabel && THEMES[theme]) curLabel.textContent = THEMES[theme].label;
    var meta = document.querySelector('meta[name="theme-color"]');
    if (meta) meta.setAttribute('content', THEMES[theme].color);
  }
  function apply(theme, persist, notify) {
    if (!THEMES[theme]) theme = 'botanical';
    document.documentElement.setAttribute('data-theme', theme);
    if (persist) {
      try { localStorage.setItem(KEY, theme); } catch (e) {}
    }
    render(theme);
    if (notify) showToast(THEMES[theme].label + ' 테마를 적용했습니다');
  }
  function init() {
    apply(selected(), false, false);
    var trigger = document.getElementById('themeBtn');
    var menu = document.getElementById('themeMenu');
    if (!trigger || !menu) return;
    trigger.addEventListener('click', function (event) {
      event.stopPropagation();
      var open = !menu.classList.contains('open');
      menu.classList.toggle('open', open);
      trigger.setAttribute('aria-expanded', String(open));
    });
    menu.addEventListener('click', function (event) { event.stopPropagation(); });
    menu.querySelectorAll('.theme-choice').forEach(function (button) {
      button.addEventListener('click', function () {
        apply(this.dataset.theme, true, true);
        close();
      });
    });
    document.addEventListener('click', close);
    document.addEventListener('keydown', function (event) { if (event.key === 'Escape') close(); });
  }
  return { init:init, apply:apply, current:selected };
})();

var UsageStats = (function () {
  var VISITOR_KEY = 'notetongAnonymousVisitor_v1';
  function isLocal() { return location.hostname === 'localhost' || location.hostname === '127.0.0.1' || /^192\.168\./.test(location.hostname) || /^10\./.test(location.hostname); }
  function setStatus(text) { var status=document.getElementById('usageStatsStatus'); if(status)status.textContent=text; }
  function setCounts(today,total) {
    var todayEl=document.getElementById('todayVisitorCount'), totalEl=document.getElementById('totalVisitCount');
    var footerToday=document.getElementById('footerTodayVisitorCount'), footerTotal=document.getElementById('footerTotalVisitCount');
    var hasToday=Number.isFinite(today), hasTotal=Number.isFinite(total);
    var todayText=hasToday?today.toLocaleString('ko-KR'):'__', totalText=hasTotal?total.toLocaleString('ko-KR'):'__';
    if(todayEl)todayEl.textContent=hasToday?todayText+'명':'__';
    if(totalEl)totalEl.textContent=hasTotal?totalText+'회':'__';
    if(footerToday)footerToday.textContent=todayText;
    if(footerTotal)footerTotal.textContent=totalText;
  }
  function koreaDateKey() {
    var parts=new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Seoul',year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(new Date());
    var map={};parts.forEach(function(p){map[p.type]=p.value;});return map.year+'-'+map.month+'-'+map.day;
  }
  function visitorId() {
    try {
      var saved=localStorage.getItem(VISITOR_KEY);if(saved)return saved;
      var id=(window.crypto&&crypto.randomUUID)?crypto.randomUUID().replace(/-/g,''):generateId().replace(/[^a-zA-Z0-9_-]/g,'');
      localStorage.setItem(VISITOR_KEY,id);return id;
    } catch(e){return generateId().replace(/[^a-zA-Z0-9_-]/g,'');}
  }
  async function dailyToken() {
    var raw=koreaDateKey()+':'+visitorId();
    if(window.crypto&&crypto.subtle){
      var bytes=new TextEncoder().encode(raw),hash=await crypto.subtle.digest('SHA-256',bytes);
      return Array.from(new Uint8Array(hash)).map(function(x){return x.toString(16).padStart(2,'0');}).join('');
    }
    return Array.from(raw).reduce(function(a,c){return ((a<<5)-a+c.charCodeAt(0))|0;},0).toString(16).replace('-','').padEnd(64,'0').slice(0,64);
  }
  async function recordVisit() {
    if(isLocal()){setCounts(NaN,NaN);setStatus('');return;}
    try {
      var response=await fetch('/api/usage',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({visitorToken:await dailyToken()}),cache:'no-store'});
      if(!response.ok)throw new Error('usage api '+response.status);
      var data=await response.json();setCounts(Number(data.todayVisitors),Number(data.totalVisits));
      setStatus('날짜별 익명 방문 수와 총 접속 횟수만 집계 · 학생 자료는 포함하지 않음');
    } catch(e){setCounts(NaN,NaN);setStatus('통계 저장소 연결 대기 중 · 학생 자료는 전송하지 않음');}
  }
  function init(){recordVisit();}
  return { init:init };
})();

var HelpMenu = (function () {
  function el(id) { return document.getElementById(id); }
  function showInfo(kind) {
    var title = el('helpInfoTitle'), body = el('helpInfoBody');
    if (!title || !body) return;
    if (kind === 'version') {
      title.textContent = 'ℹ️ 버전 안내';
      body.innerHTML = '<div class="help-version"><b>T-ong ' + escapeHtml(APP_VERSION) + '</b><span>담임·교과교사를 위한 디지털 교무수첩</span><p class="help-version-message">바쁜 선생님의 하루에 작은 여유를 더하고자 만든 무료 교무수첩입니다.</p></div>' +
        '<ul><li>방문 횟수 통계만 익명으로 집계되며, 명렬표·출결·좌석·상담 기록은 이 브라우저에만 저장됩니다.</li><li>Google 동기화를 직접 켠 경우 선택한 일정 항목만 본인 Google 계정에 직접 저장되며 T-ong 서버는 해당 일정을 수집하지 않습니다.</li><li>Google 동기화를 쓰지 않는 자료는 다른 기기에서 자동으로 보이지 않으므로 데이터 백업을 이용해 주세요.</li></ul>';
    } else if (kind === 'updates') {
      title.textContent = '✨ 업데이트 사항';
      body.innerHTML = '<ul class="help-update-list"><li><b>v' + escapeHtml(APP_VERSION) + '</b><span>모든 자리 형태에서 무작위·학번 가로순·학번 세로순 배치 방식 선택 기능 추가 · 홈 하위메뉴에 데이터 백업 바로가기 추가 · 칭찬 포인트를 알·부화 중·병아리·닭 성장 단계로 표시 · 자리배치 학생 팝업에서 오늘 출결 즉시 등록·출결관리 연동 · Google 일정 전체 다시 가져오기 버튼 추가 · Google 기본 캘린더 양방향 동기화 및 직접 생성 일정 가져오기 수정 · 월간 일정 카드에서 수정·삭제 접근 개선 · 월간·주간 일정 드래그 이동과 모바일 길게 누르기·실행 취소 추가 · 같은 성별끼리 우선 배치하는 추첨 옵션 추가 · 모바일 데이터 백업 화면 개선 · 저장한 자리배치 불러오기 버튼 추가 · 수업 중 교탁 시점 전환 추가 · 명렬표 이름 오름차순 정렬 · 게임 메뉴에 핸드 카트라이더 추가 · 즐겨찾기의 EBSi 링크를 EBSi 교사지원센터로 변경 · 컴시간알리미(교사용)를 교실운영 메뉴 최상단으로 이동 · 게임 메뉴에 SET(셋)·랜덤게임 룰렛·윷놀이 추가 · Google OAuth 클라이언트 ID 발급·연결 4단계 스크린샷 도움말 추가 · Google 캘린더 사용자별 양방향 동기화(앱 사용 중 변경 즉시·약 1분 간격)와 개인정보 보호 옵션 추가 · 자리배치에서 현재 학급 명렬표를 학번·이름·성별 엑셀 파일로 다시 다운로드하는 기능 추가 · 즐겨찾기 교무업무에 국가교육과정정보센터(NCIC) 추가 · 모바일에서 도움말 메뉴가 뒤쪽 레이아웃에 가려지지 않도록 전면 고정 표시 · 모둠형 학생용·교탁용 출력을 화면과 같은 격자·모둠 번호로 수정 · 초록 식물도감을 신규 기본 테마로 지정 · 이미지 테마 3종 추가 · 별빛 야간 하위 메뉴 가독성 개선 · 월간 달력 5주 높이 적용 · 데이터·기록 드롭다운 표시 수정 · 9가지 화면 테마와 브라우저별 선택 저장 · 사용 설명서 36장 최신 화면 개편 · 교실 게임 메뉴 추가 · 도움말의 오늘의 학교(NEIS 실험) 복원 · 일정관리 시트 복사 예시 · 속성 수정/삭제·학부모 상담 대상 · Qwen 즐겨찾기 · 출결 학생 종합 기록 연결·칭찬 +1/−1 · 자리배치/출결관리 독립 탭 전환 · 개발자 응원 설문 연결</span></li><li><b>컴시간알리미(교사용)</b><span>교실운영 맨 위에서 새 창으로 열고 학교 코드를 설정합니다.</span></li><li><b>학급 도구 정리</b><span>상단에서 학급과 명렬표를 바로 관리하고, 출결관리에서도 학급 선택과 명렬표 업로드 가능</span></li><li><b>로컬 저장 안내 강화</b><span>익명 방문 횟수 통계 외 모든 교무 자료는 현재 브라우저에만 저장</span></li>' +
        '<li><b>즐겨찾기 묶음</b><span>수학관련 사이트·수업 및 클래스·교과서 사이트·생성형 AI·교무업무·진로진학·교사 도구함 순서로 정리하고 나만의 즐겨찾기는 맨 아래 고정</span></li>' +
        '<li><b>익명 방문 통계</b><span>오늘 방문자 수와 총 접속 횟수만 집계하며 명렬표·출결·상담 자료는 전송하지 않음</span></li>' +
        '<li><b>새 독립 주소</b><span>t-ong.vercel.app에서 T-ong 업그레이드판 운영</span></li></ul>';
      var latestUpdate = body.querySelector('.help-update-list li span');
      if (latestUpdate) latestUpdate.insertAdjacentText('afterbegin', 'v5.17.0-preview.70: 게임 메뉴에 온몸으로 찰칵! 포즈 챌린지 추가 · 웹사이트와 Android 앱에서 새 창으로 열기 · ');
      if (latestUpdate) latestUpdate.insertAdjacentText('afterbegin', '좌석 설정을 기본 접힘으로 변경하고 성별·배치 순서·설정 변경을 한 묶음으로 접기 · 무작위·학번 가로순·세로순 선택지를 세로 3줄로 정리 · 앱의 실제 표시 높이 변화에 맞춰 옵션 패널 스크롤 영역 갱신 · ');
      if (latestUpdate) latestUpdate.insertAdjacentText('afterbegin', '낮은 가로 화면의 자리배치 우측 옵션 패널 스크롤·잘림 수정 · 태블릿 중간 폭 레이아웃 충돌 수정 · 배치 순서 접기와 현재 선택 방식 표시 · 좌석 설정 접기 버튼 접근성 개선 · 출력·명렬표·배치기간 하위 메뉴 잘림 방지 · ');
      if (latestUpdate) latestUpdate.insertAdjacentText('afterbegin', '사용 설명서 ' + GUIDE_SLIDES.length + '장 기능 대조 및 최신 화면 보완 · 학번 가로·세로순과 성별 우선 조건 안내 · 자리에서 오늘 출결 등록·교탁 시점·배치 복원 안내 · 성장형 칭찬 및 홈 백업·기기별 저장 안내 · Google 전체 가져오기 양방향 주의사항 보완 · ');
      if (latestUpdate) latestUpdate.insertAdjacentText('afterbegin', '\ud2b9\uc815 \uc790\ub9ac \uace0\uc815 \uce90\ub7ec\uc140 \uc0ac\uc9c4\uc744 \uc2e4\uc81c \uace0\uc815 \uc88c\uc11d \ud654\uba74\uc73c\ub85c \uad50\uccb4 \u00b7 ');
      if (latestUpdate) latestUpdate.insertAdjacentText('afterbegin', '\uce90\ub7ec\uc140 4\u00b75\u00b76\ub2e8\uacc4 \uc0ac\uc9c4\uc744 \uc6b0\uce21 \uc88c\uc11d \uc124\uc815 \ud328\ub110\u00b7\ubaa8\ub460 \uc88c\uc11d\ud45c\u00b7\uc77c\ubc18 \uc88c\uc11d\ud45c\ub85c \uad50\uccb4 \u00b7 ');
      if (latestUpdate) latestUpdate.insertAdjacentText('afterbegin', '\ud558\ub2e8 \uc800\uc791\uad8c math-tong \ub9c1\ud06c\ub97c \uc0c8 T-ong \uc0ac\uc6a9\ubc95 \ube14\ub85c\uadf8 \uae00\ub85c \ubcc0\uacbd \u00b7 ');
      if (latestUpdate) latestUpdate.insertAdjacentText('afterbegin', '\uc990\uaca8\ucc3e\uae30\uc5d0 \uc218\ud559\uad00\ub828 \uc0ac\uc774\ud2b8\uc640 \uc0ac\uc6a9\uc790 \ucd94\uac00 \uae30\ub2a5\uc744 \ub123\uace0 \ud558\uc774\ub7ec\ub2dd\uc744 AI \uad50\uc218\ud559\uc2b5 \ud50c\ub7ab\ud3fc\uc73c\ub85c \ubcc0\uacbd \u00b7 \ubaa8\ubc14\uc77c \uad50\uc2e4\uc6b4\uc601 \uba54\ub274\ub97c \uc804\uba74 \uace0\uc815\ud574 \ubaa8\ub4e0 \ubc84\ud2bc \ud074\ub9ad \uac00\ub2a5 \u00b7 ');
      if (latestUpdate) latestUpdate.insertAdjacentText('afterbegin', 'v5.17.0-preview.71: 🤖 AI 도우미 추가(내 Claude·ChatGPT·Gemini 구독 계정으로 알림장·조회 멘트·가정통신문·상담 기록 요청, 학생 이름 자동 가림, 답변을 알림장에 바로 넣기) · Google 캘린더 연결 간소화(클라이언트 ID 입력 없이 버튼으로 연결, 이전 계정 한 번에 다시 연결, 만료 시 다시 연결 안내) · ');
      if (latestUpdate) latestUpdate.insertAdjacentText('afterbegin', 'v5.17.0-preview.72: 상단에 🤖 AI 메뉴 추가(AI 도우미 열기·AI 활용 안내·담임 5분 멘트) · AI 활용 안내 창(시작하기·활용 예시·개인정보 주의·AI 도구 모음) · 자리배치 우측 메뉴를 좌측처럼 접어 좌석 화면 넓히기 · ');
      if (latestUpdate) latestUpdate.insertAdjacentText('afterbegin', 'v5.17.0-preview.73: 🤖 AI 비서 채팅 추가(내 Claude API 키로 사이드바에서 대화 · 알림장 저장과 일정 조회·추가를 확인 후 실행 · 학생 이름 자동 가림) · AI 활용 안내에 API 키 발급 방법 탭 추가 · ');
    } else {
      title.textContent = '💜 개발자 응원하기';
      body.innerHTML = '<div class="support-message"><b>선생님의 한마디가 다음 업데이트를 만드는 큰 힘이 됩니다.</b><p>T-ong이 도움이 되었다면 응원 설문에 따뜻한 한마디를 남겨 주세요.</p></div>' +
        '<a class="btn primary support-link" href="https://hot-narcissus-8cf.notion.site/be5591743c164189bc26134f8a3d2315?pvs=105" target="_blank" rel="noopener noreferrer">💜 응원 설문 작성하기</a>' +
        '<p class="support-note">응원 한마디와 유용했던 기능을 자유롭게 남겨 주세요. 학생 이름 등 개인정보는 입력하지 말아 주세요.</p>';
    }
    openModal('helpInfoModal');
  }
  function init() {
    var button = el('helpBtn'), menu = el('helpMenu');
    function positionMobileMenu() {
      if (window.innerWidth > 900) {
        menu.style.removeProperty('--mobile-help-top');
        return;
      }
      var header = button.closest('.app-header');
      var headerBottom = header ? Math.ceil(header.getBoundingClientRect().bottom) : 104;
      var safeTop = Math.max(8, Math.min(headerBottom, window.innerHeight - 80));
      menu.style.setProperty('--mobile-help-top', safeTop + 'px');
    }
    if (!button || !menu) return;
    button.addEventListener('click', function (e) {
      e.stopPropagation();
      var extra = el('extraMenu'), fav = el('favMenu'), game = el('gameMenu'), ai = el('aiMenu');
      if (extra) extra.classList.remove('open');
      if (ai) ai.classList.remove('open');
      if (fav) fav.classList.remove('open');
      menu.classList.toggle('open');
      if (menu.classList.contains('open')) positionMobileMenu();
    });
    menu.addEventListener('click', function (e) { e.stopPropagation(); });
    document.addEventListener('click', function () { menu.classList.remove('open'); });
    window.addEventListener('resize', function () { if (menu.classList.contains('open')) positionMobileMenu(); });
    menu.querySelectorAll('[data-help]').forEach(function (item) {
      item.addEventListener('click', function () {
        menu.classList.remove('open');
        if (this.dataset.help === 'guide') openGuide();
        else if (this.dataset.help === 'school') SchoolToday.open();
        else showInfo(this.dataset.help);
      });
    });
  }
  return { init:init };
})();


// ===================================================================
// NEIS 교육정보 개방 포털 공통 호출부
//  인증키 없이 호출하면 NEIS 는 요청당 앞의 5건(샘플)만 돌려줍니다. 하루 시간표는 6~7교시라
//  5교시 뒤가 잘리므로, 아래 순서로 인증키가 붙은 경로를 먼저 씁니다.
//   1) 이 기기에 저장한 개인 인증키  2) 서버 프록시(/api/neis, 환경변수 NEIS_API_KEY)
//   3) 둘 다 없으면 인증키 없이 직접 호출 — 이때는 truncated 로 잘린 사실을 알려 줍니다.
// ===================================================================
var NeisApi = (function () {
  var USER_KEY_STORAGE = 'tongNeisApiKey_v1';
  var PROXY_PATH = '/api/neis';
  var KEY_ERROR_CODES = ['ERROR-290', 'ERROR-337', 'INFO-300'];  // 잘못된 키 · 일일 호출 한도 초과 · 사용 제한된 키
  var proxyState = null;  // null: 아직 모름 · true: 사용 가능 · false: 프록시 없음(키 미설정·정적 호스팅)
  var proxyProbe = null;

  function getUserKey() {
    try { return String(localStorage.getItem(USER_KEY_STORAGE) || '').trim(); } catch (e) { return ''; }
  }
  function setUserKey(value) {
    var key = String(value || '').trim();
    try {
      if (key) localStorage.setItem(USER_KEY_STORAGE, key);
      else localStorage.removeItem(USER_KEY_STORAGE);
    } catch (e) {}
    return key;
  }

  function neisError(code, message) {
    var error = new Error(message || ('NEIS ' + code));
    error.neisCode = code;
    error.isKeyError = KEY_ERROR_CODES.indexOf(code) >= 0;
    return error;
  }

  function buildQuery(params) {
    var query = new URLSearchParams();
    query.set('pIndex', '1');
    query.set('pSize', String(params.pSize || '100'));
    Object.keys(params).forEach(function (name) {
      if (name !== 'pSize' && params[name] != null && params[name] !== '') query.set(name, params[name]);
    });
    return query;
  }

  function parse(data, setName) {
    var result = data && data.RESULT;
    if (result && result.CODE) {
      if (result.CODE === 'INFO-200') return { rows: [], total: 0 };
      throw neisError(result.CODE, result.MESSAGE);
    }
    var set = data && data[setName];
    if (!Array.isArray(set)) return { rows: [], total: 0 };
    var headBlock = set.find(function (part) { return part && Array.isArray(part.head); });
    var rowBlock = set.find(function (part) { return part && Array.isArray(part.row); });
    var rows = rowBlock ? rowBlock.row : [];
    var total = rows.length;
    if (headBlock) {
      headBlock.head.forEach(function (item) {
        if (item && item.list_total_count != null) total = Number(item.list_total_count) || rows.length;
      });
    }
    return { rows: rows, total: total };
  }

  async function requestDirect(setName, params, key) {
    var query = buildQuery(params);
    query.set('Type', 'json');
    if (key) query.set('KEY', key);
    var response = await fetch('https://open.neis.go.kr/hub/' + setName + '?' + query.toString(), { cache: 'no-store' });
    if (!response.ok) throw new Error('NEIS HTTP ' + response.status);
    return parse(await response.json(), setName);
  }

  // 프록시를 쓸 수 없으면 null 을 돌려주고, 호출한 쪽이 인증키 없는 직접 호출로 넘어갑니다.
  async function requestProxy(setName, params) {
    var query = buildQuery(params);
    query.set('set', setName);
    var response;
    try { response = await fetch(PROXY_PATH + '?' + query.toString(), { cache: 'no-store' }); }
    catch (e) { return null; }
    if (response.status === 404 || response.status === 405 || response.status === 503) { proxyState = false; return null; }
    if (!response.ok) return null;
    var data;
    try { data = await response.json(); }
    catch (e) { proxyState = false; return null; }
    try {
      var parsed = parse(data, setName);
      proxyState = true;
      return parsed;
    } catch (error) {
      if (error.isKeyError) {
        proxyState = false;
        console.warn('[NEIS] 서버에 등록된 인증키를 쓸 수 없습니다: ' + error.message);
        return null;
      }
      throw error;
    }
  }

  // 첫 화면에서 여러 조회가 한꺼번에 나가도 프록시 유무 확인은 한 번만 하도록 줄을 세웁니다.
  async function viaProxy(setName, params) {
    if (proxyState === false) return null;
    if (proxyState === null && proxyProbe) {
      await proxyProbe;
      return proxyState === false ? null : requestProxy(setName, params);
    }
    if (proxyState === null) {
      var own = requestProxy(setName, params);
      proxyProbe = own.then(function () {}, function () {});
      try { return await own; } finally { proxyProbe = null; }
    }
    return requestProxy(setName, params);
  }

  // { rows, total, keyed, truncated } — truncated 는 NEIS 가 가진 건수보다 적게 받았다는 뜻입니다.
  async function fetchSet(setName, params) {
    params = params || {};
    var userKey = getUserKey();
    var parsed = null;
    var keyed = false;
    if (userKey) {
      parsed = await requestDirect(setName, params, userKey);
      keyed = true;
    } else {
      parsed = await viaProxy(setName, params);
      if (parsed) keyed = true;
      else parsed = await requestDirect(setName, params, '');
    }
    parsed.keyed = keyed;
    parsed.truncated = parsed.total > parsed.rows.length;
    return parsed;
  }

  return {
    fetch: fetchSet,
    getUserKey: getUserKey,
    setUserKey: setUserKey,
    hasUserKey: function () { return !!getUserKey(); }
  };
})();


// ===================================================================
// 도움말 실험 기능 — NEIS 오늘의 학교
// ===================================================================
var SchoolToday = (function () {
  var PROFILE_KEY='tongSchoolProfile_v1';
  var CACHE_KEY='tongSchoolTodayCache_v1';
  var profile=null;
  var selectedSchool=null;
  var requestNumber=0;
  var lastTimetableCut=null;  // 인증키가 없어 시간표가 잘렸을 때 {shown,total}

  function el(id){return document.getElementById(id);}
  function readJson(key,fallback){try{var value=JSON.parse(localStorage.getItem(key)||'null');return value==null?fallback:value;}catch(e){return fallback;}}
  function writeJson(key,value){try{localStorage.setItem(key,JSON.stringify(value));}catch(e){}}
  function compactDate(key){return String(key||'').replace(/-/g,'');}
  function weekdayText(key){var d=new Date(String(key)+'T00:00:00');return ['일','월','화','수','목','금','토'][d.getDay()]+'요일';}
  function timetableDataset(kind){if(kind==='초등학교')return'elsTimetable';if(kind==='중학교')return'misTimetable';if(kind==='특수학교')return'spsTimetable';return'hisTimetable';}
  function semesterFor(key){var d=new Date(String(key)+'T00:00:00'),month=d.getMonth()+1,year=d.getFullYear();if(month<=2)return{year:year-1,semester:2};return{year:year,semester:month<=8?1:2};}
  // 인증키 경로(개인 키·서버 프록시)를 함께 쓰도록 공통 호출부에 맡깁니다.
  async function request(setName,params){
    var result=await NeisApi.fetch(setName,params);
    if(result.truncated&&/Timetable$/.test(setName))lastTimetableCut={shown:result.rows.length,total:result.total};
    return result.rows;
  }
  function normalizeSchool(row){return{atptCode:String(row.ATPT_OFCDC_SC_CODE||''),schoolCode:String(row.SD_SCHUL_CODE||''),schoolName:String(row.SCHUL_NM||''),schoolKind:String(row.SCHUL_KND_SC_NM||''),location:String(row.LCTN_SC_NM||''),address:String(row.ORG_RDNMA||'')};}
  function setStatus(text,type){var node=el('schoolTodayStatus');if(!node)return;node.textContent=text||'';node.className='school-today-status'+(type?' '+type:'');}
  function setSettings(open){var panel=el('schoolSettingsPanel');if(panel)panel.hidden=!open;}
  function renderSelected(){
    var node=el('schoolSelectedInfo'),school=selectedSchool||profile;if(!node)return;
    if(!school){node.className='school-selected-info';node.textContent='선택한 학교가 없습니다.';return;}
    node.className='school-selected-info selected';node.innerHTML='<b>'+escapeHtml(school.schoolName)+'</b><span>'+escapeHtml([school.schoolKind,school.location,school.address].filter(Boolean).join(' · '))+'</span>';
  }
  function renderProfile(){
    var schoolName=el('schoolTodayName'),classText=el('schoolTodayClass'),subtitle=el('schoolTodaySubtitle'),empty=el('schoolTodayEmpty'),grid=el('schoolTodayGrid');
    if(profile){schoolName.innerHTML='🏫 '+escapeHtml(profile.schoolName)+' <span class="experimental-badge">실험</span>';classText.textContent=profile.schoolKind+' · '+profile.grade+'학년 '+profile.className+'반';subtitle.textContent=profile.schoolName+' · '+profile.grade+'학년 '+profile.className+'반';empty.hidden=true;grid.hidden=false;}
    else{schoolName.innerHTML='🏫 오늘의 학교 <span class="experimental-badge">실험</span>';classText.textContent='학교 설정이 필요합니다.';subtitle.textContent='학교를 설정하면 날짜별 학급 시간표와 급식을 확인할 수 있습니다.';empty.hidden=false;grid.hidden=true;}
    renderSelected();
  }
  function loading(){var html='<div class="school-data-loading"><i></i><span>NEIS 공개정보를 불러오는 중입니다.</span></div>';el('schoolTimetableList').innerHTML=html;el('schoolMealList').innerHTML=html;}
  function emptyHtml(message,hint){return'<div class="school-data-empty"><b>'+escapeHtml(message)+'</b><small>'+escapeHtml(hint)+'</small></div>';}
  function renderTimetable(rows){
    var box=el('schoolTimetableList');rows=rows.slice().sort(function(a,b){return Number(a.PERIO)-Number(b.PERIO);});
    if(!rows.length){box.innerHTML=emptyHtml('등록된 시간표가 없어요.','휴일이거나 NEIS에 아직 등록되지 않았을 수 있습니다.');return;}
    box.innerHTML=rows.map(function(row){return'<div class="school-list-row"><b>'+escapeHtml(row.PERIO||'')+'<small>교시</small></b><span>'+escapeHtml(row.ITRT_CNTNT||'수업 정보 없음')+'</span></div>';}).join('');
  }
  function mealItems(text){return String(text||'').split(/<br\s*\/?\s*>/i).map(function(item){return item.replace(/\s+/g,' ').trim();}).filter(Boolean);}
  function renderMeal(rows){
    var box=el('schoolMealList'),meta=el('schoolMealMeta');var meal=rows.find(function(row){return String(row.MMEAL_SC_CODE)==='2'||row.MMEAL_SC_NM==='중식';})||rows[0];
    if(!meal){meta.textContent='';box.innerHTML=emptyHtml('등록된 급식이 없어요.','급식이 없는 날이거나 NEIS에 아직 등록되지 않았을 수 있습니다.');return;}
    var items=mealItems(meal.DDISH_NM);meta.textContent=meal.MMEAL_SC_NM||'급식';box.innerHTML='<ul class="meal-dishes">'+items.map(function(item){return'<li>'+escapeHtml(item)+'</li>';}).join('')+'</ul>'+(meal.CAL_INFO?'<div class="meal-calorie">'+escapeHtml(meal.CAL_INFO)+'</div>':'');
  }
  function cacheKey(date){return[profile.atptCode,profile.schoolCode,date,profile.grade,profile.className].join('|');}
  function readCache(date){var all=readJson(CACHE_KEY,{}),item=all[cacheKey(date)];return item&&Date.now()-Number(item.savedAt||0)<21600000?item:null;}
  function saveCache(date,item){var all=readJson(CACHE_KEY,{});all[cacheKey(date)]={savedAt:Date.now(),timetable:item.timetable,meal:item.meal,timetableCut:item.timetableCut||null};var keys=Object.keys(all).sort(function(a,b){return Number(all[b].savedAt)-Number(all[a].savedAt);});keys.slice(20).forEach(function(key){delete all[key];});writeJson(CACHE_KEY,all);}
  function renderData(date,data,cached){el('schoolTimetableMeta').textContent=profile.grade+'학년 '+profile.className+'반';renderTimetable(data.timetable||[]);renderMeal(data.meal||[]);setStatus((Number(date.slice(5,7)))+'월 '+Number(date.slice(8,10))+'일 '+weekdayText(date)+' · NEIS 공개정보'+(cached?' · 저장된 조회 결과':'')+(data.timetableCut?' · NEIS 인증키가 없어 '+data.timetableCut.total+'교시 중 '+data.timetableCut.shown+'교시까지만 표시됩니다':''),cached?'cached':'ok');}
  async function refresh(force){
    if(!profile){setSettings(true);setStatus('학교를 먼저 설정해 주세요.','error');return;}
    var date=el('schoolTodayDate').value||formatDate(new Date()),cached=!force&&readCache(date);if(cached){renderData(date,cached,true);return;}
    var ownRequest=++requestNumber;loading();setStatus('NEIS 공개정보를 조회하고 있습니다.');
    try{var semester=semesterFor(date),common={ATPT_OFCDC_SC_CODE:profile.atptCode,SD_SCHUL_CODE:profile.schoolCode};var tableParams=Object.assign({},common,{AY:String(semester.year),SEM:String(semester.semester),ALL_TI_YMD:compactDate(date),GRADE:String(profile.grade),CLASS_NM:String(profile.className),pSize:'100'});var mealParams=Object.assign({},common,{MLSV_YMD:compactDate(date),pSize:'10'});lastTimetableCut=null;var values=await Promise.all([request(timetableDataset(profile.schoolKind),tableParams),request('mealServiceDietInfo',mealParams)]);if(ownRequest!==requestNumber)return;var data={timetable:values[0],meal:values[1],timetableCut:lastTimetableCut};saveCache(date,data);renderData(date,data,false);}catch(error){if(ownRequest!==requestNumber)return;el('schoolTimetableList').innerHTML=emptyHtml('시간표를 불러오지 못했습니다.','잠시 후 새로고침해 주세요.');el('schoolMealList').innerHTML=emptyHtml('급식을 불러오지 못했습니다.','네트워크 연결을 확인해 주세요.');setStatus('NEIS 조회에 실패했습니다. 잠시 후 다시 시도해 주세요.','error');}
  }
  async function searchSchool(){
    var input=el('schoolSearchInput'),query=input.value.trim(),status=el('schoolSearchStatus'),results=el('schoolSearchResults');if(query.length<2){status.textContent='학교 이름을 두 글자 이상 입력해 주세요.';input.focus();return;}
    status.textContent='학교를 검색하고 있습니다.';results.innerHTML='';el('schoolSearchBtn').disabled=true;
    try{var rows=await request('schoolInfo',{SCHUL_NM:query,pSize:'30'}),schools=rows.map(normalizeSchool);status.textContent=schools.length?'검색 결과 '+schools.length+'개':'검색 결과가 없습니다. 학교 이름을 확인해 주세요.';results.innerHTML=schools.map(function(s,i){return'<button type="button" class="school-search-result" data-school-index="'+i+'"><b>'+escapeHtml(s.schoolName)+'</b><span>'+escapeHtml([s.schoolKind,s.location,s.address].filter(Boolean).join(' · '))+'</span></button>';}).join('');results.querySelectorAll('[data-school-index]').forEach(function(button){button.addEventListener('click',function(){selectedSchool=schools[Number(this.dataset.schoolIndex)];results.querySelectorAll('.school-search-result').forEach(function(node){node.classList.remove('active');});this.classList.add('active');renderSelected();});});}catch(error){status.textContent='학교 검색에 실패했습니다. 네트워크 연결을 확인해 주세요.';}finally{el('schoolSearchBtn').disabled=false;}
  }
  function saveSettings(){
    var base=selectedSchool||profile,grade=Number(el('schoolGradeInput').value),className=Number(el('schoolClassInput').value);if(!base){showToast('검색 결과에서 학교를 선택해 주세요','error');return;}if(!grade||grade<1||grade>6||!className||className<1||className>30){showToast('학년과 반을 확인해 주세요','error');return;}
    profile=Object.assign({},base,{grade:String(grade),className:String(className)});writeJson(PROFILE_KEY,profile);selectedSchool=profile;renderProfile();setSettings(false);showToast('학교 설정을 저장했습니다');refresh(true);
  }
  function open(){
    profile=readJson(PROFILE_KEY,null);selectedSchool=profile;var date=el('schoolTodayDate');if(date)date.value=formatDate(new Date());if(profile){el('schoolGradeInput').value=profile.grade||1;el('schoolClassInput').value=profile.className||1;el('schoolSearchInput').value=profile.schoolName||'';}renderProfile();setSettings(!profile);openModal('schoolTodayModal');if(profile)refresh(false);else setStatus('학교 설정 후 조회할 수 있습니다.');
  }
  function init(){
    var settings=el('schoolSettingsBtn');if(settings)settings.addEventListener('click',function(){setSettings(el('schoolSettingsPanel').hidden);renderSelected();});
    var cancel=el('schoolSettingsCancelBtn');if(cancel)cancel.addEventListener('click',function(){setSettings(false);});
    var search=el('schoolSearchBtn');if(search)search.addEventListener('click',searchSchool);var input=el('schoolSearchInput');if(input)input.addEventListener('keydown',function(e){if(e.key==='Enter'){e.preventDefault();searchSchool();}});
    var save=el('schoolSettingsSaveBtn');if(save)save.addEventListener('click',saveSettings);var refreshButton=el('schoolTodayRefreshBtn');if(refreshButton)refreshButton.addEventListener('click',function(){refresh(true);});var date=el('schoolTodayDate');if(date)date.addEventListener('change',function(){refresh(false);});
  }
  return{init:init,open:open};
})();
// ===================================================================
// 일정관리 — 브라우저 로컬 저장 · 구글시트 붙여넣기 · ICS 내보내기
// ===================================================================
var ScheduleCalendar = (function () {
  var KEY = 'seatCalendar_v1';
  var PROP_KEY = 'seatCalendarProperties_v1';
  var PROP_STATE_KEY = 'seatCalendarProperties_v2';
  var DEFAULT_PROPERTIES = ['없음', '상담', '출장', '회의'];
  var REQUIRED_PROPERTIES = ['없음', '상담'];
  var events = [];
  var properties = [];
  var selectedDate = '';
  var monthCursor = null;
  var importDraft = [];
  var dayClickTimer = null;
  var viewMode = 'month';
  var wheelTotal = 0;
  var wheelLockUntil = 0;
  var pointerStart = null;
  var suppressDayClickUntil = 0;
  var propertyEditorMode = 'add';
  var propertyEditorOriginal = '';
  var calendarDragState = null;
  var calendarPointerHold = null;
  var calendarLongPressTimer = null;
  var calendarDropTarget = null;
  var calendarDragGhost = null;

  function el(id) { return document.getElementById(id); }
  function todayKey() { return formatDate(new Date()); }
  function dateFromKey(key) {
    var p = String(key || '').split('-').map(Number);
    return new Date(p[0], (p[1] || 1) - 1, p[2] || 1);
  }
  function addDays(key, amount) {
    var d = dateFromKey(key); d.setDate(d.getDate() + amount); return formatDate(d);
  }
  function unique(values) {
    var seen = {};
    return values.map(function (x) { return String(x || '').trim(); }).filter(function (x) {
      if (!x || seen[x]) return false; seen[x] = true; return true;
    });
  }
  function eventProperty(e) { return String(e.property || e.type || '없음').trim() || '없음'; }
  function load() {
    try {
      var raw = JSON.parse(localStorage.getItem(KEY) || '[]');
      events = Array.isArray(raw) ? raw.filter(function (e) { return e && e.start && e.title; }) : [];
    } catch (e) { events = []; }
    try {
      var savedState = JSON.parse(localStorage.getItem(PROP_STATE_KEY) || 'null');
      if (Array.isArray(savedState)) {
        properties = unique(REQUIRED_PROPERTIES.concat(savedState).concat(events.map(eventProperty)));
      } else {
        var saved = JSON.parse(localStorage.getItem(PROP_KEY) || '[]');
        properties = unique(DEFAULT_PROPERTIES.concat(Array.isArray(saved) ? saved : []).concat(events.map(eventProperty)));
      }
    } catch (e) { properties = DEFAULT_PROPERTIES.slice(); }
  }
  function save() {
    try { localStorage.setItem(KEY, JSON.stringify(events)); }
    catch (e) { showToast('일정 저장 공간을 사용할 수 없습니다', 'error'); }
  }
  function saveProperties() {
    try {
      localStorage.setItem(PROP_STATE_KEY, JSON.stringify(properties));
      localStorage.setItem(PROP_KEY, JSON.stringify(properties.filter(function (x) { return DEFAULT_PROPERTIES.indexOf(x) < 0; })));
    } catch (e) { showToast('일정 속성을 저장할 수 없습니다', 'error'); }
  }
  function propertyClass(property) {
    var index = properties.indexOf(property);
    if (property === '없음') return 'type-none';
    return 'type-' + (index < 0 ? 0 : Math.max(0,index - 1) % 5);
  }
  function propertyLocked(property) { return REQUIRED_PROPERTIES.indexOf(property) >= 0; }
  function syncPropertyActions() {
    var select=el('calPropertyInput'), editButton=el('calEditPropertyBtn'), deleteButton=el('calDeletePropertyBtn');
    var locked=!select||propertyLocked(select.value);
    if(editButton)editButton.disabled=locked;
    if(deleteButton)deleteButton.disabled=locked;
  }
  function renderPropertyOptions(selected) {
    var select = el('calPropertyInput'); if (!select) return;
    select.innerHTML = properties.map(function (property) {
      return '<option value="' + escapeHtml(property) + '"' + (property === selected ? ' selected' : '') + '>' + escapeHtml(property === '없음' ? '없음 (표시 안 함)' : property) + '</option>';
    }).join('');
    if (selected && properties.indexOf(selected) < 0) {
      properties.push(selected); saveProperties(); renderPropertyOptions(selected); return;
    }
    syncPropertyActions();
  }
  function showPropertyAdder(mode) {
    var box=el('calPropertyAdder'), input=el('calNewPropertyInput'), saveButton=el('calPropertySaveBtn'), select=el('calPropertyInput'); if(!box||!input)return;
    propertyEditorMode=mode==='edit'?'edit':'add';propertyEditorOriginal=propertyEditorMode==='edit'&&select?select.value:'';
    if(propertyEditorMode==='edit'&&propertyLocked(propertyEditorOriginal)){showToast('없음과 상담 속성은 기능 연결을 위해 이름을 유지합니다','error');return;}
    box.hidden=false;input.value=propertyEditorMode==='edit'?propertyEditorOriginal:'';if(saveButton)saveButton.textContent=propertyEditorMode==='edit'?'수정':'추가';
    setTimeout(function(){input.focus();input.select();},20);
  }
  function hidePropertyAdder() {
    var box=el('calPropertyAdder'),saveButton=el('calPropertySaveBtn');if(box)box.hidden=true;if(saveButton)saveButton.textContent='추가';propertyEditorMode='add';propertyEditorOriginal='';
  }
  function confirmProperty() {
    var input=el('calNewPropertyInput'); if(!input)return;
    var value=input.value.trim().replace(/[\[\]<>]/g, '');
    if(!value){showToast('속성 이름을 입력해 주세요','error');input.focus();return;}
    if(value.length>12){showToast('속성은 12자 이내로 입력해 주세요','error');input.focus();return;}
    if(propertyEditorMode==='edit'){
      if(value!==propertyEditorOriginal&&properties.indexOf(value)>=0){showToast('같은 이름의 속성이 이미 있습니다','error');input.focus();return;}
      var index=properties.indexOf(propertyEditorOriginal);if(index>=0)properties[index]=value;
      events.forEach(function(item){if(eventProperty(item)===propertyEditorOriginal){item.property=value;item.updatedAt=Date.now();}});save();saveProperties();if(typeof GoogleCalendarSync!=='undefined')GoogleCalendarSync.onLocalChanged();
      renderPropertyOptions(value);syncConsultFields();render();hidePropertyAdder();showToast('속성 이름을 “'+value+'”(으)로 수정했습니다');return;
    }
    if(properties.indexOf(value)>=0){renderPropertyOptions(value);syncConsultFields();hidePropertyAdder();showToast('이미 등록된 속성입니다');return;}
    properties.push(value);saveProperties();renderPropertyOptions(value);syncConsultFields();hidePropertyAdder();showToast('“'+value+'” 속성을 추가했습니다');
  }
  function deleteProperty() {
    var select=el('calPropertyInput');if(!select)return;var value=select.value;
    if(propertyLocked(value)){showToast('없음과 상담 속성은 삭제할 수 없습니다','error');return;}
    var used=events.filter(function(item){return eventProperty(item)===value;}).length;
    var message='“'+value+'” 속성을 삭제할까요?'+(used?'\n사용 중인 일정 '+used+'개는 속성 없음으로 바뀝니다.':'');
    if(!confirm(message))return;
    properties=properties.filter(function(property){return property!==value;});
    events.forEach(function(item){if(eventProperty(item)===value){item.property='없음';item.updatedAt=Date.now();}});save();saveProperties();if(typeof GoogleCalendarSync!=='undefined')GoogleCalendarSync.onLocalChanged();renderPropertyOptions('없음');syncConsultFields();render();hidePropertyAdder();showToast('“'+value+'” 속성을 삭제했습니다');
  }
  function eventKey(e) {
    return [e.start, e.end || e.start, String(e.title || '').trim(), eventProperty(e)].join('|').toLowerCase();
  }
  function eventsOn(key) {
    return events.filter(function (e) { return e.start <= key && (e.end || e.start) >= key; })
      .sort(function (a,b) { return a.start.localeCompare(b.start) || a.title.localeCompare(b.title); });
  }
  function calendarDragAttributes(item, displayDate) {
    return ' data-cal-event-id="'+escapeHtml(item.id)+'" data-cal-event-date="'+escapeHtml(displayDate)+'" title="드래그해서 일정 이동" aria-label="'+escapeHtml(item.title)+' 일정, 드래그해서 날짜 이동"';
  }
  function monthTitle(d) { return d.getFullYear() + '년 ' + (d.getMonth() + 1) + '월'; }
  function selectedTitle(key) {
    var d = dateFromKey(key), day = ['일','월','화','수','목','금','토'][d.getDay()];
    return (d.getMonth()+1) + '월 ' + d.getDate() + '일 (' + day + ')';
  }
  function renderGrid() {
    var grid=el('calGrid'); if(!grid||!monthCursor)return;
    var year=monthCursor.getFullYear(), month=monthCursor.getMonth();
    var first=new Date(year,month,1), start=new Date(year,month,1-first.getDay());
    var daysInMonth=new Date(year,month+1,0).getDate();
    var weekCount=(first.getDay()+daysInMonth>35)?6:5, cellCount=weekCount*7;
    grid.classList.toggle('has-sixth-week',weekCount===6);
    grid.dataset.weeks=String(weekCount);
    grid.setAttribute('aria-label',monthTitle(monthCursor)+' 월간 달력, '+weekCount+'주');
    grid.scrollTop=0;
    var html='';
    for(var i=0;i<cellCount;i++){
      var d=new Date(start); d.setDate(start.getDate()+i); var key=formatDate(d), list=eventsOn(key);
      var cls='cal-day'+(d.getMonth()!==month?' other-month':'')+(key===todayKey()?' today':'')+(key===selectedDate?' selected':'');
      var dayText=d.getDate()===1?((d.getMonth()+1)+'/'+d.getDate()):String(d.getDate());
      html+='<button type="button" class="'+cls+'" data-date="'+key+'"><span class="cal-day-no">'+dayText+'</span><span class="cal-day-events">'+
        list.slice(0,2).map(function(e){var prop=eventProperty(e);return '<em class="'+propertyClass(prop)+' cal-event-drag"'+calendarDragAttributes(e,key)+'>'+(prop==='없음'?'':'<b>['+escapeHtml(prop)+']</b> ')+escapeHtml(e.title)+'</em>';}).join('')+
        (list.length>2?'<small>+'+(list.length-2)+'</small>':'')+'</span></button>';
    }
    grid.innerHTML=html;
    grid.querySelectorAll('.cal-event-drag').forEach(function(card){
      card.addEventListener('click',function(e){
        e.preventDefault();e.stopPropagation();
        if(Date.now()<suppressDayClickUntil)return;
        clearTimeout(dayClickTimer);editEvent(this.dataset.calEventId);
      });
    });
    var selectedCell=grid.querySelector('.cal-day.selected');
    if(selectedCell&&selectedCell.offsetTop>=grid.clientHeight){
      grid.scrollTop=Math.max(0,selectedCell.offsetTop-Math.floor(grid.clientHeight*0.8));
    }
    grid.querySelectorAll('.cal-day').forEach(function(b){
      b.addEventListener('click',function(){
        if(Date.now()<suppressDayClickUntil)return;
        var key=this.dataset.date;clearTimeout(dayClickTimer);dayClickTimer=setTimeout(function(){
          selectedDate=key;var d=dateFromKey(selectedDate);
          if(d.getMonth()!==monthCursor.getMonth()||d.getFullYear()!==monthCursor.getFullYear())monthCursor=new Date(d.getFullYear(),d.getMonth(),1);
          resetForm();render();
        },220);
      });
      b.addEventListener('dblclick',function(e){
        e.preventDefault();clearTimeout(dayClickTimer);selectedDate=this.dataset.date;var d=dateFromKey(selectedDate);
        if(d.getMonth()!==monthCursor.getMonth()||d.getFullYear()!==monthCursor.getFullYear())monthCursor=new Date(d.getFullYear(),d.getMonth(),1);
        render();openNewEditor();
      });
    });
  }
  function renderList() {
    var label=el('calSelectedLabel'), list=el('calEventList'); if(!label||!list)return;
    label.textContent=selectedTitle(selectedDate); var rows=eventsOn(selectedDate);
    if(!rows.length){list.innerHTML='<p class="cal-empty">등록된 일정이 없습니다.</p>';return;}
    list.innerHTML=rows.map(function(e){
      var range=e.start===(e.end||e.start)?e.start:e.start+' ~ '+(e.end||e.start), prop=eventProperty(e);
      return '<article class="cal-event-item '+propertyClass(prop)+'"><div><b>'+(prop==='없음'?'':'<span class="cal-prop-badge">'+escapeHtml(prop)+'</span>')+escapeHtml(e.title)+'</b><span>'+range+'</span>'+
        (e.target?'<small>대상 · '+escapeHtml(e.target)+(e.studentName?' · '+escapeHtml(e.studentName):'')+'</small>':'')+
        (e.consultMethod||e.status?'<small>'+escapeHtml([e.consultMethod,e.status].filter(Boolean).join(' · '))+'</small>':'')+
        (e.memo?'<small>'+escapeHtml(e.memo)+'</small>':'')+(e.followUp?'<small>후속 · '+escapeHtml(e.followUp)+'</small>':'')+'</div>'+
        '<div class="cal-event-actions"><button type="button" data-edit="'+e.id+'">수정</button><button type="button" data-delete="'+e.id+'">삭제</button></div></article>';
    }).join('');
    list.querySelectorAll('[data-edit]').forEach(function(b){b.addEventListener('click',function(){editEvent(this.dataset.edit);});});
    list.querySelectorAll('[data-delete]').forEach(function(b){b.addEventListener('click',function(){removeEvent(this.dataset.delete);});});
  }

  function bindEventActions(container) {
    if(!container)return;
    container.querySelectorAll('[data-edit]').forEach(function(b){b.addEventListener('click',function(e){e.stopPropagation();editEvent(this.dataset.edit);});});
    container.querySelectorAll('[data-delete]').forEach(function(b){b.addEventListener('click',function(e){e.stopPropagation();removeEvent(this.dataset.delete);});});
  }
  function compactEventHtml(e, dragDate) {
    var prop=eventProperty(e),dragClass=dragDate?' cal-event-drag':'',dragAttrs=dragDate?calendarDragAttributes(e,dragDate):'';
    return '<article class="cal-event-item '+propertyClass(prop)+dragClass+'"'+dragAttrs+'><div><b>'+(prop==='없음'?'':'<span class="cal-prop-badge">'+escapeHtml(prop)+'</span>')+escapeHtml(e.title)+'</b>'+
      (e.start!==(e.end||e.start)?'<span>'+escapeHtml(e.start+' ~ '+(e.end||e.start))+'</span>':'')+
      (e.target?'<small>대상 · '+escapeHtml(e.target)+(e.studentName?' · '+escapeHtml(e.studentName):'')+'</small>':'')+
      (e.consultMethod||e.status?'<small>'+escapeHtml([e.consultMethod,e.status].filter(Boolean).join(' · '))+'</small>':'')+
      (e.memo?'<small>'+escapeHtml(e.memo)+'</small>':'')+(e.followUp?'<small>후속 · '+escapeHtml(e.followUp)+'</small>':'')+'</div>'+
      '<div class="cal-event-actions"><button type="button" data-edit="'+e.id+'">수정</button><button type="button" data-delete="'+e.id+'">삭제</button></div></article>';
  }
  function weekStartKey(key) {
    var d=dateFromKey(key);d.setDate(d.getDate()-d.getDay());return formatDate(d);
  }
  function renderWeek() {
    var box=el('calWeekGrid');if(!box)return;
    var start=weekStartKey(selectedDate),days=['일','월','화','수','목','금','토'],html='';
    for(var i=0;i<7;i++){
      var key=addDays(start,i),d=dateFromKey(key),rows=eventsOn(key);
      var dateText=(d.getMonth()+1)+'/'+d.getDate();
      html+='<section class="cal-week-day'+(key===todayKey()?' today':'')+(key===selectedDate?' selected':'')+'" data-week-date="'+key+'">'+
        '<button type="button" class="cal-week-day-head" data-select-date="'+key+'"><span>'+days[i]+'</span><b>'+dateText+'</b></button>'+
        '<div class="cal-week-events">'+(rows.length?rows.map(function(item){return compactEventHtml(item,key);}).join(''):'<p class="cal-empty">일정 없음</p>')+'</div></section>';
    }
    box.innerHTML=html;bindEventActions(box);
    box.querySelectorAll('[data-select-date]').forEach(function(button){button.addEventListener('click',function(){selectedDate=this.dataset.selectDate;monthCursor=new Date(dateFromKey(selectedDate).getFullYear(),dateFromKey(selectedDate).getMonth(),1);render();});});
    box.querySelectorAll('[data-week-date]').forEach(function(day){day.addEventListener('dblclick',function(e){if(e.target.closest('.cal-event-item'))return;selectedDate=this.dataset.weekDate;resetForm();openNewEditor();});});
  }
  function renderTodayView() {
    var label=el('calTodayLabel'),list=el('calTodayList');if(!label||!list)return;
    label.textContent=selectedTitle(selectedDate);var rows=eventsOn(selectedDate);
    list.innerHTML=rows.length?rows.map(function(item){return compactEventHtml(item,'');}).join(''):'<p class="cal-empty cal-today-empty">등록된 일정이 없습니다.<small>오른쪽 위 버튼이나 빈 공간을 더블클릭해 일정을 추가하세요.</small></p>';
    bindEventActions(list);
  }
  function viewTitle() {
    if(viewMode==='month')return monthTitle(monthCursor);
    if(viewMode==='day')return selectedTitle(selectedDate);
    var start=weekStartKey(selectedDate),end=addDays(start,6),a=dateFromKey(start),b=dateFromKey(end);
    return a.getFullYear()+'년 '+(a.getMonth()+1)+'월 '+a.getDate()+'일 – '+(b.getMonth()+1)+'월 '+b.getDate()+'일';
  }
  function syncView() {
    var monthLayout=el('calLayout'),weekPane=el('calWeekPane'),todayPane=el('calTodayPane');
    if(monthLayout)monthLayout.hidden=viewMode!=='month';
    if(weekPane)weekPane.hidden=viewMode!=='week';
    if(todayPane)todayPane.hidden=viewMode!=='day';
    document.querySelectorAll('[data-cal-view]').forEach(function(button){var active=button.dataset.calView===viewMode;button.classList.toggle('active',active);button.setAttribute('aria-selected',active?'true':'false');});
    var label=el('calMonthLabel');if(label)label.textContent=viewTitle();
    var prev=el('calPrevBtn'),next=el('calNextBtn'),unit=viewMode==='month'?'달':viewMode==='week'?'주':'날';
    if(prev)prev.setAttribute('aria-label','이전 '+unit);if(next)next.setAttribute('aria-label','다음 '+unit);
  }
  function setView(mode) {
    viewMode=mode==='week'||mode==='day'?mode:'month';
    if(viewMode==='day'){selectedDate=todayKey();var today=dateFromKey(selectedDate);monthCursor=new Date(today.getFullYear(),today.getMonth(),1);resetForm();}
    render();
  }
  function shiftMonth(amount) {
    var current=dateFromKey(selectedDate),day=current.getDate();
    var target=new Date(current.getFullYear(),current.getMonth()+amount,1);
    var last=new Date(target.getFullYear(),target.getMonth()+1,0).getDate();target.setDate(Math.min(day,last));
    selectedDate=formatDate(target);monthCursor=new Date(target.getFullYear(),target.getMonth(),1);
  }
  function navigate(amount,animate) {
    if(viewMode==='month')shiftMonth(amount);else selectedDate=addDays(selectedDate,viewMode==='week'?amount*7:amount);
    var d=dateFromKey(selectedDate);monthCursor=new Date(d.getFullYear(),d.getMonth(),1);resetForm();render();
    var pane=viewMode==='month'?el('calMonthPane'):viewMode==='week'?el('calWeekPane'):el('calTodayPane');
    if(pane&&animate){pane.classList.remove('cal-slide-prev','cal-slide-next');void pane.offsetWidth;pane.classList.add(amount>0?'cal-slide-next':'cal-slide-prev');}
  }
  function onCalendarWheel(e) {
    if(e.ctrlKey||e.metaKey||el('calEditorOverlay')&&!el('calEditorOverlay').hidden)return;
    var delta=Math.abs(e.deltaX)>Math.abs(e.deltaY)?e.deltaX:e.deltaY;if(Math.abs(delta)<2)return;
    e.preventDefault();
    if(Date.now()<wheelLockUntil)return;wheelTotal+=delta;
    if(Math.abs(wheelTotal)>=36){var direction=wheelTotal>0?1:-1;wheelTotal=0;wheelLockUntil=Date.now()+360;navigate(direction,true);}
  }
  function onPointerDown(e) {
    if(e.button!==undefined&&e.button!==0)return;if(e.target.closest('.cal-event-drag,.cal-event-actions,.cal-nav,.cal-view-tabs'))return;
    pointerStart={x:e.clientX,y:e.clientY,id:e.pointerId};
  }
  function onPointerUp(e) {
    if(!pointerStart||pointerStart.id!==e.pointerId){pointerStart=null;return;}
    var dx=e.clientX-pointerStart.x,dy=e.clientY-pointerStart.y;pointerStart=null;
    if(Math.abs(dx)>64&&Math.abs(dx)>Math.abs(dy)*1.25){suppressDayClickUntil=Date.now()+450;navigate(dx<0?1:-1,true);}
  }
  function initGestures() {
    ['calMonthPane','calWeekPane','calTodayPane'].forEach(function(id){var node=el(id);if(!node)return;node.addEventListener('wheel',onCalendarWheel,{passive:false});node.addEventListener('pointerdown',onPointerDown);node.addEventListener('pointerup',onPointerUp);node.addEventListener('pointercancel',function(){pointerStart=null;});});
  }

  function calendarDayDiff(fromKey, toKey) {
    var from=dateFromKey(fromKey),to=dateFromKey(toKey);
    return Math.round((Date.UTC(to.getFullYear(),to.getMonth(),to.getDate())-Date.UTC(from.getFullYear(),from.getMonth(),from.getDate()))/86400000);
  }
  function moveCalendarEvent(eventId, sourceDate, targetDate) {
    var item=events.find(function(entry){return String(entry.id)===String(eventId);});
    if(!item||!sourceDate||!targetDate)return false;
    var delta=calendarDayDiff(sourceDate,targetDate);
    if(!delta){showToast('같은 날짜입니다');return false;}
    var previous={start:item.start,end:item.end||item.start,updatedAt:item.updatedAt};
    item.start=addDays(previous.start,delta);item.end=addDays(previous.end,delta);
    item.updatedAt=Math.max(Date.now(),Number(item.updatedAt||0)+1);
    var movedUpdatedAt=item.updatedAt,title=item.title;
    selectedDate=targetDate;var target=dateFromKey(targetDate);monthCursor=new Date(target.getFullYear(),target.getMonth(),1);
    save();render();if(typeof GoogleCalendarSync!=='undefined')GoogleCalendarSync.onLocalChanged();
    showToast('“'+title+'” 일정을 '+selectedTitle(targetDate)+'로 이동했습니다','', '실행 취소', function(){
      var current=events.find(function(entry){return String(entry.id)===String(eventId);});
      if(!current)return;
      current.start=previous.start;current.end=previous.end;
      current.updatedAt=Math.max(Date.now(),Number(current.updatedAt||movedUpdatedAt)+1);
      selectedDate=previous.start;var original=dateFromKey(previous.start);monthCursor=new Date(original.getFullYear(),original.getMonth(),1);
      save();render();if(typeof GoogleCalendarSync!=='undefined')GoogleCalendarSync.onLocalChanged();showToast('일정 이동을 취소했습니다');
    });
    return true;
  }
  function calendarDropDate(node) {
    if(!node||!node.closest)return'';
    var target=node.closest('[data-date],[data-week-date]');
    return target?(target.dataset.date||target.dataset.weekDate||''):'';
  }
  function clearCalendarDropTarget() {
    if(calendarDropTarget)calendarDropTarget.classList.remove('cal-drop-target');
    calendarDropTarget=null;
  }
  function setCalendarDropTargetAt(x,y) {
    var node=document.elementFromPoint(x,y),target=node&&node.closest?node.closest('.cal-day[data-date],.cal-week-day[data-week-date]'):null;
    if(target===calendarDropTarget)return;
    clearCalendarDropTarget();calendarDropTarget=target;
    if(calendarDropTarget)calendarDropTarget.classList.add('cal-drop-target');
  }
  function makeCalendarDragGhost(source,x,y) {
    if(calendarDragGhost)calendarDragGhost.remove();
    calendarDragGhost=document.createElement('div');calendarDragGhost.className='cal-drag-ghost';
    calendarDragGhost.textContent=(source.textContent||'일정 이동').trim();document.body.appendChild(calendarDragGhost);
    moveCalendarDragGhost(x,y);
  }
  function moveCalendarDragGhost(x,y) {
    if(!calendarDragGhost)return;
    calendarDragGhost.style.transform='translate3d('+(x+14)+'px,'+(y+14)+'px,0)';
  }
  function clearCalendarLongPress() {
    clearTimeout(calendarLongPressTimer);calendarLongPressTimer=null;
  }
  function cleanupCalendarDrag() {
    clearCalendarLongPress();clearCalendarDropTarget();
    document.querySelectorAll('.cal-event-drag.is-dragging').forEach(function(node){node.classList.remove('is-dragging');});
    if(calendarDragGhost){calendarDragGhost.remove();calendarDragGhost=null;}
    document.body.classList.remove('calendar-event-dragging');
    if(calendarPointerHold&&calendarPointerHold.source&&calendarPointerHold.pointerId!==undefined){try{calendarPointerHold.source.releasePointerCapture(calendarPointerHold.pointerId);}catch(e){}}
    calendarPointerHold=null;calendarDragState=null;
  }
  function onCalendarNativeDragStart(e) {
    var source=e.target.closest('.cal-event-drag');if(!source||e.target.closest('.cal-event-actions'))return;
    calendarDragState={eventId:source.dataset.calEventId,sourceDate:source.dataset.calEventDate,source:source,mode:'native'};
    source.classList.add('is-dragging');document.body.classList.add('calendar-event-dragging');suppressDayClickUntil=Date.now()+700;
    if(e.dataTransfer){e.dataTransfer.effectAllowed='move';e.dataTransfer.setData('text/plain',source.dataset.calEventId||'');}
  }
  function onCalendarNativeDragOver(e) {
    if(!calendarDragState||calendarDragState.mode!=='native')return;
    var date=calendarDropDate(e.target);if(!date)return;e.preventDefault();if(e.dataTransfer)e.dataTransfer.dropEffect='move';
    setCalendarDropTargetAt(e.clientX,e.clientY);
  }
  function onCalendarNativeDrop(e) {
    if(!calendarDragState||calendarDragState.mode!=='native')return;
    var eventId=calendarDragState.eventId,sourceDate=calendarDragState.sourceDate,targetDate=calendarDropDate(e.target);
    if(targetDate)e.preventDefault();cleanupCalendarDrag();if(targetDate)moveCalendarEvent(eventId,sourceDate,targetDate);
  }
  function activateCalendarPointerDrag() {
    if(!calendarPointerHold||calendarDragState)return;
    var source=calendarPointerHold.source;
    calendarDragState={eventId:calendarPointerHold.eventId,sourceDate:calendarPointerHold.sourceDate,source:source,mode:'pointer'};
    source.classList.add('is-dragging');document.body.classList.add('calendar-event-dragging');suppressDayClickUntil=Date.now()+900;
    makeCalendarDragGhost(source,calendarPointerHold.lastX,calendarPointerHold.lastY);setCalendarDropTargetAt(calendarPointerHold.lastX,calendarPointerHold.lastY);
    if(calendarPointerHold.pointerType!=='mouse'&&navigator.vibrate)try{navigator.vibrate(18);}catch(err){}
  }
  function onCalendarPointerDown(e) {
    if(e.isPrimary===false)return;
    var source=e.target.closest('.cal-event-drag');if(!source||e.target.closest('.cal-event-actions'))return;
    cleanupCalendarDrag();
    calendarPointerHold={source:source,eventId:source.dataset.calEventId,sourceDate:source.dataset.calEventDate,pointerId:e.pointerId,pointerType:e.pointerType||'mouse',startX:e.clientX,startY:e.clientY,lastX:e.clientX,lastY:e.clientY};
    try{source.setPointerCapture(e.pointerId);}catch(err){}
    if(calendarPointerHold.pointerType!=='mouse')calendarLongPressTimer=setTimeout(activateCalendarPointerDrag,430);
  }
  function onCalendarPointerMove(e) {
    if(!calendarPointerHold||calendarPointerHold.pointerId!==e.pointerId)return;
    calendarPointerHold.lastX=e.clientX;calendarPointerHold.lastY=e.clientY;
    if(!calendarDragState){
      var movedX=Math.abs(e.clientX-calendarPointerHold.startX),movedY=Math.abs(e.clientY-calendarPointerHold.startY);
      if(calendarPointerHold.pointerType==='mouse'&&(movedX>4||movedY>4))activateCalendarPointerDrag();
      else if(calendarPointerHold.pointerType!=='mouse'&&(movedX>9||movedY>9)){cleanupCalendarDrag();return;}
      else return;
    }
    e.preventDefault();moveCalendarDragGhost(e.clientX,e.clientY);setCalendarDropTargetAt(e.clientX,e.clientY);
  }
  function onCalendarPointerUp(e) {
    if(!calendarPointerHold||calendarPointerHold.pointerId!==e.pointerId)return;
    if(!calendarDragState){cleanupCalendarDrag();return;}
    e.preventDefault();var eventId=calendarDragState.eventId,sourceDate=calendarDragState.sourceDate;
    var targetDate=calendarDropTarget?(calendarDropTarget.dataset.date||calendarDropTarget.dataset.weekDate||''):calendarDropDate(document.elementFromPoint(e.clientX,e.clientY));
    cleanupCalendarDrag();if(targetDate)moveCalendarEvent(eventId,sourceDate,targetDate);
  }
  function initEventDragging() {
    ['calGrid','calWeekGrid'].forEach(function(id){var container=el(id);if(!container)return;
      container.addEventListener('dragstart',onCalendarNativeDragStart);container.addEventListener('dragover',onCalendarNativeDragOver);container.addEventListener('drop',onCalendarNativeDrop);container.addEventListener('dragend',cleanupCalendarDrag);
      container.addEventListener('pointerdown',onCalendarPointerDown);container.addEventListener('pointermove',onCalendarPointerMove,{passive:false});container.addEventListener('pointerup',onCalendarPointerUp);container.addEventListener('pointercancel',cleanupCalendarDrag);
      container.addEventListener('contextmenu',function(e){if(e.target.closest('.cal-event-drag'))e.preventDefault();});
    });
  }

  function populateStudents(selectedId) {
    var select=el('calStudentInput'); if(!select)return;
    select.innerHTML='<option value="">학생 선택</option>'+students.map(function(student){
      return '<option value="'+escapeHtml(student.id)+'"'+(String(student.id)===String(selectedId)?' selected':'')+'>'+escapeHtml(student.id+'번 '+student.name)+'</option>';
    }).join('');
  }
  function syncConsultFields() {
    var property=el('calPropertyInput').value, target=el('calTargetInput'), fields=el('calConsultFields'), generalLabel=el('calGeneralTargetLabel'), consultTarget=el('calConsultTargetInput');
    var isConsult=property==='상담';if(fields)fields.hidden=!isConsult;if(generalLabel)generalLabel.hidden=isConsult;
    if(target){if(isConsult)target.value=consultTarget&&consultTarget.value==='학부모'?'학부모':'학생';else if(target.value==='학생'||target.value==='학부모')target.value='전체';}
    syncPropertyActions();
  }
  function showEditor() {
    var overlay=el('calEditorOverlay'),title=el('calEditorTitle');if(!overlay)return;
    if(title)title.textContent=el('calEditId').value?'일정 수정':selectedTitle(selectedDate)+' 일정 입력';
    overlay.hidden=false;document.body.classList.add('calendar-editor-open');setTimeout(function(){el('calTitleInput').focus();},30);
  }
  function hideEditor() {
    var overlay=el('calEditorOverlay');if(overlay)overlay.hidden=true;
    document.body.classList.remove('calendar-editor-open');hidePropertyAdder();
  }
  function openNewEditor(){resetForm();showEditor();}
  function render(){syncView();renderGrid();renderList();renderWeek();renderTodayView();}
  function resetForm(){
    el('calEditId').value='';el('calTitleInput').value='';el('calStartInput').value=selectedDate||todayKey();el('calEndInput').value='';
    renderPropertyOptions('없음');el('calTargetInput').value='전체';el('calMemoInput').value='';
    el('calConsultTargetInput').value='학생';el('calConsultMethodInput').value='학생 상담';el('calStatusInput').value='예정';el('calFollowUpInput').value='';
    populateStudents('');syncConsultFields();el('calSaveBtn').textContent='일정 저장';
    var deleteButton=el('calDeleteEventBtn');if(deleteButton)deleteButton.hidden=true;
  }
  function editEvent(id){
    var item=events.find(function(e){return e.id===id;});if(!item)return;
    el('calEditId').value=item.id;el('calTitleInput').value=item.title;el('calStartInput').value=item.start;
    el('calEndInput').value=item.end&&item.end!==item.start?item.end:'';renderPropertyOptions(eventProperty(item));
    el('calTargetInput').value=item.target||(item.studentId?'학생':'전체');el('calConsultTargetInput').value=item.target==='학부모'?'학부모':'학생';el('calMemoInput').value=item.memo||'';
    el('calConsultMethodInput').value=item.consultMethod||'학생 상담';el('calStatusInput').value=item.status||'예정';
    el('calFollowUpInput').value=item.followUp||'';populateStudents(item.studentId||'');syncConsultFields();
    el('calSaveBtn').textContent='수정 저장';var deleteButton=el('calDeleteEventBtn');if(deleteButton)deleteButton.hidden=false;showEditor();
  }
  function removeEvent(id){
    var item=events.find(function(e){return e.id===id;});if(!item||!confirm('"'+item.title+'" 일정을 삭제할까요?'))return false;
    if(typeof GoogleCalendarSync!=='undefined')GoogleCalendarSync.noteLocalDelete(item);events=events.filter(function(e){return e.id!==id;});save();resetForm();render();if(typeof GoogleCalendarSync!=='undefined')GoogleCalendarSync.onLocalChanged();showToast('일정을 삭제했습니다');return true;
  }
  function submit(e){
    e.preventDefault();var title=el('calTitleInput').value.trim(),start=el('calStartInput').value,end=el('calEndInput').value||start;
    if(!title||!start){showToast('날짜와 일정명을 입력해 주세요','error');return;}if(end<start){showToast('종료일은 시작일보다 빠를 수 없습니다','error');return;}
    var property=el('calPropertyInput').value||'없음';
    var previous=events.find(function(x){return x.id===el('calEditId').value;});
    var item={id:el('calEditId').value||generateId(),title:title,start:start,end:end,property:property,target:el('calTargetInput').value||'전체',
      memo:el('calMemoInput').value.trim(),studentId:'',studentName:'',consultMethod:'',status:'',followUp:'',createdAt:previous&&previous.createdAt||Date.now(),updatedAt:Math.max(Date.now(),Number(previous&&previous.updatedAt||0)+1)};
    if(property==='상담'){
      var studentSelect=el('calStudentInput');item.studentId=studentSelect.value||'';item.studentName=studentSelect.selectedOptions.length?studentSelect.selectedOptions[0].textContent:'';
      item.target=el('calConsultTargetInput').value||'학생';item.consultMethod=el('calConsultMethodInput').value;item.status=el('calStatusInput').value;item.followUp=el('calFollowUpInput').value.trim();
      if(!item.studentId){showToast('상담과 연결할 학생을 선택해 주세요','error');return;}
    }
    var idx=events.findIndex(function(x){return x.id===item.id;});if(idx>=0)events[idx]=item;else events.push(item);
    selectedDate=start;var d=dateFromKey(start);monthCursor=new Date(d.getFullYear(),d.getMonth(),1);save();render();hideEditor();resetForm();if(typeof GoogleCalendarSync!=='undefined')GoogleCalendarSync.onLocalChanged();showToast(idx>=0?'일정을 수정했습니다':'일정을 저장했습니다');
  }
  function normalizeDate(value){
    var x=String(value||'').trim().replace(/[년.\/]/g,'-').replace(/월/g,'-').replace(/일/g,'').replace(/\s+/g,'').replace(/-+/g,'-').replace(/-$/,'');
    var m=x.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);if(!m){m=x.match(/^(\d{1,2})-(\d{1,2})$/);if(m)m=[m[0],String((monthCursor||new Date()).getFullYear()),m[1],m[2]];}
    if(!m)return'';var y=Number(m[1]),mo=Number(m[2]),da=Number(m[3]),d=new Date(y,mo-1,da);return d.getFullYear()===y&&d.getMonth()===mo-1&&d.getDate()===da?formatDate(d):'';
  }
  function previewPaste(){
    var text=el('calPasteInput').value.trim(),box=el('calImportPreview');importDraft=[];
    if(!text){box.innerHTML='<p class="cal-import-error">붙여넣은 내용이 없습니다.</p>';el('calImportBtn').disabled=true;return;}
    var rows=text.split(/\r?\n/).filter(function(r){return r.trim();}).map(function(r){return r.split('\t');});
    var headers=rows[0].map(function(h){return h.trim().replace(/\s/g,'');});
    function col(names){for(var i=0;i<names.length;i++){var x=headers.indexOf(names[i]);if(x>=0)return x;}return-1;}
    var map={start:col(['날짜','시작일','일자']),end:col(['종료일','마침일']),title:col(['일정명','제목','일정']),property:col(['속성','유형','구분','분류']),memo:col(['메모','내용','비고'])};
    var hasHeader=map.start>=0||map.title>=0;if(!hasHeader)map={start:0,end:1,title:2,property:3,memo:4};var dataRows=hasHeader?rows.slice(1):rows,errors=[];
    dataRows.forEach(function(r,i){var start=normalizeDate(r[map.start]),title=String(r[map.title]||'').trim(),end=map.end>=0?normalizeDate(r[map.end]):'';
      if(!start||!title){errors.push((i+1+(hasHeader?1:0))+'행: 날짜 또는 일정명을 확인하세요.');return;}if(end&&end<start){errors.push((i+1+(hasHeader?1:0))+'행: 종료일이 시작일보다 빠릅니다.');return;}
      var property=map.property>=0?String(r[map.property]||'').trim():'없음';if(!property)property='없음';if(properties.indexOf(property)<0)properties.push(property);
      importDraft.push({id:generateId(),start:start,end:end||start,title:title,property:property,target:'전체',memo:map.memo>=0?String(r[map.memo]||'').trim():'',createdAt:Date.now(),updatedAt:Date.now()});
    });saveProperties();
    var preview=importDraft.slice(0,8).map(function(e){return '<div><b>'+escapeHtml(e.start)+'</b><span>'+escapeHtml(e.title)+'</span><em>'+escapeHtml(e.property)+'</em></div>';}).join('');
    box.innerHTML='<p class="cal-import-ok">가져올 수 있는 일정 <b>'+importDraft.length+'개</b>'+(errors.length?' · 확인 필요 '+errors.length+'개':'')+'</p>'+preview+
      (importDraft.length>8?'<small>외 '+(importDraft.length-8)+'개</small>':'')+(errors.length?'<ul class="cal-import-errors">'+errors.slice(0,5).map(function(x){return'<li>'+escapeHtml(x)+'</li>';}).join('')+'</ul>':'');
    el('calImportBtn').disabled=importDraft.length===0;
  }
  function importEvents(){
    if(!importDraft.length)return;var known={};events.forEach(function(e){known[eventKey(e)]=true;});var added=0;
    importDraft.forEach(function(e){var key=eventKey(e);if(!known[key]){events.push(e);known[key]=true;added++;}});save();render();if(added&&typeof GoogleCalendarSync!=='undefined')GoogleCalendarSync.onLocalChanged();importDraft=[];el('calImportBtn').disabled=true;showToast(added+'개 일정을 가져왔습니다'+(added===0?' (모두 중복)':''));
  }
  function icsEscape(value){return String(value||'').replace(/\\/g,'\\\\').replace(/;/g,'\\;').replace(/,/g,'\\,').replace(/\r?\n/g,'\\n');}
  function compactDate(key){return String(key).replace(/-/g,'');}
  function exportIcs(){
    if(!events.length){showToast('내보낼 일정이 없습니다','error');return;}var lines=['BEGIN:VCALENDAR','VERSION:2.0','PRODID:-//NoteTong//School Calendar//KO','CALSCALE:GREGORIAN','METHOD:PUBLISH'];
    events.forEach(function(e){var prop=eventProperty(e);lines.push('BEGIN:VEVENT','UID:'+e.id+'@t-ong.vercel.app','DTSTAMP:'+compactDate(todayKey())+'T000000Z','DTSTART;VALUE=DATE:'+compactDate(e.start),'DTEND;VALUE=DATE:'+compactDate(addDays(e.end||e.start,1)),
      'SUMMARY:'+icsEscape((prop==='없음'?'':'['+prop+'] ')+e.title),'DESCRIPTION:'+icsEscape([prop==='없음'?'':prop,e.target,e.studentName,e.consultMethod,e.status,e.memo,e.followUp?'후속: '+e.followUp:''].filter(Boolean).join(' · ')),'END:VEVENT');});
    lines.push('END:VCALENDAR');var blob=new Blob([lines.join('\r\n')],{type:'text/calendar;charset=utf-8'}),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='T-ong-일정-'+todayKey()+'.ics';a.click();setTimeout(function(){URL.revokeObjectURL(a.href);},1000);showToast('구글 캘린더용 ICS 파일을 만들었습니다');
  }
  function consultationRows(){
    var classData=allClassData[currentClassIndex]||{};
    return events.filter(function(e){return eventProperty(e)==='상담';}).sort(function(a,b){return a.start.localeCompare(b.start);}).map(function(e){return{
      '학급':classData.className||'','날짜':e.start,'종료일':e.end||e.start,'상담대상':e.target||'학생','학생번호':e.studentId||'','학생명':String(e.studentName||'').replace(/^\d+번\s*/,''),
      '상담방식':e.consultMethod||'','진행상태':e.status||'','일정명':e.title||'','상담내용':e.memo||'','후속조치':e.followUp||''
    };});
  }
  function exportConsultExcel(){
    var rows=consultationRows();if(!rows.length){showToast('내보낼 상담 기록이 없습니다','error');return;}if(typeof XLSX==='undefined'){showToast('엑셀 기능을 불러오지 못했습니다','error');return;}
    var wb=XLSX.utils.book_new(),ws=XLSX.utils.json_to_sheet(rows);XLSX.utils.book_append_sheet(wb,ws,'상담기록');XLSX.writeFile(wb,'상담기록-'+todayKey()+'.xlsx');showToast('상담 기록 엑셀을 만들었습니다');
  }
  function copyConsultSheet(){
    var rows=consultationRows();if(!rows.length){showToast('복사할 상담 기록이 없습니다','error');return;}
    var headers=Object.keys(rows[0]),text=headers.join('\t')+'\n'+rows.map(function(row){return headers.map(function(h){return String(row[h]||'').replace(/[\t\r\n]+/g,' ');}).join('\t');}).join('\n');
    function fallback(){var ta=document.createElement('textarea');ta.value=text;ta.style.position='fixed';ta.style.opacity='0';document.body.appendChild(ta);ta.select();document.execCommand('copy');ta.remove();showToast('상담 기록을 복사했습니다');}
    if(navigator.clipboard&&navigator.clipboard.writeText)navigator.clipboard.writeText(text).then(function(){showToast('상담 기록을 복사했습니다');}).catch(fallback);else fallback();
  }
  function open(){load();if(!selectedDate)selectedDate=todayKey();var d=dateFromKey(selectedDate);monthCursor=new Date(d.getFullYear(),d.getMonth(),1);hideEditor();resetForm();render();openModal('calendarModal');}
  function init(){
    load();selectedDate=todayKey();monthCursor=new Date(new Date().getFullYear(),new Date().getMonth(),1);
    var form=el('calForm');if(form)form.addEventListener('submit',submit);
    var prev=el('calPrevBtn');if(prev)prev.addEventListener('click',function(){navigate(-1,true);});
    var next=el('calNextBtn');if(next)next.addEventListener('click',function(){navigate(1,true);});
    var today=el('calTodayBtn');if(today)today.addEventListener('click',function(){selectedDate=todayKey();monthCursor=new Date(new Date().getFullYear(),new Date().getMonth(),1);resetForm();render();});
    document.querySelectorAll('[data-cal-view]').forEach(function(button){button.addEventListener('click',function(){setView(this.dataset.calView);});});
    var todayAdd=el('calTodayAddBtn');if(todayAdd)todayAdd.addEventListener('click',openNewEditor);
    var todayPane=el('calTodayPane');if(todayPane)todayPane.addEventListener('dblclick',function(e){if(!e.target.closest('.cal-event-item,.btn'))openNewEditor();});
    initGestures();
    initEventDragging();
    var reset=el('calFormResetBtn');if(reset)reset.addEventListener('click',hideEditor);
    var deleteEvent=el('calDeleteEventBtn');if(deleteEvent)deleteEvent.addEventListener('click',function(){var id=el('calEditId').value;if(id&&removeEvent(id))hideEditor();});
    var property=el('calPropertyInput');if(property)property.addEventListener('change',function(){syncConsultFields();hidePropertyAdder();});
    var addPropertyButton=el('calAddPropertyBtn');if(addPropertyButton)addPropertyButton.addEventListener('click',function(){showPropertyAdder('add');});
    var editPropertyButton=el('calEditPropertyBtn');if(editPropertyButton)editPropertyButton.addEventListener('click',function(){showPropertyAdder('edit');});
    var deletePropertyButton=el('calDeletePropertyBtn');if(deletePropertyButton)deletePropertyButton.addEventListener('click',deleteProperty);
    var consultTarget=el('calConsultTargetInput');if(consultTarget)consultTarget.addEventListener('change',function(){var method=el('calConsultMethodInput');if(method&&this.value==='학부모'&&method.value==='학생 상담')method.value='학부모 상담';if(method&&this.value==='학생'&&method.value==='학부모 상담')method.value='학생 상담';syncConsultFields();});
    var propertySave=el('calPropertySaveBtn');if(propertySave)propertySave.addEventListener('click',confirmProperty);
    var propertyCancel=el('calPropertyCancelBtn');if(propertyCancel)propertyCancel.addEventListener('click',hidePropertyAdder);
    var propertyInput=el('calNewPropertyInput');if(propertyInput)propertyInput.addEventListener('keydown',function(e){if(e.key==='Enter'){e.preventDefault();confirmProperty();}else if(e.key==='Escape'){hidePropertyAdder();}});
    var add=el('calAddBtn');if(add)add.addEventListener('click',openNewEditor);
    var editorClose=el('calEditorCloseBtn');if(editorClose)editorClose.addEventListener('click',hideEditor);
    var editorOverlay=el('calEditorOverlay');if(editorOverlay)editorOverlay.addEventListener('click',function(e){if(e.target===this)hideEditor();});
    document.addEventListener('keydown',function(e){if(e.key==='Escape'&&editorOverlay&&!editorOverlay.hidden)hideEditor();});
    var help=el('calHelpBtn');if(help)help.addEventListener('click',function(){openGuideSection('calendar');});
    var googleButton=el('calGoogleBtn');if(googleButton)googleButton.addEventListener('click',function(){if(typeof GoogleCalendarSync!=='undefined')GoogleCalendarSync.openPanel();});
    var transfer=el('calTransferBtn');if(transfer)transfer.addEventListener('click',function(){var panel=el('calTransferPanel');panel.open=!panel.open;if(panel.open)setTimeout(function(){panel.scrollIntoView({behavior:'smooth',block:'start'});},20);});
    var sample=el('calSampleBtn');if(sample)sample.addEventListener('click',function(){el('calPasteInput').value='날짜\t종료일\t일정명\t속성\t메모\n2026-08-17\t\t교직원 회의\t회의\t2학기 준비\n2026-09-10\t2026-09-11\t교육청 출장\t출장\t연수 참석';previewPaste();});
    var preview=el('calPreviewBtn');if(preview)preview.addEventListener('click',previewPaste);var imp=el('calImportBtn');if(imp)imp.addEventListener('click',importEvents);
    var ics=el('calIcsBtn');if(ics)ics.addEventListener('click',exportIcs);var excel=el('calConsultExcelBtn');if(excel)excel.addEventListener('click',exportConsultExcel);
    var copy=el('calConsultCopyBtn');if(copy)copy.addEventListener('click',copyConsultSheet);
    var copyHelp=el('calConsultCopyHelpBtn');if(copyHelp)copyHelp.addEventListener('click',function(){openModal('sheetCopyHelpModal');});
  }
  function getData(){return events.slice();}function setData(value){events=Array.isArray(value)?value:[];save();render();}
  function getProperties(){return properties.slice();}
  function setProperties(value){var incoming=Array.isArray(value)?value:[];properties=unique((incoming.indexOf('없음')>=0?REQUIRED_PROPERTIES:DEFAULT_PROPERTIES).concat(incoming));saveProperties();renderPropertyOptions(properties[0]);}
  return{init:init,open:open,render:render,getData:getData,setData:setData,getProperties:getProperties,setProperties:setProperties};
})();

// ===================================================================
// Google Calendar — 사용자 기본(primary) 캘린더 양방향 동기화
// 접근 토큰은 메모리에만 보관하고, 계정별 매핑·syncToken만 localStorage에 저장합니다.
// ===================================================================
var GoogleCalendarSync = (function () {
  var CONFIG_KEY='tongGoogleCalendarConfig_v1',STATE_KEY='tongGoogleCalendarSync_v1';
  var SCOPES='openid email profile https://www.googleapis.com/auth/calendar.events.owned';
  var CALENDAR_ID='primary',SYNC_MODE='primary-v1';
  var API='https://www.googleapis.com/calendar/v3',USERINFO='https://openidconnect.googleapis.com/v1/userinfo';
  var token='',expiresAt=0,tokenClient=null,loadingPromise=null,syncing=false,timer=null,debounceTimer=null;
  var chooseAccount=false,expiryNotified=false,calendarOnlyAuthMode=false; // calendarOnlyAuthMode=true 이면 캘린더만 다른 계정으로 변경
  var config={clientId:'',activeAccountId:'',autoSync:true,includeConsult:false,includeMemo:false,customCalAccount:false};
  var accountStates={};
  function el(id){return document.getElementById(id);}
  function parse(value,fallback){try{var data=JSON.parse(value);return data&&typeof data==='object'?data:fallback;}catch(e){return fallback;}}
  function metaClientId(){var node=document.querySelector('meta[name="google-oauth-client-id"]');return node?String(node.content||'').trim():'';}
  function load(){
    var stored={};try{stored=parse(localStorage.getItem(CONFIG_KEY)||'{}',{});accountStates=parse(localStorage.getItem(STATE_KEY)||'{}',{});}catch(e){}
    config.clientId=String(metaClientId()||stored.clientId||'').trim();config.activeAccountId=String(stored.activeAccountId||'');
    config.autoSync=stored.autoSync!==false;config.includeConsult=stored.includeConsult===true;config.includeMemo=stored.includeMemo===true;
    config.customCalAccount=stored.customCalAccount===true;
  }
  function saveConfig(){try{localStorage.setItem(CONFIG_KEY,JSON.stringify(config));}catch(e){}}
  function saveStates(){try{localStorage.setItem(STATE_KEY,JSON.stringify(accountStates));}catch(e){}}
  function activeState(){return config.activeAccountId&&accountStates[config.activeAccountId]||null;}
  function ensureState(user){var id=String(user.sub||user.email||'');var state=accountStates[id]||{};state.accountId=id;state.email=user.email||state.email||'Google 계정';state.links=state.links||{};state.pendingDeletes=Array.isArray(state.pendingDeletes)?state.pendingDeletes:[];accountStates[id]=state;config.activeAccountId=id;saveConfig();saveStates();return state;}
  function validToken(){return !!token&&Date.now()<expiresAt-30000;}
  function getAccountSummary(){
    var st=activeState();
    return {
      connected: validToken(),
      email: st && st.email ? st.email : '',
      customCalAccount: !!config.customCalAccount
    };
  }
  function setBusy(value){syncing=value;var button=el('googleCalSyncBtn');if(button){button.disabled=!validToken()||value;button.textContent=value?'동기화 중…':'지금 양방향 동기화';}var full=el('googleCalFullSyncBtn');if(full)full.disabled=!validToken()||value;var connect=el('googleCalConnectBtn');if(connect)connect.disabled=value;var sw=el('googleCalSwitchOnlyBtn');if(sw)sw.disabled=value;}
  function setStatus(text,kind,detail){var status=el('googleCalStatus'),badge=el('googleCalStateBadge'),summary=el('googleCalSummaryStatus');if(status)status.textContent=text;if(badge){badge.textContent=kind==='ok'?'연결됨':kind==='busy'?'동기화 중':kind==='error'?'확인 필요':'대기';badge.dataset.state=kind||'idle';}if(summary)summary.textContent=kind==='ok'?'연결됨':kind==='busy'?'동기화 중':kind==='error'?'확인 필요':'연결 안 됨';if(detail&&el('googleCalAccount'))el('googleCalAccount').textContent=detail;}
  function render(){
    var state=activeState(),connected=validToken();if(el('googleClientIdInput'))el('googleClientIdInput').value=config.clientId||'';
    if(el('googleCalAutoInput'))el('googleCalAutoInput').checked=config.autoSync;if(el('googleCalConsultInput'))el('googleCalConsultInput').checked=config.includeConsult;if(el('googleCalMemoInput'))el('googleCalMemoInput').checked=config.includeMemo;
    var mainAuth=(typeof CloudSheetSync!=='undefined'&&CloudSheetSync.getAuth)?CloudSheetSync.getAuth():null;
    var isDiffFromMain=!!(config.customCalAccount||(mainAuth&&mainAuth.email&&state&&state.email&&mainAuth.email!==state.email));
    if(el('googleCalConnectBtn'))el('googleCalConnectBtn').textContent=connected?'통합 계정 다시 연결':(quickHint()?'통합 로그인 · '+quickHint():'Google 통합 로그인 · 연결');
    var useMainBtn=el('googleCalUseMainBtn');
    if(useMainBtn)useMainBtn.hidden=!isDiffFromMain||!(mainAuth&&mainAuth.email);
    var dev=el('googleCalDeveloperSettings');if(dev)dev.hidden=!!metaClientId()&&!/[?&]admin\b/.test(location.search);if(el('googleCalSyncBtn'))el('googleCalSyncBtn').disabled=!connected||syncing;if(el('googleCalFullSyncBtn'))el('googleCalFullSyncBtn').disabled=!connected||syncing;if(el('googleCalDisconnectBtn'))el('googleCalDisconnectBtn').disabled=!connected;
    var modeLabel=config.customCalAccount?'캘린더 별도 계정':'교무수첩 기본 계정 통합';
    if(connected&&state)setStatus('Google 캘린더와 연결되었습니다.','ok',(state.email||'Google 계정')+' ('+modeLabel+') · '+(state.calendarName||'Google 기본 캘린더'));
    else if(!config.clientId)setStatus('앱 관리자 설정이 필요합니다.','error','OAuth 클라이언트 ID를 입력한 뒤 계정을 연결해 주세요.');
    else setStatus('Google 계정을 연결해 주세요.','idle',state&&state.email?'대상 계정 · '+state.email+' ('+modeLabel+')':'Google 로그인 한 번으로 교무수첩과 캘린더가 함께 동기화됩니다.');
    var last=el('googleCalLastSync');if(last)last.textContent='마지막 동기화: '+(state&&state.lastSyncAt?new Date(state.lastSyncAt).toLocaleString('ko-KR'):'없음');
    if(typeof CloudSheetSync!=='undefined'&&CloudSheetSync.updateUI)CloudSheetSync.updateUI();
  }
  function loadIdentity(){
    if(window.google&&google.accounts&&google.accounts.oauth2)return Promise.resolve();if(loadingPromise)return loadingPromise;
    loadingPromise=new Promise(function(resolve,reject){var old=document.querySelector('script[data-tong-google-identity]');if(old){old.addEventListener('load',resolve,{once:true});old.addEventListener('error',function(){reject(new Error('Google 로그인 모듈을 불러오지 못했습니다.'));},{once:true});return;}var script=document.createElement('script');script.src='https://accounts.google.com/gsi/client';script.async=true;script.defer=true;script.dataset.tongGoogleIdentity='1';script.onload=resolve;script.onerror=function(){reject(new Error('Google 로그인 모듈을 불러오지 못했습니다.'));};document.head.appendChild(script);});return loadingPromise;
  }
  function initTokenClient(){tokenClient=google.accounts.oauth2.initTokenClient({client_id:config.clientId,scope:SCOPES,callback:function(response){if(response&&response.error){calendarOnlyAuthMode=false;setStatus('Google 연결을 완료하지 못했습니다.','error',response.error_description||response.error);setBusy(false);return;}token=response.access_token||'';expiresAt=Date.now()+(Number(response.expires_in)||3600)*1000;chooseAccount=false;expiryNotified=false;finishConnection().catch(handleError);}});}
  // 예전에 연결한 계정이 있으면 계정 선택·동의 화면 없이 한 번에 다시 연결합니다.
  function quickHint(){
    var state=activeState();
    if (!chooseAccount&&state&&state.email&&state.email.indexOf('@')>0) return state.email;
    if (!config.customCalAccount && typeof CloudSheetSync !== 'undefined') {
      var ca = CloudSheetSync.getAuth();
      if (ca && ca.email && ca.email.indexOf('@') > 0) return ca.email;
    }
    return '';
  }
  function connect(silent, forceSelectAccount){
    if(!config.clientId){
      var detail=el('googleCalDeveloperSettings');
      if(detail)detail.open=true;
      render();
      if(!silent) showToast('Google OAuth 클라이언트 ID를 먼저 설정해 주세요','error');
      return;
    }
    var hint=(forceSelectAccount||validToken())?'':quickHint();
    if(silent && !hint) return;
    setBusy(true);
    setStatus('Google 로그인 창을 준비하고 있습니다.','busy','연결할 계정을 선택해 주세요.');
    loadIdentity().then(function(){
      initTokenClient();
      if (forceSelectAccount) {
        tokenClient.requestAccessToken({prompt:'select_account'});
      } else {
        tokenClient.requestAccessToken(hint?{prompt:'',login_hint:hint}:{prompt:'select_account'});
      }
    }).catch(function(err){
      calendarOnlyAuthMode=false;
      if(!silent) handleError(err);
      else setBusy(false);
    });
  }
  // 교무수첩 + Google 캘린더 통합 1클릭 로그인
  function loginUnified(){
    calendarOnlyAuthMode = false;
    config.customCalAccount = false;
    saveConfig();
    connect(false, false);
  }
  // 교무수첩 기본 계정은 유지하고 캘린더만 다른 Google 계정으로 변경
  function switchCalendarAccountOnly(){
    calendarOnlyAuthMode = true;
    chooseAccount = true;
    config.customCalAccount = true;
    saveConfig();
    showToast('캘린더에 연결할 별도 Google 계정을 선택해 주세요 (교무수첩 기본 계정은 유지됩니다).', 'info');
    connect(false, true);
  }
  // 캘린더 계정을 다시 교무수첩 기본 계정으로 통일
  function useMainAccountForCalendar(){
    if (typeof CloudSheetSync === 'undefined' || !CloudSheetSync.getAuth) return;
    var mainAuth = CloudSheetSync.getAuth();
    if (!mainAuth || !mainAuth.email) {
      showToast('먼저 교무수첩 기본 Google 계정으로 로그인해 주세요.', 'error');
      return;
    }
    calendarOnlyAuthMode = false;
    chooseAccount = false;
    config.customCalAccount = false;
    var id = String(mainAuth.sub || mainAuth.email);
    config.activeAccountId = id;
    if (!accountStates[id]) {
      accountStates[id] = { accountId: id, email: mainAuth.email, links: {}, pendingDeletes: [] };
    } else {
      accountStates[id].email = mainAuth.email;
    }
    saveConfig();
    saveStates();
    token = '';
    expiresAt = 0;
    render();
    showToast('캘린더 계정을 교무수첩 기본 계정(' + mainAuth.email + ')으로 통일합니다.', 'info');
    connect(false, false);
  }
  function syncWithAccount(email, sub, passiveOnly){
    if (!email) return;
    // 사용자가 캘린더 전용 별도 계정을 설정해 둔 경우 기본 계정 로그인으로 덮어쓰지 않음
    if (config.customCalAccount && activeState() && activeState().email && activeState().email !== email) {
      render();
      return;
    }
    var id = String(sub || email);
    config.activeAccountId = id;
    if (!accountStates[id]) {
      accountStates[id] = { accountId: id, email: email, links: {}, pendingDeletes: [] };
    } else {
      accountStates[id].email = email;
    }
    saveConfig();
    saveStates();
    render();
    if (validToken()) {
      syncNow(true);
    } else if (!passiveOnly) {
      connect(true, false);
    }
  }
  async function api(pathname,options){
    options=options||{};var headers=Object.assign({'Authorization':'Bearer '+token},options.headers||{});if(options.body)headers['Content-Type']='application/json';var response=await fetch(pathname.indexOf('http')===0?pathname:API+pathname,Object.assign({},options,{headers:headers}));
    if(response.status===204)return null;var body=null;try{body=await response.json();}catch(e){}if(!response.ok){var err=new Error(body&&body.error&&body.error.message||('Google API 오류 '+response.status));err.status=response.status;throw err;}return body;
  }
  async function finishConnection(){
    var user=await api(USERINFO);
    var isCalOnly = calendarOnlyAuthMode;
    calendarOnlyAuthMode = false;
    // 통합 로그인 모드이거나 아직 교무수첩 기본 계정이 없는 경우 교무수첩 로그인도 한 번에 완료
    if (!isCalOnly && !config.customCalAccount && typeof CloudSheetSync !== 'undefined' && CloudSheetSync.setUnifiedAuthFromUser) {
      CloudSheetSync.setUnifiedAuthFromUser(user, false);
      showToast((user.name || user.email) + ' 계정으로 교무수첩 및 Google 캘린더가 동시 연결되었습니다.', 'success');
    } else if (isCalOnly) {
      showToast('Google 캘린더만 [' + (user.email || '별도 계정') + '] 계정으로 변경되었습니다.', 'success');
    }
    var state=ensureState(user);
    await ensureCalendar(state);
    setBusy(false);
    render();
    await syncNow(true);
    startTimer();
  }
  async function ensureCalendar(state){
    var migrated=state.calendarId!==CALENDAR_ID||state.syncMode!==SYNC_MODE;
    state.calendarId=CALENDAR_ID;state.calendarName='Google 기본 캘린더';state.syncMode=SYNC_MODE;
    if(migrated){state.syncToken=null;state.links={};state.pendingDeletes=[];state.lastSyncAt=0;}
    saveStates();
  }
  function addDays(key,amount){var p=String(key||'').split('-').map(Number),d=new Date(p[0],(p[1]||1)-1,p[2]||1);d.setDate(d.getDate()+amount);return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');}
  function propertyOf(item){return String(item.property||item.type||'없음').trim()||'없음';}
  function shouldSync(item){return config.includeConsult||propertyOf(item)!=='상담';}
  function toGoogle(item){var prop=propertyOf(item),description=config.includeMemo&&item.memo?String(item.memo).slice(0,160):'';return{summary:item.title,start:{date:item.start},end:{date:addDays(item.end||item.start,1)},description:description,extendedProperties:{private:{tongLocalId:String(item.id),tongUpdatedAt:String(Number(item.updatedAt)||Date.now()),tongProperty:prop,tongTarget:String(item.target||'전체')}}};}
  function fromGoogle(g,existing){var allDay=!!(g.start&&g.start.date),start=allDay?g.start.date:String(g.start&&g.start.dateTime||'').slice(0,10),endRaw=allDay?g.end&&g.end.date:String(g.end&&g.end.dateTime||'').slice(0,10),end=allDay&&endRaw?addDays(endRaw,-1):(endRaw||start);if(!start)return null;if(end<start)end=start;var priv=g.extendedProperties&&g.extendedProperties.private||{},stamp=Date.parse(g.updated)||Date.now(),item=Object.assign({},existing||{});item.id=existing&&existing.id||priv.tongLocalId||('gcal-'+g.id);item.title=String(g.summary||'(제목 없는 일정)').slice(0,80);item.start=start;item.end=end;item.property=priv.tongProperty||item.property||'없음';item.target=priv.tongTarget||item.target||'전체';if(config.includeMemo)item.memo=String(g.description||'').slice(0,160);else if(!existing)item.memo='';item.studentId=item.studentId||'';item.studentName=item.studentName||'';item.consultMethod=item.consultMethod||'';item.status=item.status||'';item.followUp=item.followUp||'';item.createdAt=item.createdAt||stamp;item.updatedAt=stamp;return item;}
  async function listChanges(state,forceFull){var items=[],pageToken='',nextSyncToken='',requestSyncToken=forceFull?'':state.syncToken;do{var query=['showDeleted=true','singleEvents=true','maxResults=2500'];if(requestSyncToken)query.push('syncToken='+encodeURIComponent(requestSyncToken));if(pageToken)query.push('pageToken='+encodeURIComponent(pageToken));var data;try{data=await api('/calendars/'+encodeURIComponent(state.calendarId)+'/events?'+query.join('&'));}catch(e){if(e.status===410&&requestSyncToken){state.syncToken=null;saveStates();return listChanges(state,true);}throw e;}items=items.concat(data.items||[]);pageToken=data.nextPageToken||'';nextSyncToken=data.nextSyncToken||nextSyncToken;}while(pageToken);return{items:items,nextSyncToken:nextSyncToken};}
  async function createRemote(state,item){var result=await api('/calendars/'+encodeURIComponent(state.calendarId)+'/events',{method:'POST',body:JSON.stringify(toGoogle(item))});state.links[item.id]={googleId:result.id,lastLocalUpdated:Number(item.updatedAt)||0,lastGoogleUpdated:result.updated||''};return result;}
  async function updateRemote(state,item,link){var result=await api('/calendars/'+encodeURIComponent(state.calendarId)+'/events/'+encodeURIComponent(link.googleId),{method:'PATCH',body:JSON.stringify(toGoogle(item))});link.lastLocalUpdated=Number(item.updatedAt)||0;link.lastGoogleUpdated=result.updated||link.lastGoogleUpdated||'';return result;}
  async function deleteRemote(state,link){try{await api('/calendars/'+encodeURIComponent(state.calendarId)+'/events/'+encodeURIComponent(link.googleId),{method:'DELETE'});}catch(e){if(e.status!==404&&e.status!==410)throw e;}}
  async function syncNow(silent,options){
    var forceFull=!!(options&&options.forceFull);
    if(syncing)return;if(!validToken()){render();if(!silent)showToast('Google 계정을 다시 연결해 주세요','error');else promptReconnect();return;}var state=activeState();if(!state)return;
    if(state.calendarId!==CALENDAR_ID||state.syncMode!==SYNC_MODE)await ensureCalendar(state);
    setBusy(true);setStatus(forceFull?'Google 기본 캘린더 전체를 확인하고 있습니다.':'양방향 일정을 확인하고 있습니다.','busy',(state.email||'Google 계정')+' · '+(state.calendarName||'Google 기본 캘린더'));
    try{
      var pulledCount=0,pushedCount=0;
      var local=ScheduleCalendar.getData().map(function(item){if(!item.updatedAt)item.updatedAt=0;return item;}),byId={};local.forEach(function(item){byId[item.id]=item;});
      for(var p=0;p<state.pendingDeletes.length;p++){var pending=state.pendingDeletes[p];if(pending.googleId){await deleteRemote(state,pending);pushedCount++;}delete state.links[pending.localId];}state.pendingDeletes=[];
      var changes=await listChanges(state,forceFull),localChanged=false;
      for(var i=0;i<changes.items.length;i++){
        var g=changes.items[i],priv=g.extendedProperties&&g.extendedProperties.private||{},localId=priv.tongLocalId||'',link=localId&&state.links[localId];
        if(!link){Object.keys(state.links).some(function(id){if(state.links[id].googleId===g.id){localId=id;link=state.links[id];return true;}return false;});}
        if(!localId&&byId['gcal-'+g.id])localId='gcal-'+g.id;
        var item=localId&&byId[localId];
        if(g.status==='cancelled'){
          if(item&&link){var localWasChanged=Number(item.updatedAt||0)>Number(link.lastLocalUpdated||0);if(localWasChanged&&Number(item.updatedAt||0)>Date.parse(g.updated||0)){delete state.links[localId];await createRemote(state,item);pushedCount++;}else{local=local.filter(function(x){return x.id!==localId;});delete byId[localId];delete state.links[localId];localChanged=true;pulledCount++;}}continue;
        }
        if(!item){var imported=fromGoogle(g,null);if(imported){local.push(imported);byId[imported.id]=imported;state.links[imported.id]={googleId:g.id,lastLocalUpdated:Number(imported.updatedAt)||0,lastGoogleUpdated:g.updated||''};localChanged=true;pulledCount++;}continue;}
        if(!link){link=state.links[item.id]={googleId:g.id,lastLocalUpdated:0,lastGoogleUpdated:''};}
        var localWasChanged=Number(item.updatedAt||0)>Number(link.lastLocalUpdated||0),googleWasChanged=(g.updated||'')!==(link.lastGoogleUpdated||'');
        if(googleWasChanged&&(!localWasChanged||Date.parse(g.updated||0)>Number(item.updatedAt||0))){var pulled=fromGoogle(g,item);var pos=local.findIndex(function(x){return x.id===item.id;});local[pos]=pulled;byId[pulled.id]=pulled;link.lastLocalUpdated=Number(pulled.updatedAt)||0;link.lastGoogleUpdated=g.updated||'';localChanged=true;pulledCount++;}
        else if(localWasChanged){await updateRemote(state,item,link);pushedCount++;}else{link.lastGoogleUpdated=g.updated||link.lastGoogleUpdated||'';}
      }
      for(var j=0;j<local.length;j++){
        var current=local[j],currentLink=state.links[current.id];
        if(!shouldSync(current)){if(currentLink){await deleteRemote(state,currentLink);delete state.links[current.id];pushedCount++;}continue;}
        if(!current.updatedAt){current.updatedAt=Date.now();localChanged=true;}
        if(!currentLink){await createRemote(state,current);pushedCount++;}else if(Number(current.updatedAt||0)>Number(currentLink.lastLocalUpdated||0)){await updateRemote(state,current,currentLink);pushedCount++;}
      }
      if(localChanged)ScheduleCalendar.setData(local);state.syncToken=changes.nextSyncToken||state.syncToken||null;state.lastSyncAt=Date.now();saveStates();setBusy(false);render();if(!silent)showToast((forceFull?'전체 다시 가져오기 완료':'양방향 동기화 완료')+' · Google에서 '+pulledCount+'개 · Google로 '+pushedCount+'개');
    }catch(e){setBusy(false);handleError(e);}
  }
  function fullSync(){
    if(syncing)return;if(!validToken()){render();showToast('Google 계정을 다시 연결해 주세요','error');return;}if(!activeState())return;
    if(!confirm('Google 기본 캘린더의 모든 일정을 다시 확인할까요?\n\nT-ong의 기존 일정을 일괄 삭제하지는 않습니다. 다만 양방향 동기화이므로 T-ong에서 수정·삭제한 내용도 Google에 반영되며, 일정이 많으면 시간이 걸릴 수 있습니다.'))return;
    syncNow(false,{forceFull:true});
  }
  function handleError(error){setBusy(false);if(error&&error.status===401){token='';expiresAt=0;stopTimer();setStatus('Google 계정을 다시 연결해 주세요.','error','접근 권한 시간이 만료되었습니다.');render();if(quickHint()){expiryNotified=true;setBusy(false);showToast('Google 캘린더 연결 시간이 끝났습니다','error','다시 연결',connect);return;}}else if(error&&error.status===403){setStatus('Google Calendar API 설정을 확인해 주세요.','error','API 사용 설정과 OAuth 권한을 확인해 주세요.');}else setStatus('동기화하지 못했습니다.','error',error&&error.message||'네트워크 상태를 확인해 주세요.');renderButtonsOnly();showToast(error&&error.message||'Google 캘린더 동기화 오류','error');}
  function renderButtonsOnly(){if(el('googleCalSyncBtn'))el('googleCalSyncBtn').disabled=!validToken()||syncing;if(el('googleCalFullSyncBtn'))el('googleCalFullSyncBtn').disabled=!validToken()||syncing;if(el('googleCalDisconnectBtn'))el('googleCalDisconnectBtn').disabled=!validToken();}
  function startTimer(){stopTimer();if(config.autoSync&&validToken())timer=setInterval(function(){if(document.visibilityState==='visible')syncNow(true);},60000);}
  function stopTimer(){if(timer){clearInterval(timer);timer=null;}}
  // 토큰(약 1시간)이 끝나면 탭 한 번으로 다시 연결하도록 안내합니다. 팝업은 사용자 탭에서만 열 수 있습니다.
  function promptReconnect(){if(expiryNotified||!config.clientId||!quickHint())return;expiryNotified=true;stopTimer();showToast('Google 캘린더 연결 시간이 끝났습니다','error','다시 연결',connect);}
  function onLocalChanged(){clearTimeout(debounceTimer);if(config.autoSync&&token&&!validToken()){promptReconnect();return;}if(!config.autoSync||!validToken())return;debounceTimer=setTimeout(function(){syncNow(true);},1400);}
  function noteLocalDelete(item){var state=activeState(),link=state&&state.links&&state.links[item.id];if(!state||!link)return;state.pendingDeletes=state.pendingDeletes.filter(function(x){return x.localId!==item.id;});state.pendingDeletes.push({localId:item.id,googleId:link.googleId,deletedAt:Date.now()});saveStates();}
  function disconnect(){var old=token;token='';expiresAt=0;stopTimer();chooseAccount=true;if(old&&window.google&&google.accounts&&google.accounts.oauth2&&google.accounts.oauth2.revoke)google.accounts.oauth2.revoke(old,function(){});render();showToast('Google 계정 연결을 해제했습니다');}
  function saveClientId(){var input=el('googleClientIdInput'),value=String(input&&input.value||'').trim();if(value&&!/^[0-9A-Za-z._-]+\.apps\.googleusercontent\.com$/.test(value)){showToast('OAuth 클라이언트 ID 형식을 확인해 주세요','error');return;}config.clientId=value;tokenClient=null;saveConfig();render();showToast(value?'Google OAuth 클라이언트 ID를 저장했습니다':'클라이언트 ID를 비웠습니다');}
  function optionChanged(){config.autoSync=!!el('googleCalAutoInput').checked;config.includeConsult=!!el('googleCalConsultInput').checked;config.includeMemo=!!el('googleCalMemoInput').checked;saveConfig();startTimer();if(validToken())onLocalChanged();}
  function openPanel(){var panel=el('calGooglePanel');if(!panel)return;panel.open=!panel.open;render();if(panel.open)setTimeout(function(){panel.scrollIntoView({behavior:'smooth',block:'start'});},20);}
  function init(){load();render();var connectButton=el('googleCalConnectBtn'),switchOnlyButton=el('googleCalSwitchOnlyBtn'),useMainButton=el('googleCalUseMainBtn'),syncButton=el('googleCalSyncBtn'),fullSyncButton=el('googleCalFullSyncBtn'),disconnectButton=el('googleCalDisconnectBtn'),saveButton=el('googleClientIdSaveBtn');if(connectButton)connectButton.addEventListener('click',loginUnified);if(switchOnlyButton)switchOnlyButton.addEventListener('click',switchCalendarAccountOnly);if(useMainButton)useMainButton.addEventListener('click',useMainAccountForCalendar);if(syncButton)syncButton.addEventListener('click',function(){syncNow(false);});if(fullSyncButton)fullSyncButton.addEventListener('click',fullSync);if(disconnectButton)disconnectButton.addEventListener('click',disconnect);if(saveButton)saveButton.addEventListener('click',saveClientId);var helpButton=el('googleOAuthHelpBtn');if(helpButton)helpButton.addEventListener('click',function(e){e.preventDefault();e.stopPropagation();openModal('googleOAuthHelpModal');});['googleCalAutoInput','googleCalConsultInput','googleCalMemoInput'].forEach(function(id){var node=el(id);if(node)node.addEventListener('change',optionChanged);});document.addEventListener('visibilitychange',function(){if(document.visibilityState==='visible'&&validToken()&&config.autoSync)syncNow(true);});}
  return{init:init,openPanel:openPanel,onLocalChanged:onLocalChanged,noteLocalDelete:noteLocalDelete,syncNow:syncNow,syncWithAccount:syncWithAccount,validToken:validToken,connect:connect,loginUnified:loginUnified,switchCalendarAccountOnly:switchCalendarAccountOnly,useMainAccountForCalendar:useMainAccountForCalendar,getAccountSummary:getAccountSummary};
})();
// ===================================================================
// AI 비서 — 선생님 본인의 Claude API 키로 사이드바에서 대화하고 T-ong 기능을 실행합니다.
// 키와 대화는 이 브라우저에서 Anthropic으로 바로 가며, T-ong 서버로는 보내지 않습니다.
// ===================================================================
var AiAssistant = (function () {
  var PREF_KEY = 'tongAiAssistant_v2', KEY_KEY = 'tongAiApiKey_v1';
  var SDK_SRC = '/vendor/anthropic-sdk-0.132.1.js';
  var MODELS = ['claude-haiku-5-5', 'claude-sonnet-5-5', 'claude-opus-5-5'];
  var FALLBACK_MODELS = ['claude-opus-5-5', 'claude-sonnet-5-5'];   // 서버 측 대체 모델(fallbacks: "default")을 쓰는 모델
  var WEEKDAYS = ['일', '월', '화', '수', '목', '금', '토'];
  var pref = { model: 'claude-haiku-5-5', mask: true, remember: true };
  var apiKey = '', client = null, clientKey = '', sdkPromise = null;
  var history = [];            // API에 보내는 대화 — 응답 content를 그대로 덧붙이는 방식(append-only)
  var nameMap = {}, maskCount = 0;   // '[학생A]' → 실제 이름, 대화 하나 안에서는 같은 가명 유지
  var busy = false, currentStream = null;

  var TOOLS = [
    { name: 'get_class_info', description: '현재 학급 이름, 학생 수, 오늘 날짜, 지금 저장된 알림장 내용을 확인합니다.',
      input_schema: { type: 'object', properties: {} } },
    { name: 'list_events', description: 'T-ong 일정관리에서 기간 안의 일정을 조회합니다. 날짜는 YYYY-MM-DD.',
      input_schema: { type: 'object', properties: { from: { type: 'string', description: '시작일 YYYY-MM-DD' }, to: { type: 'string', description: '종료일 YYYY-MM-DD' } }, required: ['from', 'to'] } },
    { name: 'add_event', description: 'T-ong 일정관리에 일정을 추가합니다. 실행 전에 교사가 화면에서 확인 버튼을 누릅니다. 상담 일정은 추가할 수 없습니다.',
      input_schema: { type: 'object', properties: {
        title: { type: 'string', description: '일정명 (80자 이내)' },
        start: { type: 'string', description: '시작일 YYYY-MM-DD' },
        end: { type: 'string', description: '종료일 YYYY-MM-DD, 하루 일정이면 생략' },
        property: { type: 'string', description: '일정 속성. get_class_info의 eventProperties 중 하나, 모르면 생략' },
        memo: { type: 'string', description: '짧은 메모 (선택)' } }, required: ['title', 'start'] } },
    { name: 'set_notice', description: '오늘의 알림장 내용을 저장합니다. 실행 전에 교사가 화면에서 확인 버튼을 누릅니다.',
      input_schema: { type: 'object', properties: {
        text: { type: 'string', description: '알림장 본문. 줄마다 한 항목' },
        mode: { type: 'string', enum: ['replace', 'append'], description: 'replace는 기존 내용을 바꾸고 append는 뒤에 덧붙입니다' } }, required: ['text', 'mode'] } }
  ].map(function (tool) { tool.eager_input_streaming = true; return tool; });

  function el(id) { return document.getElementById(id); }
  function todayKey() { return formatDate(new Date()); }

  // ---- 설정·키 저장 -------------------------------------------------
  function loadPref() {
    try {
      var v = JSON.parse(localStorage.getItem(PREF_KEY) || '{}');
      if (MODELS.indexOf(v.model) >= 0) pref.model = v.model;
      if (v.mask === false) pref.mask = false;
      if (v.remember === false) pref.remember = false;
    } catch (e) {}
    try { apiKey = localStorage.getItem(KEY_KEY) || sessionStorage.getItem(KEY_KEY) || ''; } catch (e) { apiKey = ''; }
  }
  function savePref() { try { localStorage.setItem(PREF_KEY, JSON.stringify(pref)); } catch (e) {} }
  function storeKey(value) {
    apiKey = value;
    try {
      localStorage.removeItem(KEY_KEY); sessionStorage.removeItem(KEY_KEY);
      if (value) (pref.remember ? localStorage : sessionStorage).setItem(KEY_KEY, value);
    } catch (e) {}
  }
  function maskedKey(value) { return value ? value.slice(0, 7) + '…' + value.slice(-4) : ''; }

  function renderSettings() {
    el('aiModelInput').value = pref.model;
    el('aiMaskInput').checked = pref.mask;
    el('aiRememberInput').checked = pref.remember;
    el('aiKeyInput').value = '';
    el('aiKeyStatus').textContent = apiKey ? '저장된 키: ' + maskedKey(apiKey) + (pref.remember ? ' · 이 기기에 저장됨' : ' · 이 창을 닫으면 지워짐') : '아직 키가 없습니다.';
    el('aiSetupCard').hidden = !!apiKey || !el('aiSettings').hidden;
  }
  function toggleSettings(show) {
    var box = el('aiSettings');
    box.hidden = typeof show === 'boolean' ? !show : !box.hidden;
    el('aiSettingsBtn').setAttribute('aria-expanded', box.hidden ? 'false' : 'true');
    renderSettings();
    if (!box.hidden) el('aiKeyInput').focus();
  }
  function saveKey() {
    var value = String(el('aiKeyInput').value || '').trim();
    if (!value) { showToast('API 키를 붙여넣어 주세요', 'error'); return; }
    if (!/^sk-ant-[A-Za-z0-9_\-]{20,}$/.test(value)) { showToast('sk-ant-로 시작하는 Claude API 키인지 확인해 주세요', 'error'); return; }
    storeKey(value); client = null; toggleSettings(false);
    showToast('API 키를 저장했습니다. 이제 질문해 보세요');
    el('aiPromptInput').focus();
  }
  function deleteKey() {
    if (!apiKey) { showToast('저장된 키가 없습니다'); return; }
    if (!confirm('이 기기에 저장된 API 키를 지울까요?')) return;
    storeKey(''); client = null; renderSettings(); showToast('API 키를 지웠습니다');
  }
  function openKeyGuide() {
    openModal('aiGuideModal');
    var tab = document.querySelector('[data-ai-guide="apikey"]'); if (tab) tab.click();
  }

  // ---- 학생 이름 가리기 ----------------------------------------------
  function rosterNames() {
    return (Array.isArray(students) ? students : [])
      .map(function (s) { return String(s && s.name || '').trim(); })
      .filter(function (n, i, arr) { return n.length >= 2 && !/^학생\d+$/.test(n) && arr.indexOf(n) === i; })
      .sort(function (a, b) { return b.length - a.length; });   // 긴 이름부터 — '김민'이 '김민준'을 먼저 먹지 않게
  }
  function mask(text) {
    text = String(text == null ? '' : text);
    if (!pref.mask) return text;
    rosterNames().forEach(function (name) {
      if (text.indexOf(name) === -1) return;
      var label = Object.keys(nameMap).find(function (k) { return nameMap[k] === name; });
      if (!label) { label = '[학생' + (maskCount < 26 ? String.fromCharCode(65 + maskCount) : (maskCount + 1)) + ']'; maskCount++; nameMap[label] = name; }
      text = text.split(name).join(label);
    });
    return text;
  }
  function unmask(text) {
    text = String(text == null ? '' : text);
    Object.keys(nameMap).forEach(function (label) { text = text.split(label).join(nameMap[label]); });
    return text;
  }

  // ---- 화면 ------------------------------------------------------------
  function open() {
    loadPref(); renderSettings();
    el('aiDrawer').classList.add('show'); el('aiDrawer').setAttribute('aria-hidden', 'false'); el('aiDrawerBackdrop').classList.add('show');
    if (!history.length && !el('aiChat').children.length) addBubble('ai', '안녕하세요! 알림장 작성, 일정 조회·추가, 조회 멘트나 가정통신문 쓰기를 도와드릴게요. 무엇을 할까요?');
    setTimeout(function () { (apiKey ? el('aiPromptInput') : el('aiSetupOpenBtn')).focus(); }, 60);
  }
  function close() { el('aiDrawer').classList.remove('show'); el('aiDrawer').setAttribute('aria-hidden', 'true'); el('aiDrawerBackdrop').classList.remove('show'); }
  function scrollChat() { var c = el('aiChat'); c.scrollTop = c.scrollHeight; }
  function addBubble(kind, text) {
    var div = document.createElement('div');
    div.className = 'ai-msg ai-msg-' + kind;
    div.textContent = text || '';
    el('aiChat').appendChild(div); scrollChat();
    return div;
  }
  function setBusy(value) {
    busy = value;
    var send = el('aiSendBtn');
    send.textContent = value ? '중단' : '보내기';
    send.classList.toggle('stop', value);
    el('aiPromptInput').disabled = value;
  }
  function newChat() {
    if (busy) return;
    history = []; nameMap = {}; maskCount = 0;
    el('aiChat').innerHTML = '';
    addBubble('ai', '새 대화를 시작했습니다. 무엇을 도와드릴까요?');
  }

  // 알림장 저장·일정 추가는 선생님이 확인 버튼을 눌러야 실행됩니다.
  function confirmAction(title, previewText) {
    var card = document.createElement('div');
    card.className = 'ai-confirm';
    var head = document.createElement('b'); head.textContent = title;
    var body = document.createElement('pre'); body.textContent = previewText;
    var row = document.createElement('div'); row.className = 'ai-confirm-row';
    var ok = document.createElement('button'); ok.type = 'button'; ok.className = 'btn primary'; ok.textContent = '실행';
    var no = document.createElement('button'); no.type = 'button'; no.className = 'btn'; no.textContent = '취소';
    row.appendChild(no); row.appendChild(ok);
    card.appendChild(head); card.appendChild(body); card.appendChild(row);
    el('aiChat').appendChild(card); scrollChat();
    return new Promise(function (resolve) {
      function done(answer) { ok.disabled = no.disabled = true; card.classList.add(answer ? 'done' : 'cancelled'); row.textContent = answer ? '✅ 실행했습니다' : '취소했습니다'; resolve(answer); }
      ok.addEventListener('click', function () { done(true); });
      no.addEventListener('click', function () { done(false); });
    });
  }

  // ---- 도구 ------------------------------------------------------------
  function isDateKey(v) { return typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v) && !isNaN(Date.parse(v)); }
  function eventProperties() { return ScheduleCalendar.getProperties().filter(function (p) { return p !== '상담'; }); }
  function validate(name, input) {
    input = input && typeof input === 'object' ? input : null;
    if (!input) return '입력이 비어 있습니다';
    if (name === 'list_events') return isDateKey(input.from) && isDateKey(input.to) ? '' : 'from과 to는 YYYY-MM-DD 형식이어야 합니다';
    if (name === 'add_event') {
      if (typeof input.title !== 'string' || !input.title.trim()) return 'title이 필요합니다';
      if (!isDateKey(input.start)) return 'start는 YYYY-MM-DD 형식이어야 합니다';
      if (input.end != null && !isDateKey(input.end)) return 'end는 YYYY-MM-DD 형식이어야 합니다';
      if (input.end && input.end < input.start) return 'end가 start보다 빠릅니다';
      return '';
    }
    if (name === 'set_notice') return typeof input.text === 'string' && input.text.trim() && (input.mode === 'replace' || input.mode === 'append') ? '' : 'text와 mode(replace 또는 append)가 필요합니다';
    return '';
  }
  async function runTool(name, input) {
    var cls = allClassData[currentClassIndex] || {};
    if (name === 'get_class_info') {
      return JSON.stringify({ className: cls.className || ((currentClassIndex + 1) + '반'), studentCount: (Array.isArray(students) ? students : []).length,
        today: todayKey(), weekday: WEEKDAYS[new Date().getDay()], notice: mask(cls.notice || ''), eventProperties: eventProperties() });
    }
    if (name === 'list_events') {
      var list = ScheduleCalendar.getData().filter(function (e) { return (e.end || e.start) >= input.from && e.start <= input.to; })
        .sort(function (a, b) { return a.start < b.start ? -1 : a.start > b.start ? 1 : 0; })
        .slice(0, 60)
        .map(function (e) { return { title: mask(e.title), start: e.start, end: e.end || e.start, property: e.property || '없음', memo: mask(e.memo || '') }; });
      return JSON.stringify({ count: list.length, events: list });
    }
    if (name === 'add_event') {
      var props = eventProperties(), property = props.indexOf(input.property) >= 0 ? input.property : '없음';
      var item = { id: generateId(), title: unmask(input.title).trim().slice(0, 80), start: input.start, end: input.end || input.start, property: property, target: '전체',
        memo: unmask(input.memo || '').trim().slice(0, 160), studentId: '', studentName: '', consultMethod: '', status: '', followUp: '', createdAt: Date.now(), updatedAt: Date.now() };
      var range = item.start + (item.end !== item.start ? ' ~ ' + item.end : '');
      var ok = await confirmAction('📅 일정을 추가할까요?', item.title + '\n' + range + (property !== '없음' ? '\n속성: ' + property : '') + (item.memo ? '\n메모: ' + item.memo : ''));
      if (!ok) return '교사가 일정 추가를 취소했습니다.';
      ScheduleCalendar.setData(ScheduleCalendar.getData().concat([item]));
      if (typeof GoogleCalendarSync !== 'undefined') GoogleCalendarSync.onLocalChanged();
      showToast('일정을 추가했습니다');
      return '일정을 추가했습니다: ' + mask(item.title) + ' (' + range + ')';
    }
    if (name === 'set_notice') {
      var text = unmask(input.text).trim(), before = (cls.notice || '').trim();
      var next = input.mode === 'append' && before ? before + '\n' + text : text;
      var ok2 = await confirmAction(input.mode === 'append' && before ? '📢 알림장에 덧붙일까요?' : '📢 알림장을 이렇게 저장할까요?', next);
      if (!ok2) return '교사가 알림장 저장을 취소했습니다.';
      if (!allClassData[currentClassIndex]) return '현재 학급 정보를 찾지 못했습니다.';
      allClassData[currentClassIndex].notice = next;
      allClassData[currentClassIndex].noticeAt = formatDateTime(new Date());
      saveToLocalStorage(); Notice.renderPreview();
      showToast('알림장을 저장했습니다', '', '열기', function () { var b = el('noticeBtn'); if (b) b.click(); });
      return '알림장을 저장했습니다.';
    }
    throw new Error('알 수 없는 도구: ' + name);
  }
  var TOOL_LABELS = { get_class_info: '🏫 학급 정보 확인', list_events: '📅 일정 확인', add_event: '📅 일정 추가 준비', set_notice: '📢 알림장 준비' };

  // ---- Claude API -------------------------------------------------------
  function loadSdk() {
    if (window.TongAnthropic) return Promise.resolve(window.TongAnthropic);
    if (sdkPromise) return sdkPromise;
    sdkPromise = new Promise(function (resolve, reject) {
      var script = document.createElement('script');
      script.src = SDK_SRC; script.async = true;
      script.onload = function () { window.TongAnthropic ? resolve(window.TongAnthropic) : reject(new Error('AI 모듈을 불러오지 못했습니다')); };
      script.onerror = function () { sdkPromise = null; reject(new Error('AI 모듈을 불러오지 못했습니다. 인터넷 연결을 확인해 주세요')); };
      document.head.appendChild(script);
    });
    return sdkPromise;
  }
  async function getClient() {
    var Anthropic = await loadSdk();
    if (!client || clientKey !== apiKey) {
      // 각 선생님이 자기 키를 자기 브라우저에서만 쓰는 구조라 브라우저 직접 호출을 허용합니다.
      client = new Anthropic({ apiKey: apiKey, dangerouslyAllowBrowser: true, maxRetries: 2 });
      clientKey = apiKey;
    }
    return { client: client, Anthropic: Anthropic };
  }
  function systemPrompt() {
    var d = new Date();
    return '당신은 한국 학교 교사가 쓰는 학급 관리 앱 "T-ong 교무수첩"의 AI 비서입니다. 오늘은 ' + todayKey() + ' (' + WEEKDAYS[d.getDay()] + '요일)입니다.\n' +
      '- 한국어로, 교사가 바로 쓸 수 있게 간결하고 실용적으로 답합니다.\n' +
      '- 학생 이름은 개인정보 보호를 위해 [학생A] 같은 가명으로 바뀌어 있을 수 있습니다. 가명은 그대로 쓰고 실제 이름을 추측하지 않습니다.\n' +
      '- 알림장 저장이나 일정 추가를 부탁받으면 도구를 사용합니다. 도구는 실행 전에 교사의 확인을 받으므로 따로 허락을 구하지 않아도 됩니다. 취소되면 다시 시도하지 말고 어떻게 고칠지 묻습니다.\n' +
      '- "다음 주 수요일"처럼 상대적인 날짜는 오늘 날짜를 기준으로 YYYY-MM-DD로 계산합니다.\n' +
      '- 준비물, 시간, 장소처럼 교사가 알려 주지 않은 사실은 지어내지 말고 물어보거나 빈칸으로 남깁니다.\n' +
      '- 알림장은 번호나 기호를 붙인 짧은 줄 3~6개로, 학생과 학부모가 읽기 쉽게 씁니다.';
  }
  function friendlyError(err, Anthropic) {
    if (Anthropic && err instanceof Anthropic.APIUserAbortError) return '중단했습니다.';
    if (Anthropic && err instanceof Anthropic.AuthenticationError) return 'API 키가 올바르지 않습니다. ⚙️ 설정에서 키를 다시 확인해 주세요.';
    if (Anthropic && err instanceof Anthropic.PermissionDeniedError) return '이 키로는 선택한 모델을 쓸 수 없습니다. 다른 모델을 고르거나 Console의 권한을 확인해 주세요.';
    if (Anthropic && err instanceof Anthropic.RateLimitError) return '요청이 많아 잠시 제한되었습니다. 조금 뒤에 다시 시도해 주세요.';
    if (Anthropic && err instanceof Anthropic.BadRequestError) return '요청을 처리하지 못했습니다. 크레딧 잔액이 부족한지 Console의 Billing에서 확인해 주세요.\n(' + (err.message || '') + ')';
    if (Anthropic && err instanceof Anthropic.APIConnectionError) return '인터넷 연결을 확인해 주세요.';
    if (Anthropic && err instanceof Anthropic.APIError) return 'AI 서버 오류가 발생했습니다 (' + (err.status || '') + '). 잠시 뒤 다시 시도해 주세요.';
    return (err && err.message) || '알 수 없는 오류가 발생했습니다.';
  }

  async function send(promptText) {
    if (busy) { if (currentStream) currentStream.abort(); return; }
    var text = String(promptText != null ? promptText : el('aiPromptInput').value || '').trim();
    if (!text) { el('aiPromptInput').focus(); return; }
    if (!apiKey) { toggleSettings(true); showToast('먼저 Claude API 키를 넣어 주세요', 'error'); return; }
    el('aiPromptInput').value = '';
    addBubble('me', text);
    history.push({ role: 'user', content: mask(text) });
    setBusy(true);
    var Anthropic = null, jsonRetries = 0;
    try {
      var api = await getClient(); Anthropic = api.Anthropic;
      for (var turn = 0; turn < 8; turn++) {
        var bubble = addBubble('ai pending', '생각 중…'), acc = '';
        var params = { model: pref.model, max_tokens: 64000, system: systemPrompt(), tools: TOOLS, messages: history,
          output_config: { effort: 'low' }, cache_control: { type: 'ephemeral' } };
        // 안전 분류기가 요청을 거절하면 서버가 같은 요청을 다른 모델로 이어서 처리합니다.
        if (FALLBACK_MODELS.indexOf(pref.model) >= 0) { params.betas = ['server-side-fallback-2026-07-01']; params.fallbacks = 'default'; }
        currentStream = api.client.beta.messages.stream(params);
        currentStream.on('text', function (delta) {
          acc += delta;
          bubble.classList.remove('pending');
          bubble.textContent = unmask(acc);
          scrollChat();
        });
        var message;
        try {
          message = await currentStream.finalMessage();
          jsonRetries = 0;
        } catch (err) {
          bubble.remove();
          // 도구 입력이 JSON으로 읽히지 않은 경우만 같은 턴을 다시 요청합니다. API 오류는 그대로 처리합니다.
          if (err instanceof Anthropic.APIError || err instanceof Anthropic.APIUserAbortError || jsonRetries++ >= 2) throw err;
          turn--; continue;
        }
        if (!acc) bubble.remove();
        if (message.stop_reason === 'refusal') { addBubble('ai error', 'AI가 이 요청에는 답하지 않았습니다. 표현을 바꿔 다시 부탁해 주세요.'); break; }
        var toolUses = message.content.filter(function (b) { return b.type === 'tool_use'; });
        if (message.stop_reason === 'max_tokens' && toolUses.length) { addBubble('ai error', '답변이 너무 길어 중간에 끊겼습니다. 요청을 나눠서 다시 부탁해 주세요.'); break; }
        history.push({ role: 'assistant', content: message.content });
        if (message.stop_reason === 'pause_turn') continue;
        if (!toolUses.length) break;
        var results = [];
        for (var i = 0; i < toolUses.length; i++) {
          var use = toolUses[i], problem = validate(use.name, use.input);
          addBubble('tool', TOOL_LABELS[use.name] || use.name);
          if (problem) { results.push({ type: 'tool_result', tool_use_id: use.id, is_error: true, content: problem }); continue; }
          try { results.push({ type: 'tool_result', tool_use_id: use.id, content: await runTool(use.name, use.input) }); }
          catch (toolErr) { results.push({ type: 'tool_result', tool_use_id: use.id, is_error: true, content: String(toolErr && toolErr.message || toolErr) }); }
        }
        history.push({ role: 'user', content: results });   // 도구 결과는 한 메시지에 모아서 보냅니다
      }
    } catch (err) {
      addBubble('ai error', friendlyError(err, Anthropic));
      if (Anthropic && err instanceof Anthropic.AuthenticationError) toggleSettings(true);
    } finally {
      currentStream = null; setBusy(false); el('aiPromptInput').focus();
    }
  }

  function init() {
    if (!el('aiDrawer')) return;
    loadPref(); renderSettings();
    Array.prototype.forEach.call(document.querySelectorAll('[data-ai-open]'), function (b) { b.addEventListener('click', open); });
    el('aiDrawerClose').addEventListener('click', close);
    el('aiDrawerBackdrop').addEventListener('click', close);
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && el('aiDrawer').classList.contains('show')) close(); });
    el('aiSettingsBtn').addEventListener('click', function () { toggleSettings(); });
    el('aiNewChatBtn').addEventListener('click', newChat);
    el('aiKeySaveBtn').addEventListener('click', saveKey);
    el('aiKeyInput').addEventListener('keydown', function (e) { if (e.key === 'Enter') saveKey(); });
    el('aiKeyDeleteBtn').addEventListener('click', deleteKey);
    el('aiKeyGuideBtn').addEventListener('click', openKeyGuide);
    el('aiSetupGuideBtn').addEventListener('click', openKeyGuide);
    el('aiSetupOpenBtn').addEventListener('click', function () { toggleSettings(true); });
    el('aiModelInput').addEventListener('change', function () { pref.model = this.value; savePref(); });
    el('aiMaskInput').addEventListener('change', function () { pref.mask = this.checked; savePref(); });
    el('aiRememberInput').addEventListener('change', function () { pref.remember = this.checked; savePref(); if (apiKey) storeKey(apiKey); renderSettings(); });
    el('aiSendBtn').addEventListener('click', function () { send(); });
    el('aiPromptInput').addEventListener('keydown', function (e) { if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) { e.preventDefault(); send(); } });
    el('aiQuick').addEventListener('click', function (e) {
      var b = e.target.closest('[data-ai-prompt]'); if (!b || busy) return;
      var input = el('aiPromptInput'); input.value = b.dataset.aiPrompt; input.focus();
      input.setSelectionRange(input.value.length, input.value.length);
    });
  }

  return { init: init, open: open };
})();
// ===================================================================
// 조회·종례 타이머 — 시작하면 전체화면 모래시계로
// ===================================================================
var Timer = (function () {
  var total = 180, left = 180, running = false, tick = null, warned = false;
  var RING = 2 * Math.PI * 88;

  function el(id) { return document.getElementById(id); }

  function fmt(s) {
    var m = Math.floor(Math.max(0, s) / 60);
    var r = Math.max(0, s) % 60;
    return String(m).padStart(2, '0') + ':' + String(r).padStart(2, '0');
  }

  function setPresetState(minutes, custom) {
    document.querySelectorAll('#tsPresets .ts-preset').forEach(function (c) {
      c.classList.toggle('active', custom ? !!c.dataset.tsCustom : parseFloat(c.dataset.tsMin) === minutes);
    });
  }

  function setManualVisible(show) {
    var manual = el('tsManual');
    if (manual) manual.hidden = !show;
  }

  function fillManualInputs() {
    var mi = el('tsMinInput'), si = el('tsSecInput');
    if (mi) mi.value = String(Math.floor(total / 60));
    if (si) si.value = total % 60 ? String(total % 60) : '';
    if (mi) { mi.focus(); mi.select(); }
  }

  function updateLabel() {
    var cls = allClassData[currentClassIndex];
    var mm = Math.floor(total / 60), ss = total % 60;
    var duration = (mm ? mm + '분' : '') + (mm && ss ? ' ' : '') + (ss ? ss + '초' : '');
    var label = el('tsLabel');
    if (label) label.textContent = ((cls && cls.className) ? cls.className + ' · ' : '') + duration + ' 타이머';
  }

  function paint() {
    var text = fmt(left);
    var urgent = left <= 10 && left > 0 && running;
    updateLabel();

    var time = el('tsTime');
    if (time) {
      time.textContent = text;
      time.classList.toggle('warning', urgent);
      time.classList.toggle('done', left === 0);
    }
    var ratio = total > 0 ? Math.max(0, left) / total : 0;
    var ring = el('ringFill');
    if (ring) {
      ring.style.strokeDasharray = RING;
      ring.style.strokeDashoffset = RING * (1 - ratio);
      ring.classList.toggle('warning', urgent);
    }
    var top = el('sandTop'), bot = el('sandBot'), stream = el('sandStream');
    if (top && bot) {
      var topH = 76 * ratio;
      top.setAttribute('y', String(18 + (76 - topH)));
      top.setAttribute('height', String(topH));
      var botH = 74 * (1 - ratio);
      bot.setAttribute('y', String(172 - botH));
      bot.setAttribute('height', String(botH));
    }
    if (stream) stream.style.opacity = (running && left > 0) ? '1' : '0';

    var toggleBtn = el('tsToggleBtn');
    if (toggleBtn) {
      toggleBtn.textContent = running ? '⏸ 멈춤' : '▶ 시작';
      toggleBtn.classList.toggle('play', !running);
      toggleBtn.classList.toggle('pause', running);
    }
    var hg = el('hourglass');
    if (hg) hg.classList.toggle('flowing', running && left > 0);
  }

  function openStage() {
    var stage = el('timerStage');
    if (!stage) return;
    var custom = !!(el('tsCustomBtn') && el('tsCustomBtn').classList.contains('active'));
    setManualVisible(custom);
    stage.classList.add('show');
    paint();
  }

  function closeStage() {
    var stage = el('timerStage');
    if (stage) stage.classList.remove('show');
  }

  function setMinutes(m) {
    stop();
    total = Math.round(m * 60);
    left = total;
    warned = false;
    setPresetState(m, false);
    setManualVisible(false);
    paint();
  }

  function start() {
    if (running || left <= 0) return;
    Sound.init();
    running = true;
    openStage();
    paint();
    tick = setInterval(function () {
      left--;
      if (left <= 10 && left > 0 && !warned) warned = true;
      if (left <= 10 && left > 0) Sound.playTick();
      if (left <= 0) {
        left = 0;
        stop();
        Sound.playAlarm();
        Confetti.celebrate();
        showToast('⏰ 시간이 다 되었습니다');
      }
      paint();
    }, 1000);
  }

  function stop() {
    running = false;
    if (tick) { clearInterval(tick); tick = null; }
    paint();
  }

  function toggle() { running ? stop() : start(); }
  function reset() { stop(); left = total; warned = false; paint(); }

  function setManual() {
    var mi = el('tsMinInput'), si = el('tsSecInput');
    var mm = Math.max(0, Math.min(180, parseInt(mi && mi.value, 10) || 0));
    var ss = Math.max(0, Math.min(59, parseInt(si && si.value, 10) || 0));
    var secs = mm * 60 + ss;
    if (secs <= 0) { showToast('시간을 입력해 주세요', 'error'); return; }
    stop();
    total = secs;
    left = secs;
    warned = false;
    setPresetState(null, true);
    paint();
    showToast('타이머를 ' + (mm ? mm + '분 ' : '') + (ss ? ss + '초' : '') + '로 맞췄습니다');
  }

  function init() {
    document.querySelectorAll('#tsPresets .ts-preset').forEach(function (c) {
      c.addEventListener('click', function () {
        if (this.dataset.tsCustom) {
          setPresetState(null, true);
          setManualVisible(true);
          fillManualInputs();
          return;
        }
        setMinutes(parseFloat(this.dataset.tsMin));
      });
    });
    var setBtn = el('tsSetBtn'); if (setBtn) setBtn.addEventListener('click', setManual);
    ['tsMinInput', 'tsSecInput'].forEach(function (id) {
      var input = el(id);
      if (input) input.addEventListener('keydown', function (e) { if (e.key === 'Enter') setManual(); });
    });
    var toggleBtn = el('tsToggleBtn'); if (toggleBtn) toggleBtn.addEventListener('click', toggle);
    var resetBtn = el('tsResetBtn'); if (resetBtn) resetBtn.addEventListener('click', reset);
    var closeBtn = el('tsCloseBtn'); if (closeBtn) closeBtn.addEventListener('click', closeStage);
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeStage(); });
    paint();
  }

  return { init: init, setMinutes: setMinutes, openStage: openStage };
})();
// ===================================================================
// 룰렛 뽑기 — 발표자 · 도우미 · 심부름
// ===================================================================
var Picker = (function () {
  var picked = [];
  var rolling = false;

  function pool() {
    var noRepeat = document.getElementById('pickNoRepeat');
    if (noRepeat && noRepeat.checked) {
      var rest = students.filter(function (s) { return picked.indexOf(s.id) === -1; });
      return rest.length ? rest : null;     // 전원 소진
    }
    return students.slice();
  }

  function render() {
    var box = document.getElementById('pickPicked');
    if (!box) return;
    box.innerHTML = '';
    picked.forEach(function (id) {
      var s = students.find(function (x) { return x.id === id; });
      if (!s) return;
      var chip = document.createElement('span');
      chip.className = 'picked-chip';
      chip.textContent = s.name;
      box.appendChild(chip);
    });
  }

  function howMany() {
    var i = document.getElementById('pickHowMany');
    return Math.max(1, Math.min(10, parseInt(i && i.value, 10) || 1));
  }

  function go() {
    if (rolling) return;
    if (!students.length) { showToast('명렬표를 먼저 등록해 주세요', 'error'); return; }
    var list = pool();
    if (!list) { showToast('모두 뽑았습니다. 기록을 지워주세요'); return; }

    var n = Math.min(howMany(), list.length);
    rolling = true;
    Sound.init();
    var noEl = document.querySelector('#pickResult .pick-no');
    var nameEl = document.querySelector('#pickResult .pick-name');
    var result = document.getElementById('pickResult');
    result.classList.remove('settled');

    // 뽑을 사람들을 미리 정하고, 화면에는 돌아가는 연출만 보여줍니다
    var shuffled = list.slice();
    shuffleArray(shuffled);
    var winners = shuffled.slice(0, n);

    var spins = 16, i = 0;
    var timer = setInterval(function () {
      var s = list[Math.floor(Math.random() * list.length)];
      noEl.textContent = s.id + '번';
      nameEl.textContent = s.name;
      Sound.playSpin(i, spins);          // 돌아가는 동안 째깍째깍
      if (++i >= spins) {
        clearInterval(timer);
        if (n === 1) {
          noEl.textContent = winners[0].id + '번';
          nameEl.textContent = winners[0].name;
        } else {
          noEl.textContent = n + '명 당첨';
          nameEl.textContent = winners.map(function (w) { return w.name; }).join(', ');
        }
        result.classList.add('settled');
        winners.forEach(function (w) { picked.push(w.id); });
        render();
        Sound.playWin();
        Confetti.burstAt(result, 40 + n * 8);
        rolling = false;
      }
    }, 70);
  }

  function reset() { picked = []; render(); showToast('기록을 지웠습니다'); }

  function open() {
    var noEl = document.querySelector('#pickResult .pick-no');
    var nameEl = document.querySelector('#pickResult .pick-name');
    if (noEl) noEl.textContent = '';
    if (nameEl) nameEl.textContent = '준비!';
    document.getElementById('pickResult').classList.remove('settled');
    render();
    openModal('pickModal');
  }

  function init() {
    var b = document.getElementById('pickStudentBtn'); if (b) b.addEventListener('click', open);
    var g = document.getElementById('pickGoBtn'); if (g) g.addEventListener('click', go);
    var r = document.getElementById('pickResetBtn'); if (r) r.addEventListener('click', reset);
    var pm = document.getElementById('pickMinus');
    var pp = document.getElementById('pickPlus');
    var ph = document.getElementById('pickHowMany');
    if (pm && ph) pm.addEventListener('click', function () { ph.value = Math.max(1, (parseInt(ph.value,10)||1) - 1); });
    if (pp && ph) pp.addEventListener('click', function () { ph.value = Math.min(10, (parseInt(ph.value,10)||1) + 1); });
  }

  return { init: init };
})();

// ===================================================================
// 오늘의 알림장 — 학급별로 저장, 좌석표 옆에 상시 표시
// ===================================================================
var Notice = (function () {
  function data() { return allClassData[currentClassIndex] || {}; }

  function renderPreview() {
    var box = document.getElementById('noticePreview');
    if (!box) return;
    var text = (data().notice || '').trim();
    box.classList.toggle('empty', !text);
    box.textContent = text || '아직 작성한 알림장이 없습니다';
  }

  function open() {
    var input = document.getElementById('noticeInput');
    if (input) input.value = data().notice || '';
    var at = document.getElementById('noticeSavedAt');
    if (at) at.textContent = data().noticeAt ? ('마지막 저장 ' + data().noticeAt) : '';
    openModal('noticeModal');
  }

  function save() {
    var input = document.getElementById('noticeInput');
    if (!allClassData[currentClassIndex]) return;
    allClassData[currentClassIndex].notice = input ? input.value : '';
    allClassData[currentClassIndex].noticeAt = formatDateTime(new Date());
    saveToLocalStorage();
    renderPreview();
    closeModal('noticeModal');
    showToast('알림장을 저장했습니다');
  }

  function clear() {
    var input = document.getElementById('noticeInput');
    if (input) input.value = '';
  }

  // 교실 TV에 띄우는 큰 화면
  function showBoard() {
    var text = (document.getElementById('noticeInput').value || data().notice || '').trim();
    if (!text) { showToast('먼저 내용을 적어주세요', 'error'); return; }
    var cls = data().className || ((currentClassIndex + 1) + '반');
    document.getElementById('nbClass').textContent = cls;
    document.getElementById('nbDate').textContent = formatDate(new Date());
    var body = document.getElementById('nbBody');
    body.innerHTML = '';
    text.split('\n').forEach(function (line) {
      if (!line.trim()) return;
      var p = document.createElement('p');
      p.textContent = line;
      body.appendChild(p);
    });
    document.getElementById('noticeBoard').classList.add('show');
  }

  function hideBoard() { document.getElementById('noticeBoard').classList.remove('show'); }

  function init() {
    var b = document.getElementById('noticeBtn'); if (b) b.addEventListener('click', open);
    var card = document.getElementById('noticePreviewCard'); if (card) card.addEventListener('click', open);
    var s = document.getElementById('noticeSaveBtn'); if (s) s.addEventListener('click', save);
    var c = document.getElementById('noticeClearBtn'); if (c) c.addEventListener('click', clear);
    var big = document.getElementById('noticeBigBtn'); if (big) big.addEventListener('click', showBoard);
    var x = document.getElementById('nbCloseBtn'); if (x) x.addEventListener('click', hideBoard);
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') hideBoard();
    });
    renderPreview();
  }

  return { init: init, renderPreview: renderPreview };
})();

// ===================================================================
// 칭찬 포인트 — 학생별 점수, 좌석 카드에 표시
// ===================================================================
var Praise = (function () {
  function points() {
    var d = allClassData[currentClassIndex];
    if (!d) return {};
    if (!d.praise) d.praise = {};
    return d.praise;
  }

  function stage(score) {
    score = Math.max(0, Number(score) || 0);
    if (score >= 10) return { key: 'chicken', icon: '🐔', label: '닭', score: score };
    if (score >= 5) return { key: 'chick', icon: '🐥', label: '병아리', score: score };
    if (score >= 1) return { key: 'hatching', icon: '🐣', label: '부화 중', score: score };
    return { key: 'egg', icon: '🥚', label: '알', score: 0 };
  }

  function add(studentId, delta) {
    var p = points();
    p[studentId] = Math.max(0, get(studentId) + delta);
    saveCurrentClassData();
    render();
    renderSeating();
    if (delta > 0) Sound.playPick(0, 5, false);
  }

  function get(studentId) { return Math.max(0, Number(points()[studentId]) || 0); }

  function render() {
    var list = document.getElementById('praiseList');
    if (!list) return;
    list.innerHTML = '';
    if (!students.length) {
      list.innerHTML = '<p class="hint">명렬표를 먼저 등록해 주세요</p>';
      return;
    }
    var sorted = students.slice().sort(function (a, b) {
      var diff = get(b.id) - get(a.id);
      return diff !== 0 ? diff : (parseInt(a.id) - parseInt(b.id));
    });
    sorted.forEach(function (s) {
      var n = get(s.id);
      var growth = stage(n);
      var row = document.createElement('div');
      row.className = 'praise-row stage-' + growth.key + (n > 0 ? ' has' : '');
      var plus = document.createElement('button');
      plus.type = 'button';
      plus.className = 'praise-name';
      plus.innerHTML = '<span class="pr-no">' + s.id + '</span><span class="pr-name">' + escapeHtml(s.name) + '</span>' +
        '<span class="pr-score"><span class="pr-stage-icon">' + growth.icon + '</span><span class="pr-stage-text">' + growth.label + '<small>' + n + '점</small></span></span>';
      plus.title = s.name + ' 학생 칭찬 +1 · 현재 ' + growth.label + ' ' + n + '점';
      plus.addEventListener('click', function () { add(s.id, 1); });
      var minus = document.createElement('button');
      minus.type = 'button';
      minus.className = 'praise-minus';
      minus.textContent = '−';
      minus.disabled = n <= 0;
      minus.setAttribute('aria-label', s.name + ' 학생 칭찬 1점 차감');
      minus.addEventListener('click', function () { add(s.id, -1); });
      row.appendChild(plus);
      row.appendChild(minus);
      list.appendChild(row);
    });
  }

  function reset() {
    if (!confirm('칭찬 포인트를 모두 0으로 되돌릴까요?')) return;
    if (allClassData[currentClassIndex]) allClassData[currentClassIndex].praise = {};
    saveCurrentClassData();
    render();
    renderSeating();
    showToast('칭찬 포인트를 초기화했습니다');
  }

  function open() { render(); openModal('praiseModal'); }

  function init() {
    var b = document.getElementById('praiseBtn'); if (b) b.addEventListener('click', open);
    var r = document.getElementById('praiseResetBtn'); if (r) r.addEventListener('click', reset);
  }

  return { init: init, get: get, add: add, stage: stage, render: render };
})();

// ===================================================================
// 출결 체크 — 날짜별로 저장, 좌석표에 표시
// ===================================================================
var Attend = (function () {
  var STATES = ['출석', '결석', '지각', '조퇴', '결과'];
  var ATT_STATES = ['', '결석', '지각', '조퇴', '결과'];
  var ATT_REASONS = ['', '질병', '인정', '미인정', '기타'];
  var MARK = { '결석': '결', '지각': '지', '조퇴': '조', '결과': '과' };

  // 보고 있는 날짜 (null 이면 오늘). 캘린더로 지난 날짜도 소급 입력합니다.
  var viewDate = null;
  function todayKey() { return viewDate || formatDate(new Date()); }

  function setDate(key) {
    var today = formatDate(new Date());
    viewDate = (key && key !== today) ? key : null;
    render();
    renderSeating();
  }

  function shiftDate(days) {
    var d = new Date(todayKey() + 'T00:00:00');
    d.setDate(d.getDate() + days);
    setDate(formatDate(d));
  }

  function bag(dateKey) {
    var d = allClassData[currentClassIndex];
    if (!d) return {};
    if (!d.attendance) d.attendance = {};
    var k = dateKey || todayKey();
    if (!d.attendance[k]) d.attendance[k] = {};
    return d.attendance[k];
  }

  function get(studentId, dateKey) { return bag(dateKey)[studentId] || '출석'; }

  function memoBag(dateKey) {
    var d = allClassData[currentClassIndex];
    if (!d) return {};
    if (!d.attendMemos) d.attendMemos = {};
    var k = dateKey || todayKey();
    if (!d.attendMemos[k]) d.attendMemos[k] = {};
    return d.attendMemos[k];
  }
  function getMemo(studentId, dateKey) { return memoBag(dateKey)[studentId] || ''; }

  function setState(studentId, v, dateKey) {
    var b = bag(dateKey), r = reasonBag(dateKey), m = memoBag(dateKey);
    if (!v) {
      b[studentId] = undefined; r[studentId] = undefined; m[studentId] = undefined;
      clean(b, studentId); clean(r, studentId); clean(m, studentId);
    } else {
      b[studentId] = v;
    }
    saveCurrentClassData();
    render();
    renderSeating();
  }

  function setReason(studentId, v, dateKey) {
    var r = reasonBag(dateKey);
    if (v) r[studentId] = v; else { r[studentId] = undefined; clean(r, studentId); }
    saveCurrentClassData();
  }

  function setMemo(studentId, v, dateKey) {
    v = (v || '').trim();
    var m = memoBag(dateKey);
    if (v) m[studentId] = v; else { m[studentId] = undefined; clean(m, studentId); }
    saveCurrentClassData();
  }

  function clean(obj, key) { if (obj[key] === undefined) { try { Reflect.deleteProperty(obj, key); } catch (e) { obj[key] = ''; } } }

  function mark(studentId) { return MARK[get(studentId, formatDate(new Date()))] || ''; }

  function render() {
    var dInput = document.getElementById('avDate');
    if (dInput) dInput.value = todayKey();
    var dSpan = document.getElementById('attendDate');
    if (dSpan) dSpan.textContent = todayKey();
    var tBtn = document.getElementById('avToday');
    if (tBtn) tBtn.classList.toggle('is-today', !viewDate);
    renderInto(document.getElementById('attendList'), document.getElementById('attendSummary'));
    renderInto(document.getElementById('avList'), document.getElementById('avSummary'));
  }

  function renderInto(list, sum) {
    if (!list) return;
    list.innerHTML = '';
    if (!students.length) {
      list.innerHTML = '<p class="hint">명렬표를 먼저 등록해 주세요</p>';
      if (sum) sum.innerHTML = '';
      return;
    }
    var counts = { '결석': 0, '지각': 0, '조퇴': 0, '결과': 0 };

    students.forEach(function (s) {
      var st = get(s.id);
      var isAbsent = st !== '출석';
      if (isAbsent && counts[st] !== undefined) counts[st]++;

      var row = document.createElement('div');
      row.className = 'attend-item' + (isAbsent ? ' state-' + st : '');

      var who = document.createElement('button');
      who.type = 'button';
      who.className = 'at-who attend-student-link';
      who.title = s.name + ' 학생 종합 기록 열기';
      who.innerHTML = '<span class="at-no">' + s.id + '</span><span class="at-name">' + escapeHtml(s.name) + '</span>';
      who.addEventListener('click', function () { openStudentSettings(s.id); });
      row.appendChild(who);

      // 근태 종류 (출석이면 빈칸)
      var sel = document.createElement('select');
      sel.className = 'at-sel state';
      ATT_STATES.forEach(function (v) {
        var o = document.createElement('option');
        o.value = v;
        o.textContent = v === '' ? '—' : v;
        if ((v === '' && !isAbsent) || v === st) o.selected = true;
        sel.appendChild(o);
      });
      sel.addEventListener('change', function () { setState(s.id, this.value); });
      row.appendChild(sel);

      // 사유 (근태가 있을 때만)
      var rsel = document.createElement('select');
      rsel.className = 'at-sel reason';
      ATT_REASONS.forEach(function (v) {
        var o = document.createElement('option');
        o.value = v;
        o.textContent = v === '' ? '사유' : v;
        if (v === getReason(s.id)) o.selected = true;
        rsel.appendChild(o);
      });
      rsel.disabled = !isAbsent;
      rsel.addEventListener('change', function () { setReason(s.id, this.value); });
      row.appendChild(rsel);

      // 메모
      var memo = document.createElement('input');
      memo.type = 'text';
      memo.className = 'at-memo';
      memo.placeholder = '메모';
      memo.value = getMemo(s.id);
      memo.disabled = !isAbsent;
      memo.addEventListener('change', function () { setMemo(s.id, this.value); });
      row.appendChild(memo);

      list.appendChild(row);
    });

    if (sum) {
      var absent = counts['결석'] + counts['지각'] + counts['조퇴'] + counts['결과'];
      sum.innerHTML =
        '<span class="as-chip ok">출석 <b>' + (students.length - absent) + '</b></span>' +
        '<span class="as-chip ab">결석 <b>' + counts['결석'] + '</b></span>' +
        '<span class="as-chip la">지각 <b>' + counts['지각'] + '</b></span>' +
        '<span class="as-chip ea">조퇴 <b>' + counts['조퇴'] + '</b></span>' +
        '<span class="as-chip rs">결과 <b>' + counts['결과'] + '</b></span>';
    }
  }

  function resetAll() {
    var d = allClassData[currentClassIndex];
    var key = todayKey();
    if (d) {
      if (!d.attendance) d.attendance = {};
      d.attendance[key] = {};
      if (d.attendReasons) d.attendReasons[key] = {};
      if (d.attendMemos) d.attendMemos[key] = {};
    }
    saveCurrentClassData();
    render();
    renderSeating();
    showToast('전원 출석으로 되돌렸습니다');
  }

  // ── 사유 ──
  var reasonTarget = null;

  function reasonBag(dateKey) {
    var d = allClassData[currentClassIndex];
    if (!d) return {};
    if (!d.attendReasons) d.attendReasons = {};
    var k = dateKey || todayKey();
    if (!d.attendReasons[k]) d.attendReasons[k] = {};
    return d.attendReasons[k];
  }

  function getReason(studentId, dateKey) { return reasonBag(dateKey)[studentId] || ''; }

  // ── 월별 엑셀 ──
  // 기간 선택: 이번 달 / 지난 달 / 전체
  function periodMode() {
    var sel = document.getElementById('avPeriod');
    return sel ? sel.value : 'this';
  }

  function monthKeys(mode) {
    var d = allClassData[currentClassIndex];
    var att = (d && d.attendance) || {};
    var keys = Object.keys(att).sort();
    mode = mode || periodMode();
    if (mode === 'all') return keys;
    var now = new Date();
    if (mode === 'prev') now.setMonth(now.getMonth() - 1);
    var prefix = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
    return keys.filter(function (k) { return k.indexOf(prefix) === 0; });
  }

  function periodLabel(mode) {
    mode = mode || periodMode();
    if (mode === 'all') return '전체기간';
    var now = new Date();
    if (mode === 'prev') now.setMonth(now.getMonth() - 1);
    return now.getFullYear() + '년 ' + (now.getMonth() + 1) + '월';
  }

  // 출결만 담은 표 (자리배치와 분리 — 교과교사도 그대로 사용)
  function buildTable() {
    var d = allClassData[currentClassIndex];
    var keys = monthKeys();
    var head = ['학번', '이름', '성별'];
    keys.forEach(function (k) { head.push(k); });
    head.push('결석', '지각', '조퇴', '결과', '사유·메모');

    var rows = students.map(function (s) {
      var row = [s.id, s.name, s.gender || ''];
      var c = { '결석': 0, '지각': 0, '조퇴': 0, '결과': 0 };
      var notes = [];
      keys.forEach(function (k) {
        var st = ((d.attendance || {})[k] || {})[s.id] || '';
        var rz = ((d.attendReasons || {})[k] || {})[s.id];
        row.push(st && st !== '출석' ? st + (rz ? '(' + rz + ')' : '') : '');
        if (c[st] !== undefined) c[st]++;
        var mm = ((d.attendMemos || {})[k] || {})[s.id];
        if (st && st !== '출석' && (rz || mm)) {
          notes.push(k.slice(5) + ' ' + st + (rz ? '(' + rz + ')' : '') + (mm ? ' - ' + mm : ''));
        }
      });
      row.push(c['결석'], c['지각'], c['조퇴'], c['결과'], notes.join(' / '));
      return row;
    });
    return { head: head, rows: rows, keys: keys };
  }

  function copyToSheet() {
    var t = buildTable();
    if (!t.keys.length) { showToast('해당 기간 기록이 없습니다', 'error'); return; }
    var tsv = [t.head].concat(t.rows).map(function (r) {
      return r.map(function (v) { return String(v == null ? '' : v).replace(/[\t\n\r]/g, ' '); }).join('\t');
    }).join('\n');
    var finish = function () {
      showToast('복사했습니다! 구글시트에서 Ctrl+V 하세요');
      window.open('https://sheets.new', '_blank', 'noopener');
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(tsv).then(finish, function () { showToast('복사 실패', 'error'); });
    } else { showToast('이 브라우저에서는 복사를 지원하지 않습니다', 'error'); }
  }

  function exportMonth(silent) {
    var d = allClassData[currentClassIndex];
    var t = buildTable();
    if (!t.keys.length) { if (!silent) showToast('해당 기간 기록이 없습니다', 'error'); return false; }
    var title = (d.className || ((currentClassIndex + 1) + '반')) + ' ' + periodLabel() + ' 출결';
    var ws = XLSX.utils.aoa_to_sheet([t.head].concat(t.rows));
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '출결');
    XLSX.writeFile(wb, title + '.xlsx');
    return true;
  }

  // ── 다운로드 후 초기화 (학생 기록 보호) ──
  function purge() {
    if (!confirm('이번 달 출결 기록을 엑셀로 내려받은 뒤,\n저장된 출결 기록·사유·메모를 모두 지웁니다.\n\n계속할까요?')) return;
    var ok = exportMonth(true);
    if (!ok && !confirm('내려받을 기록이 없습니다. 그래도 초기화할까요?')) return;
    var d = allClassData[currentClassIndex];
    if (d) { d.attendance = {}; d.attendReasons = {}; d.attendMemos = {}; }
    saveCurrentClassData();
    render();
    renderSeating();
    showToast('출결 기록을 초기화했습니다');
  }

  function open() { render(); openModal('attendModal'); }

  function init() {
    var b = document.getElementById('attendBtn'); if (b) b.addEventListener('click', open);
    var r = document.getElementById('attendResetBtn'); if (r) r.addEventListener('click', resetAll);
    var m = document.getElementById('attendMonthBtn'); if (m) m.addEventListener('click', function () { exportMonth(); });
    var p = document.getElementById('attendPurgeBtn'); if (p) p.addEventListener('click', purge);
    var vm = document.getElementById('avMonthBtn'); if (vm) vm.addEventListener('click', function () { exportMonth(); });
    var vp = document.getElementById('avPurgeBtn'); if (vp) vp.addEventListener('click', purge);
    var vr = document.getElementById('avResetBtn'); if (vr) vr.addEventListener('click', resetAll);
    var vs = document.getElementById('avSheetBtn'); if (vs) vs.addEventListener('click', copyToSheet);
    var dp = document.getElementById('avDate');
    if (dp) dp.addEventListener('change', function () { setDate(this.value); });
    var pd = document.getElementById('avPrevDay'); if (pd) pd.addEventListener('click', function () { shiftDate(-1); });
    var nd = document.getElementById('avNextDay'); if (nd) nd.addEventListener('click', function () { shiftDate(1); });
    var td = document.getElementById('avToday'); if (td) td.addEventListener('click', function () { setDate(null); });
    var vpd = document.getElementById('avPeriod'); if (vpd) vpd.addEventListener('change', render);
  }

  return { init: init, mark: mark, get: get, render: render,
           getReason: getReason, getMemo: getMemo,
           setState: setState, setReason: setReason, setMemo: setMemo };
})();

// ===================================================================
// 1인 1역 배정 — 역할을 무작위로 나눠 줍니다
// ===================================================================
var Duty = (function () {
  function data() {
    var d = allClassData[currentClassIndex];
    if (!d) return {};
    if (!d.duties) d.duties = {};
    return d.duties;
  }

  function roleText() {
    var d = allClassData[currentClassIndex];
    return (d && d.dutyRoles) || '';
  }

  function assign() {
    var raw = document.getElementById('dutyRoles').value;
    var roles = raw.split('\n').map(function (r) { return r.trim(); }).filter(Boolean);
    if (!roles.length) { showToast('역할을 한 줄에 하나씩 적어주세요', 'error'); return; }
    if (!students.length) { showToast('명렬표를 먼저 등록해 주세요', 'error'); return; }

    var pool = students.slice();
    shuffleArray(pool);
    var d = allClassData[currentClassIndex];
    d.duties = {};
    d.dutyRoles = raw;
    pool.forEach(function (s, i) { d.duties[s.id] = roles[i % roles.length]; });
    saveCurrentClassData();
    render();
    Sound.init();
    Sound.playFanfare();
    showToast('1인 1역을 배정했습니다');
  }

  function clear() {
    var d = allClassData[currentClassIndex];
    if (d) d.duties = {};
    saveCurrentClassData();
    render();
  }

  function render() {
    var box = document.getElementById('dutyResult');
    if (!box) return;
    var duties = data();
    var ids = Object.keys(duties);
    if (!ids.length) { box.innerHTML = '<p class="hint">아직 배정하지 않았습니다.</p>'; return; }

    // 역할별로 묶어서 보여줍니다
    var byRole = {};
    students.forEach(function (s) {
      var r = duties[s.id];
      if (!r) return;
      if (!byRole[r]) byRole[r] = [];
      byRole[r].push(s);
    });
    box.innerHTML = '';
    Object.keys(byRole).forEach(function (role) {
      var card = document.createElement('div');
      card.className = 'duty-card';
      card.innerHTML = '<div class="dc-role">' + escapeHtml(role) +
        '<span class="dc-count">' + byRole[role].length + '명</span></div>';
      var names = document.createElement('div');
      names.className = 'dc-names';
      byRole[role].forEach(function (s) {
        var chip = document.createElement('span');
        chip.textContent = s.id + ' ' + s.name;
        names.appendChild(chip);
      });
      card.appendChild(names);
      box.appendChild(card);
    });
  }

  function print() {
    var duties = data();
    if (!Object.keys(duties).length) { showToast('먼저 배정해 주세요', 'error'); return; }
    var cls = allClassData[currentClassIndex];
    var name = (cls && cls.className) || ((currentClassIndex + 1) + '반');
    var byRole = {};
    students.forEach(function (s) {
      var r = duties[s.id];
      if (!r) return;
      if (!byRole[r]) byRole[r] = [];
      byRole[r].push(s.id + ' ' + s.name);
    });
    var html = '<div class="print-header"><h1>' + name + ' 1인 1역</h1><p>' + formatDate(new Date()) + '</p></div>';
    html += '<table class="print-duty"><thead><tr><th>역할</th><th>담당 학생</th></tr></thead><tbody>';
    Object.keys(byRole).forEach(function (r) {
      html += '<tr><td>' + escapeHtml(r) + '</td><td>' + byRole[r].map(escapeHtml).join(', ') + '</td></tr>';
    });
    html += '</tbody></table>';
    document.getElementById('printArea').innerHTML = html;
    setTimeout(function () { window.print(); }, 100);
  }

  // ── 엑셀 양식 내려받기 (두 가지 방식) ──
  function downloadTemplate() {
    var d = allClassData[currentClassIndex] || {};
    var name = d.className || ((currentClassIndex + 1) + '반');

    // ① 직접 배정 — 학번·이름은 명렬표에서 채워 둡니다
    var s1 = [['학번', '이름', '역할']];
    if (students.length) {
      students.forEach(function (s) { s1.push([s.id, s.name, '']); });
    } else {
      s1.push(['1', '홍길동', '칠판 담당'], ['2', '김영희', '분리수거']);
    }

    // ② 역할 목록 — 역할만 적으면 무작위로 나눠 줍니다
    var s2 = [['역할'], ['칠판 담당'], ['분리수거'], ['창문 담당'], ['우유 담당'], ['책상 줄맞추기']];

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s1), '①직접배정');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s2), '②역할목록');
    XLSX.writeFile(wb, name + ' 1인1역 양식.xlsx');
    showToast('양식을 내려받았습니다');
  }

  // ── 엑셀 업로드 ──
  function handleFile(file) {
    if (!students.length) { showToast('명렬표를 먼저 등록해 주세요', 'error'); return; }
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        var direct = {}, roleList = [], matched = 0;

        wb.SheetNames.forEach(function (sn) {
          var rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1 });
          if (!rows.length) return;
          var head = (rows[0] || []).map(function (c) { return String(c || '').trim(); });
          var iId = head.indexOf('학번'), iNm = head.indexOf('이름'), iRo = head.indexOf('역할');
          if (iRo === -1) return;

          for (var r = 1; r < rows.length; r++) {
            var row = rows[r] || [];
            var role = String(row[iRo] == null ? '' : row[iRo]).trim();
            if (!role) continue;

            // 학번이나 이름이 있으면 직접 배정, 없으면 역할 목록으로 봅니다
            var id = iId >= 0 ? String(row[iId] == null ? '' : row[iId]).trim() : '';
            var nm = iNm >= 0 ? String(row[iNm] == null ? '' : row[iNm]).trim() : '';
            if (id || nm) {
              var st = students.find(function (s) {
                return (id && s.id === id) || (!id && nm && s.name === nm);
              });
              if (st) { direct[st.id] = role; matched++; }
            } else if (roleList.indexOf(role) === -1) {
              roleList.push(role);
            }
          }
        });

        var d = allClassData[currentClassIndex];
        if (matched) {
          d.duties = direct;
          d.dutyRoles = Object.keys(direct).map(function (k) { return direct[k]; })
            .filter(function (v, i, a) { return a.indexOf(v) === i; }).join('\n');
          saveCurrentClassData();
          document.getElementById('dutyRoles').value = d.dutyRoles;
          render();
          renderSeating();
          Sound.init(); Sound.playFanfare();
          showToast(matched + '명 역할을 불러왔습니다');
        } else if (roleList.length) {
          document.getElementById('dutyRoles').value = roleList.join('\n');
          showToast('역할 ' + roleList.length + '개를 불러왔습니다. 무작위 배정을 눌러주세요');
        } else {
          showToast('읽을 내용이 없습니다. 양식을 확인해 주세요', 'error');
        }
      } catch (err) {
        showToast('엑셀을 읽지 못했습니다', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function open() {
    document.getElementById('dutyRoles').value = roleText() ||
      '칠판 담당\n분리수거\n창문 담당\n우유 담당\n책상 줄맞추기';
    render();
    openModal('dutyModal');
  }

  function init() {
    var b = document.getElementById('dutyBtn'); if (b) b.addEventListener('click', open);
    var a = document.getElementById('dutyAssignBtn'); if (a) a.addEventListener('click', assign);
    var c = document.getElementById('dutyClearBtn'); if (c) c.addEventListener('click', clear);
    var p = document.getElementById('dutyPrintBtn'); if (p) p.addEventListener('click', print);
    var tp = document.getElementById('dutyTemplateBtn'); if (tp) tp.addEventListener('click', downloadTemplate);
    var up = document.getElementById('dutyUploadBtn');
    var fi = document.getElementById('dutyFileInput');
    if (up && fi) {
      up.addEventListener('click', function () { fi.value = ''; fi.click(); });
      fi.addEventListener('change', function () { if (this.files[0]) handleFile(this.files[0]); });
    }
  }

  return { init: init, render: render, get: function (id) { return data()[id] || ''; } };
})();


// ===================================================================
// 관찰·세특 기록 — 학급 + 학생번호를 키로 연결하는 다건 관계형 데이터
// ===================================================================
var SpecialNotes = (function () {
  var editingId = '';
  var LAST_SUBJECT_KEY = 'notetongLastSpecialSubject_v1';
  function el(id) { return document.getElementById(id); }
  function cls() { return allClassData[currentClassIndex] || {}; }
  function data() {
    var d = cls();
    if (!Array.isArray(d.specialNotes)) d.specialNotes = [];
    return d.specialNotes;
  }
  function compare(a, b) {
    return String(b.date || '').localeCompare(String(a.date || '')) || String(b.updatedAt || b.createdAt || '').localeCompare(String(a.updatedAt || a.createdAt || ''));
  }
  function forStudent(studentId) {
    return data().filter(function (item) { return String(item.studentId) === String(studentId); }).sort(compare);
  }
  function resetForm(keepSubject) {
    editingId = '';
    if (el('specialNoteId')) el('specialNoteId').value = '';
    if (el('specialNoteDate')) el('specialNoteDate').value = formatDate(new Date());
    if (el('specialNoteCategory')) el('specialNoteCategory').value = '수업 참여';
    if (el('specialNoteContent')) el('specialNoteContent').value = '';
    if (el('specialNoteSubject') && !keepSubject) {
      try { el('specialNoteSubject').value = localStorage.getItem(LAST_SUBJECT_KEY) || ''; } catch (e) { el('specialNoteSubject').value = ''; }
    }
    if (el('specialNoteSaveBtn')) el('specialNoteSaveBtn').textContent = '＋ 기록 저장';
    if (el('specialNoteCancelBtn')) el('specialNoteCancelBtn').hidden = true;
  }
  function render(studentId) {
    var box = el('specialNoteList'), count = el('specialNoteCount');
    if (!box) return;
    var list = forStudent(studentId);
    if (count) count.textContent = list.length + '건';
    if (!list.length) {
      box.innerHTML = '<p class="special-note-empty">아직 저장한 관찰·세특 기록이 없습니다.</p>';
      return;
    }
    box.innerHTML = list.map(function (item) {
      return '<article class="special-note-item">' +
        '<div class="special-note-meta"><span>' + escapeHtml(item.date || '') + '</span>' +
        '<b>' + escapeHtml(item.subject || '교과 미지정') + '</b><em>' + escapeHtml(item.category || '기타') + '</em></div>' +
        '<p>' + escapeHtml(item.content || '') + '</p>' +
        '<div class="special-note-actions"><button type="button" data-special-edit="' + escapeHtml(item.id) + '">수정</button>' +
        '<button type="button" data-special-delete="' + escapeHtml(item.id) + '">삭제</button></div></article>';
    }).join('');
    box.querySelectorAll('[data-special-edit]').forEach(function (button) {
      button.addEventListener('click', function () { edit(this.dataset.specialEdit); });
    });
    box.querySelectorAll('[data-special-delete]').forEach(function (button) {
      button.addEventListener('click', function () { remove(this.dataset.specialDelete); });
    });
  }
  function openForStudent(student) {
    if (!student) return;
    resetForm(false);
    render(student.id);
  }
  function activate() {
    document.querySelectorAll('#studentModal .tab-btn').forEach(function (b) { b.classList.toggle('active', b.dataset.tab === 'special'); });
    document.querySelectorAll('#studentModal .tab-content').forEach(function (c) { c.classList.toggle('active', c.id === 'specialTab'); });
    setTimeout(function () { if (el('specialNoteContent')) el('specialNoteContent').focus(); }, 30);
  }
  function save() {
    if (!selectedStudent) return;
    var date = el('specialNoteDate').value || formatDate(new Date());
    var subject = el('specialNoteSubject').value.trim();
    var category = el('specialNoteCategory').value || '기타';
    var content = el('specialNoteContent').value.trim();
    if (!subject) { showToast('교과를 입력해 주세요', 'error'); el('specialNoteSubject').focus(); return; }
    if (!content) { showToast('관찰 내용을 입력해 주세요', 'error'); el('specialNoteContent').focus(); return; }
    var list = data(), now = new Date().toISOString();
    var index = list.findIndex(function (item) { return item.id === editingId; });
    var record = {
      id: editingId || generateId(), studentId: selectedStudent.id, date: date,
      subject: subject, category: category, content: content,
      createdAt: index >= 0 ? (list[index].createdAt || now) : now, updatedAt: now
    };
    if (index >= 0) list[index] = record; else list.push(record);
    try { localStorage.setItem(LAST_SUBJECT_KEY, subject); } catch (e) { /* 선택 교과 저장 실패 무시 */ }
    saveCurrentClassData();
    resetForm(true);
    render(selectedStudent.id);
    renderStudentSummary(selectedStudent);
    showToast(index >= 0 ? '관찰·세특 기록을 수정했습니다' : '관찰·세특 기록을 저장했습니다');
  }
  function edit(id) {
    var item = data().find(function (record) { return record.id === id; });
    if (!item) return;
    editingId = item.id;
    el('specialNoteId').value = item.id;
    el('specialNoteDate').value = item.date || formatDate(new Date());
    el('specialNoteSubject').value = item.subject || '';
    el('specialNoteCategory').value = item.category || '기타';
    el('specialNoteContent').value = item.content || '';
    el('specialNoteSaveBtn').textContent = '수정 저장';
    el('specialNoteCancelBtn').hidden = false;
    el('specialNoteContent').focus();
  }
  function remove(id) {
    var list = data(), item = list.find(function (record) { return record.id === id; });
    if (!item || !confirm('이 관찰·세특 기록을 삭제할까요?')) return;
    cls().specialNotes = list.filter(function (record) { return record.id !== id; });
    saveCurrentClassData();
    resetForm(true);
    render(selectedStudent ? selectedStudent.id : '');
    if (selectedStudent) renderStudentSummary(selectedStudent);
    showToast('관찰·세특 기록을 삭제했습니다');
  }
  function count(studentId) { return forStudent(studentId).length; }
  function init() {
    if (el('specialNoteSaveBtn')) el('specialNoteSaveBtn').addEventListener('click', save);
    if (el('specialNoteCancelBtn')) el('specialNoteCancelBtn').addEventListener('click', function () { resetForm(true); });
    if (el('studentSpecialQuickBtn')) el('studentSpecialQuickBtn').addEventListener('click', activate);
  }
  return { init:init, openForStudent:openForStudent, activate:activate, forStudent:forStudent, count:count, all:function () { return data().slice(); } };
})();

// ===================================================================
// 통합 기록 — 생기부 참고용 관계형 데이터
// 학생 1명 = 1행. 자리이력 · 짝 · 출결 · 1인1역 · 칭찬 · 메모를 한 줄로 잇습니다.
// ===================================================================
var Records = (function () {
  function cls() { return allClassData[currentClassIndex] || {}; }

  function currentSeatOf(studentId) {
    var layout = getCurrentLayout();
    var text = '';
    Object.keys(currentSeats).forEach(function (sid) {
      var sd = currentSeats[sid];
      if (!sd || sd.studentId !== studentId) return;
      var p = sid.split('-');
      var row = parseInt(p[0], 10), col = parseInt(p[1], 10);
      var gi = getGroupIndexByCol(col, layout);
      text = config.seatType === 'group'
        ? (row * layout.length + gi + 1) + '모둠'
        : (gi + 1) + '분단 ' + (row + 1) + '행';
    });
    return text;
  }

  // 표 데이터 생성 (항목을 열로 분리 — 생기부 정리용)
  function build() {
    var d = cls();
    var recs = (d.savedRecords || []).slice().reverse();   // 오래된 것부터
    var head = ['학번', '이름', '성별', '1인 1역', '현재 좌석'];
    recs.forEach(function (r) { head.push('[' + r.memo + '] 자리'); });
    recs.forEach(function (r) { head.push('[' + r.memo + '] 짝'); });
    head.push('결석', '지각', '조퇴', '결과', '출결 사유·메모', '칭찬 포인트', '담임 메모', '관찰·세특 건수', '최근 관찰·세특', '관찰·세특 전체');

    var att = d.attendance || {}, rz = d.attendReasons || {}, am = d.attendMemos || {};
    var dates = Object.keys(att).sort();

    var rows = students.map(function (s) {
      var row = [s.id, s.name, s.gender || '',
                 (typeof Duty !== 'undefined' ? Duty.get(s.id) : '') || '',
                 currentSeatOf(s.id)];

      var hist = getSeatHistory(s.id);
      var byId = {};
      hist.forEach(function (h) { byId[h.memo] = h; });
      recs.forEach(function (r) {
        var h = byId[r.memo];
        row.push(h ? (h.group + '분단 ' + h.row + '행') : '');
      });
      recs.forEach(function (r) {
        var h = byId[r.memo];
        row.push(h && h.neighbors && h.neighbors.length
          ? h.neighbors.map(function (n) { return n.name; }).join(', ') : '');
      });

      var c = { '결석': 0, '지각': 0, '조퇴': 0, '결과': 0 };
      var notes = [];
      dates.forEach(function (dk) {
        var st = att[dk][s.id];
        if (!st || st === '출석') return;
        if (c[st] !== undefined) c[st]++;
        var r1 = (rz[dk] || {})[s.id], m1 = (am[dk] || {})[s.id];
        notes.push(dk.slice(5) + ' ' + st + (r1 ? '(' + r1 + ')' : '') + (m1 ? ' - ' + m1 : ''));
      });
      row.push(c['결석'], c['지각'], c['조퇴'], c['결과'], notes.join(' / '));
      row.push((typeof Praise !== 'undefined' ? Praise.get(s.id) : 0) || '');
      row.push(studentMemos[s.id] || '');
      var special = (Array.isArray(d.specialNotes) ? d.specialNotes : []).filter(function (item) {
        return String(item.studentId) === String(s.id);
      }).sort(function (a,b) { return String(b.date || '').localeCompare(String(a.date || '')); });
      var specialText = special.map(function (item) {
        return (item.date || '') + ' [' + (item.subject || '교과 미지정') + '·' + (item.category || '기타') + '] ' + (item.content || '');
      });
      row.push(special.length, specialText[0] || '', specialText.join('\n'));
      return row;
    });

    return { head: head, rows: rows };
  }


  function specialExportData() {
    var d = cls(), className = d.className || ((currentClassIndex + 1) + '반');
    var head = ['학급', '학번', '이름', '날짜', '교과', '기록 유형', '관찰 내용', '최종 수정'];
    var notes = Array.isArray(d.specialNotes) ? d.specialNotes.slice() : [];
    notes.sort(function (a,b) { return String(a.date || '').localeCompare(String(b.date || '')) || String(a.createdAt || '').localeCompare(String(b.createdAt || '')); });
    var rows = notes.map(function (item) {
      var student = students.find(function (s) { return String(s.id) === String(item.studentId); });
      return [className, item.studentId, student ? student.name : '', item.date || '', item.subject || '', item.category || '', item.content || '', item.updatedAt || item.createdAt || ''];
    });
    return { head:head, rows:rows };
  }

  function specialSummaryData() {
    var d = cls(), className = d.className || ((currentClassIndex + 1) + '반');
    var notes = Array.isArray(d.specialNotes) ? d.specialNotes : [];
    var head = ['학급', '학번', '이름', '기록 건수', '학생별 관찰·세특 정리'];
    var rows = students.map(function (student) {
      var list = notes.filter(function (item) { return String(item.studentId) === String(student.id); })
        .sort(function (a,b) { return String(a.date || '').localeCompare(String(b.date || '')); });
      return [className, student.id, student.name, list.length, list.map(function (item) {
        return (item.date || '') + ' [' + (item.subject || '교과 미지정') + '·' + (item.category || '기타') + '] ' + (item.content || '');
      }).join('\n')];
    });
    return { head:head, rows:rows };
  }

  function title() {
    var d = cls();
    return (d.className || ((currentClassIndex + 1) + '반')) + ' 통합기록 ' + formatDate(new Date());
  }

  function renderTable() {
    var data = build();
    var wrap = document.getElementById('recTableWrap');
    var cnt = document.getElementById('recCount');
    if (cnt) cnt.textContent = students.length + '명 · ' + data.head.length + '개 항목';
    if (!wrap) return;
    if (!students.length) { wrap.innerHTML = '<p class="hint">명렬표를 먼저 등록해 주세요</p>'; return; }

    var h = '<table class="rec-table"><thead><tr>';
    data.head.forEach(function (c) { h += '<th>' + escapeHtml(String(c)) + '</th>'; });
    h += '</tr></thead><tbody>';
    data.rows.forEach(function (r) {
      h += '<tr>';
      r.forEach(function (c) { h += '<td>' + escapeHtml(String(c == null ? '' : c)) + '</td>'; });
      h += '</tr>';
    });
    wrap.innerHTML = h + '</tbody></table>';
  }

  function toXlsx() {
    var data = build();
    if (!data.rows.length) { showToast('내보낼 자료가 없습니다', 'error'); return; }
    var ws = XLSX.utils.aoa_to_sheet([data.head].concat(data.rows));
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '통합기록');
    var raw = specialExportData();
    var rawWs = XLSX.utils.aoa_to_sheet([raw.head].concat(raw.rows));
    XLSX.utils.book_append_sheet(wb, rawWs, '관찰세특원본');
    var summary = specialSummaryData();
    var summaryWs = XLSX.utils.aoa_to_sheet([summary.head].concat(summary.rows));
    XLSX.utils.book_append_sheet(wb, summaryWs, '학생별세특정리');
    XLSX.writeFile(wb, title() + '.xlsx');
  }

  function toCsv() {
    var data = build();
    if (!data.rows.length) { showToast('내보낼 자료가 없습니다', 'error'); return; }
    var esc = function (v) {
      v = String(v == null ? '' : v);
      return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
    };
    var csv = [data.head].concat(data.rows).map(function (r) { return r.map(esc).join(','); }).join('\r\n');
    var blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });   // 엑셀 한글 깨짐 방지
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = title() + '.csv';
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }

  // 구글시트는 탭 구분 텍스트를 붙여넣으면 셀이 그대로 나뉩니다
  function toSheet() {
    var data = build();
    if (!data.rows.length) { showToast('내보낼 자료가 없습니다', 'error'); return; }
    var tsv = [data.head].concat(data.rows).map(function (r) {
      return r.map(function (v) { return String(v == null ? '' : v).replace(/[\t\n\r]/g, ' '); }).join('\t');
    }).join('\n');

    function done() {
      showToast('복사했습니다! 구글시트에서 Ctrl+V 하세요');
      window.open('https://sheets.new', '_blank', 'noopener');
    }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(tsv).then(done, fallback);
    } else fallback();

    function fallback() {
      var ta = document.createElement('textarea');
      ta.value = tsv;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand('copy'); done(); }
      catch (e) { showToast('복사에 실패했습니다', 'error'); }
      ta.remove();
    }
  }


  function toSpecialSheet() {
    var data = specialExportData();
    if (!data.rows.length) { showToast('복사할 관찰·세특 기록이 없습니다', 'error'); return; }
    var tsv = [data.head].concat(data.rows).map(function (row) {
      return row.map(function (value) { return String(value == null ? '' : value).replace(/[\t\n\r]/g, ' '); }).join('\t');
    }).join('\n');
    function done() { showToast('세특 원본을 복사했습니다! 구글시트에서 Ctrl+V 하세요'); window.open('https://sheets.new', '_blank', 'noopener'); }
    function fallback() {
      var ta=document.createElement('textarea');ta.value=tsv;ta.style.position='fixed';ta.style.opacity='0';document.body.appendChild(ta);ta.select();
      try { document.execCommand('copy'); done(); } catch (e) { showToast('복사에 실패했습니다', 'error'); }
      ta.remove();
    }
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(tsv).then(done, fallback); else fallback();
  }

  function open() { renderTable(); openModal('recordsModal'); }

  function init() {
    var b = document.getElementById('downloadHistoryBtn');
    if (b) { var n = b.cloneNode(true); b.parentNode.replaceChild(n, b); n.addEventListener('click', open); }
    var x = document.getElementById('recXlsxBtn'); if (x) x.addEventListener('click', toXlsx);
    var g = document.getElementById('recSheetBtn'); if (g) g.addEventListener('click', toSheet);
    var c = document.getElementById('recCsvBtn'); if (c) c.addEventListener('click', toCsv);
    var s = document.getElementById('recSpecialSheetBtn'); if (s) s.addEventListener('click', toSpecialSheet);
  }

  return { init: init, open: open };
})();

// ===================================================================
// 사다리타기
// ===================================================================
var Ladder = (function () {
  var names = [], prizes = [], rungs = [], cols = 4;
  var W = 720, H = 240, TOP = 34, BOT = 206;

  function el(id) { return document.getElementById(id); }
  function colX(i) { return 46 + i * ((W - 92) / Math.max(1, cols - 1)); }

  function readLines(id, n, fallback) {
    var raw = (el(id).value || '').split('\n').map(function (t) { return t.trim(); });
    var out = [];
    for (var i = 0; i < n; i++) out.push(raw[i] || (fallback + (i + 1)));
    return out;
  }

  function build() {
    cols = Math.max(2, Math.min(8, parseInt(el('ladderCount').value, 10) || 4));
    names = readLines('ladderNames', cols, '참여자');
    prizes = readLines('ladderPrizes', cols, '결과');

    // 가로줄: 같은 높이에 이웃끼리 겹치지 않게
    rungs = [];
    var levels = 8 + cols;
    for (var L = 0; L < levels; L++) {
      var y = TOP + 22 + (BOT - TOP - 40) * (L / (levels - 1));
      var used = {};
      for (var c = 0; c < cols - 1; c++) {
        if (used[c] || used[c - 1]) continue;
        if (Math.random() < 0.42) { rungs.push({ y: y, c: c }); used[c] = true; }
      }
    }
    draw(null);
    el('ladderResult').innerHTML = '<p class="hint">위쪽 이름을 누르면 길을 따라갑니다.</p>';
  }

  function draw(path) {
    var cv = el('ladderCanvas');
    if (!cv) return;
    var g = cv.getContext('2d');
    g.clearRect(0, 0, W, H);

    g.strokeStyle = '#cdd3e8'; g.lineWidth = 3; g.lineCap = 'round';
    for (var i = 0; i < cols; i++) {
      g.beginPath(); g.moveTo(colX(i), TOP); g.lineTo(colX(i), BOT); g.stroke();
    }
    rungs.forEach(function (r) {
      g.beginPath(); g.moveTo(colX(r.c), r.y); g.lineTo(colX(r.c + 1), r.y); g.stroke();
    });

    if (path && path.length) {
      g.strokeStyle = '#f06292'; g.lineWidth = 5;
      g.beginPath(); g.moveTo(path[0].x, path[0].y);
      for (var p = 1; p < path.length; p++) g.lineTo(path[p].x, path[p].y);
      g.stroke();
    }

    g.font = '600 13px sans-serif'; g.textAlign = 'center'; g.fillStyle = '#2f3550';
    for (var n = 0; n < cols; n++) {
      g.fillText(names[n] || '', colX(n), TOP - 12);
      g.fillText(prizes[n] || '', colX(n), BOT + 22);
    }
  }

  function trace(start) {
    var c = start, y = TOP;
    var path = [{ x: colX(c), y: TOP }];
    var sorted = rungs.slice().sort(function (a, b) { return a.y - b.y; });
    sorted.forEach(function (r) {
      if (r.c === c) {
        path.push({ x: colX(c), y: r.y });
        c = c + 1;
        path.push({ x: colX(c), y: r.y });
      } else if (r.c === c - 1) {
        path.push({ x: colX(c), y: r.y });
        c = c - 1;
        path.push({ x: colX(c), y: r.y });
      }
    });
    path.push({ x: colX(c), y: BOT });
    return { path: path, end: c };
  }

  function run(start) {
    if (!names.length) { build(); }
    var t = trace(start);
    // 한 점씩 그려 애니메이션
    var i = 2;
    Sound.init();
    (function step() {
      draw(t.path.slice(0, i));
      i++;
      if (i <= t.path.length) setTimeout(step, 45);
      else {
        Sound.playFanfare();
        Confetti.burstAt(el('ladderCanvas'), 30);
        el('ladderResult').innerHTML =
          '<div class="lr-hit"><b>' + escapeHtml(names[start]) + '</b> → <span>' +
          escapeHtml(prizes[t.end]) + '</span></div>';
      }
    })();
  }

  function runAll() {
    if (!names.length) build();
    var rows = names.map(function (n, i) {
      return '<div class="lr-row"><b>' + escapeHtml(n) + '</b><span>' +
        escapeHtml(prizes[trace(i).end]) + '</span></div>';
    });
    el('ladderResult').innerHTML = rows.join('');
    draw(null);
    Sound.init(); Sound.playFanfare();
  }

  function pickFromStudents() {
    if (!students.length) { showToast('명렬표를 먼저 등록해 주세요', 'error'); return; }
    var n = Math.max(2, Math.min(8, parseInt(el('ladderCount').value, 10) || 4));
    var pool = students.slice();
    shuffleArray(pool);
    el('ladderNames').value = pool.slice(0, n).map(function (s) { return s.name; }).join('\n');
    build();
  }

  function onCanvasClick(e) {
    var cv = el('ladderCanvas');
    var r = cv.getBoundingClientRect();
    var x = (e.clientX - r.left) * (W / r.width);
    var y = (e.clientY - r.top) * (H / r.height);
    if (y > TOP + 30) return;
    var best = 0, bd = 1e9;
    for (var i = 0; i < cols; i++) {
      var d = Math.abs(colX(i) - x);
      if (d < bd) { bd = d; best = i; }
    }
    if (bd < 60) run(best);
  }

  function open() {
    if (!el('ladderNames').value.trim() && students.length) pickFromStudents();
    else build();
    openModal('ladderModal');
  }

  function init() {
    var b = el('ladderBtn'); if (b) b.addEventListener('click', open);
    var mk = el('ladderMakeBtn'); if (mk) mk.addEventListener('click', build);
    var pk = el('ladderPickBtn'); if (pk) pk.addEventListener('click', pickFromStudents);
    var mi = el('ladderMinus'); if (mi) mi.addEventListener('click', function () {
      el('ladderCount').value = Math.max(2, (parseInt(el('ladderCount').value, 10) || 4) - 1); build();
    });
    var pl = el('ladderPlus'); if (pl) pl.addEventListener('click', function () {
      el('ladderCount').value = Math.min(8, (parseInt(el('ladderCount').value, 10) || 4) + 1); build();
    });
    var cv = el('ladderCanvas'); if (cv) cv.addEventListener('click', onCanvasClick);
    var lk = el('linksBtn'); if (lk) lk.addEventListener('click', function () { openModal('linksModal'); });
  }

  return { init: init, runAll: runAll };
})();

// ===================================================================
// 추첨 시작 스플래시 — 3 · 2 · 1 카운트다운
// ===================================================================
function runLotterySplash(done) {
  var splash = document.getElementById('lotterySplash');
  if (!splash) { done(); return; }

  var cls = allClassData[currentClassIndex];
  document.getElementById('splashClass').textContent =
    (cls && cls.className) ? cls.className : ((currentClassIndex + 1) + '반');
  document.getElementById('splashSub').textContent = students.length + '명의 새로운 자리';

  var countEl = document.getElementById('splashCount');
  var n = 3;
  var finished = false;
  var timer = null;

  function finish() {
    if (finished) return;
    finished = true;
    clearInterval(timer);
    splash.classList.remove('show');
    document.removeEventListener('keydown', onKey);
    done();
  }
  function onKey(e) { if (e.key === 'Escape' || e.key === 'Enter') finish(); }

  splash.classList.add('show');
  countEl.textContent = n;
  countEl.classList.remove('pulse');
  void countEl.offsetWidth;
  countEl.classList.add('pulse');
  Sound.init();
  Sound.playCountdown(n);

  timer = setInterval(function () {
    n--;
    if (n <= 0) {
      countEl.textContent = 'GO!';
      countEl.classList.remove('pulse');
      void countEl.offsetWidth;
      countEl.classList.add('pulse', 'go');
      Sound.playCountdown(0);
      Confetti.celebrate();
      setTimeout(finish, 700);
      clearInterval(timer);
      return;
    }
    countEl.textContent = n;
    countEl.classList.remove('pulse');
    void countEl.offsetWidth;
    countEl.classList.add('pulse');
    Sound.playCountdown(n);
  }, 850);

  document.getElementById('splashSkipBtn').onclick = finish;
  document.addEventListener('keydown', onKey);
}

// ===================================================================
// 시작 화면 단계 진행 (1 교실 크기 → 2 자리 형태 → 3 학생 수)
// ===================================================================
var wizardStep = 1;
var WIZARD_LAST = 3;

function showWizardStep(n) {
  wizardStep = Math.max(1, Math.min(WIZARD_LAST, n));

  document.querySelectorAll('.wizard-content .step').forEach(function (el) {
    var s = parseInt(el.dataset.step, 10);
    if (!s) return;
    // 레이아웃/모둠 인원 단계는 자체 표시 규칙이 있으므로 건드리지 않습니다
    if (el.id === 'layoutStep') {
      var avail = el.dataset.available !== '0';
      var onLayoutStep = s === wizardStep;
      el.style.display = onLayoutStep ? '' : 'none';
      el.style.visibility = (onLayoutStep && avail) ? 'visible' : 'hidden';
      return;
    }
    el.style.display = (s === wizardStep) ? '' : 'none';
  });

  document.querySelectorAll('.step-tab').forEach(function (tab) {
    var s = parseInt(tab.dataset.step, 10);
    tab.classList.toggle('active', s === wizardStep);
    tab.classList.toggle('done', s < wizardStep);
  });

  var fill = document.getElementById('stepProgressFill');
  if (fill) fill.style.width = (wizardStep / WIZARD_LAST * 100) + '%';

  var prev = document.getElementById('stepPrevBtn');
  var next = document.getElementById('stepNextBtn');
  var apply = document.getElementById('applyConfigBtn');
  if (prev) prev.style.visibility = wizardStep === 1 ? 'hidden' : 'visible';
  if (next) next.style.display = wizardStep === WIZARD_LAST ? 'none' : '';
  if (apply) apply.style.display = wizardStep === WIZARD_LAST ? '' : 'none';

  if (wizardStep === 3) updateEmptySeatInfo();
}

// 시작 화면의 기능 버튼 → 해당 안내 슬라이드로 캐러셀 열기
// 좌측 접기 메뉴 — 상단 버튼을 대신 눌러주는 방식이라 기존 동작을 그대로 씁니다
var RAIL_KEY = 'seatRail_open';

function setRailOpen(open) {
  var rail = document.getElementById('sideRail');
  if (!rail) return;
  rail.classList.toggle('open', !!open);
  var tx = rail.querySelector('#railToggle .rail-tx');
  if (tx) tx.textContent = open ? '메뉴 접기' : '메뉴 펼치기';
  try { localStorage.setItem(RAIL_KEY, open ? '1' : '0'); } catch (e) { /* 무시 */ }
  scheduleFit();
}

function initSideRail() {
  var rail = document.getElementById('sideRail');
  if (!rail) return;

  var saved = '0';
  try { saved = localStorage.getItem(RAIL_KEY) || '0'; } catch (e) { /* 무시 */ }
  setRailOpen(saved === '1');

  var t = document.getElementById('railToggle');
  if (t) t.addEventListener('click', function () {
    setRailOpen(!rail.classList.contains('open'));
  });

  // 상단에 숨겨둔 원래 버튼을 대신 클릭합니다
  rail.querySelectorAll('[data-proxy]').forEach(function (b) {
    b.addEventListener('click', function () {
      var target = document.getElementById(this.dataset.proxy);
      if (target) target.click();
    });
  });

  // 데이터 · 기록 하위 메뉴
  var bk = document.getElementById('railBackup');
  var bkMenu = document.getElementById('railBackupMenu');
  if (bk && bkMenu) {
    bk.addEventListener('click', function (e) {
      e.stopPropagation();
      var menuOpen=!bkMenu.classList.contains('open');
      bkMenu.classList.toggle('open',menuOpen);
      bk.setAttribute('aria-expanded',menuOpen?'true':'false');
    });
    bkMenu.addEventListener('click', function () { bkMenu.classList.remove('open');bk.setAttribute('aria-expanded','false'); });
    document.addEventListener('click', function () { bkMenu.classList.remove('open');bk.setAttribute('aria-expanded','false'); });
  }
  var bkOpen = document.getElementById('railBackupOpen');
  if (bkOpen) bkOpen.addEventListener('click', function () {
    openBackupModal();
  });
}

// 좌석 설정 패널 아래 버튼 + 드롭다운
var SP_KEY = 'seatPanel_settingsCollapsed_v2';
var SP_FOLD_KEY = 'seatPanel_folded_v1';

function initSeatActions() {
  // 요약·성별·배치 순서·설정 변경을 함께 접고, 주요 동작은 남겨 둡니다.
  var panel = document.getElementById('seatPanel');
  var head = panel && panel.querySelector('.sp-head');
  if (panel && head) {
    var saved = '1';
    try { saved = localStorage.getItem(SP_KEY) || '1'; } catch (e) { /* 무시 */ }
    panel.classList.toggle('sp-collapsed', saved === '1');
    head.setAttribute('aria-expanded', saved === '1' ? 'false' : 'true');
    head.addEventListener('click', function () {
      var on = !panel.classList.contains('sp-collapsed');
      panel.classList.toggle('sp-collapsed', on);
      head.setAttribute('aria-expanded', on ? 'false' : 'true');
      try { localStorage.setItem(SP_KEY, on ? '1' : '0'); } catch (e) { /* 무시 */ }
      scheduleFit();
    });
  }

  // 우측 패널 전체 접기 — 좌측 메뉴처럼 좁은 띠로 줄이고 자주 쓰는 동작만 아이콘으로 남깁니다
  var fold = document.getElementById('spFoldBtn');
  if (panel && fold) {
    var setFolded = function (on) {
      panel.classList.toggle('sp-folded', on);
      fold.setAttribute('aria-expanded', on ? 'false' : 'true');
      fold.title = on ? '오른쪽 메뉴 펼치기' : '오른쪽 메뉴 접기';
      fold.querySelector('.sp-fold-ic').textContent = on ? '◀' : '▶';
      fold.querySelector('.sp-fold-tx').textContent = on ? '' : '메뉴 접기';
      try { localStorage.setItem(SP_FOLD_KEY, on ? '1' : '0'); } catch (e) { /* 무시 */ }
      scheduleFit();
    };
    var folded = '0';
    try { folded = localStorage.getItem(SP_FOLD_KEY) || '0'; } catch (e) { /* 무시 */ }
    setFolded(folded === '1');
    fold.addEventListener('click', function () { setFolded(!panel.classList.contains('sp-folded')); });
  }

  document.querySelectorAll('.seat-panel [data-proxy]').forEach(function (b) {
    b.addEventListener('click', function () {
      var t = document.getElementById(this.dataset.proxy);
      if (t) t.click();
    });
  });
  document.querySelectorAll('.sp-drop').forEach(function (d) {
    var t = d.querySelector('.drop-t');
    var m = d.querySelector('.sp-menu');
    if (!t || !m) return;
    t.addEventListener('click', function (e) {
      e.stopPropagation();
      document.querySelectorAll('.sp-menu.open').forEach(function (o) { if (o !== m) o.classList.remove('open'); });
      m.classList.toggle('open');
    });
    m.addEventListener('click', function (e) {
      e.stopPropagation();
      if (e.target.closest('button')) m.classList.remove('open');
    });
  });
  document.addEventListener('click', function () {
    document.querySelectorAll('.sp-menu.open').forEach(function (m) { m.classList.remove('open'); });
  });
}

function initHeroTags() {
  // 슬라이드 순서가 바뀌어도 어긋나지 않도록 제목 키워드로 찾습니다
  function slideIndexBy(keyword) {
    for (var i = 0; i < GUIDE_SLIDES.length; i++) {
      if (GUIDE_SLIDES[i].title.indexOf(keyword) !== -1) return i;
    }
    return 0;
  }
  document.querySelectorAll('.hero-tag[data-key]').forEach(function (t) {
    t.addEventListener('click', function () {
      openGuide();
      goToSlide(slideIndexBy(this.dataset.key || ''));
    });
  });
}

function initWizardSteps() {
  document.querySelectorAll('.step-tab').forEach(function (tab) {
    tab.addEventListener('click', function () {
      showWizardStep(parseInt(this.dataset.step, 10));
    });
  });
  var prev = document.getElementById('stepPrevBtn');
  var next = document.getElementById('stepNextBtn');
  if (prev) prev.addEventListener('click', function () { showWizardStep(wizardStep - 1); });
  if (next) next.addEventListener('click', function () { showWizardStep(wizardStep + 1); });
  showWizardStep(1);
}

// 설정 모달 단계 탭 (스크롤 없이 한 화면에 담기 위해)
function showSettingsTab(n) {
  document.querySelectorAll('.set-tab').forEach(function (t) {
    t.classList.toggle('active', t.dataset.set === String(n));
  });
  document.querySelectorAll('#settingsModal .si-group').forEach(function (g) {
    g.style.display = (g.dataset.set === String(n)) ? '' : 'none';
  });
}

function initSettingsTabs() {
  document.querySelectorAll('.set-tab').forEach(function (t) {
    t.addEventListener('click', function () { showSettingsTab(this.dataset.set); });
  });
  var open = document.getElementById('settingsBtn');
  if (open) open.addEventListener('click', function () { showSettingsTab(1); });
  showSettingsTab(1);
}

// ===================================================================
// 학생 종합 카드 — 이름을 누르면 그 학생의 모든 기록을 한눈에
// ===================================================================
function renderStudentSummary(student) {
  try { StudentTimetable.renderForStudent(student); } catch (e) { /* ignore */ }
  var box = document.getElementById('studentSummary');
  if (!box || !student) return;

  var cls = allClassData[currentClassIndex] || {};

  // 현재 좌석
  var seatText = '아직 배치되지 않았습니다';
  var layout = getCurrentLayout();
  Object.keys(currentSeats).forEach(function (sid) {
    var sd = currentSeats[sid];
    if (!sd || sd.studentId !== student.id) return;
    var parts = sid.split('-');
    var row = parseInt(parts[0], 10), col = parseInt(parts[1], 10);
    var gi = getGroupIndexByCol(col, layout);
    seatText = config.seatType === 'group'
      ? (row * layout.length + gi + 1) + '모둠'
      : (gi + 1) + '분단 ' + (row + 1) + '행';
  });

  var actualToday = formatDate(new Date());
  var duty = (typeof Duty !== 'undefined') ? Duty.get(student.id) : '';
  var pts = (typeof Praise !== 'undefined') ? Praise.get(student.id) : 0;
  var praiseStage = (typeof Praise !== 'undefined') ? Praise.stage(pts) : null;
  var att = (typeof Attend !== 'undefined') ? Attend.get(student.id, actualToday) : '출석';
  var reason = (typeof Attend !== 'undefined') ? Attend.getReason(student.id, actualToday) : '';
  var memo = studentMemos[student.id] || '';
  var specialCount = (typeof SpecialNotes !== 'undefined') ? SpecialNotes.count(student.id) : 0;

  // 이번 달 근태 요약
  var counts = { '결석': 0, '지각': 0, '조퇴': 0 };
  var recent = [];
  var att0 = cls.attendance || {};
  Object.keys(att0).sort().reverse().forEach(function (dk) {
    var st = att0[dk][student.id];
    if (!st || st === '출석') return;
    if (counts[st] !== undefined) counts[st]++;
    if (recent.length < 4) {
      var rz = ((cls.attendReasons || {})[dk] || {})[student.id];
      recent.push(dk.slice(5) + ' ' + st + (rz ? ' · ' + rz : ''));
    }
  });

  // 최근 짝
  var pairs = [];
  var hist = getSeatHistory(student.id);
  hist.slice(0, 2).forEach(function (h) {
    if (h.neighbors && h.neighbors.length) pairs.push(h.memo + ': ' + h.neighbors.join(', '));
  });

  function row(icon, label, value, cls2) {
    return '<div class="ss-row ' + (cls2 || '') + '"><span class="ss-ic">' + icon + '</span>' +
      '<span class="ss-lb">' + label + '</span>' +
      '<span class="ss-vl">' + value + '</span></div>';
  }

  var attClass = att === '출석' ? 'ok' : 'warn';
  box.innerHTML =
    '<div class="ss-head">' +
      '<span class="ss-no">' + student.id + '번</span>' +
      '<span class="ss-name">' + escapeHtml(student.name) + '</span>' +
      (student.gender ? '<span class="ss-gender ' + (student.gender === '남' ? 'm' : 'f') + '">' + student.gender + '</span>' : '') +
      (praiseStage ? '<span class="ss-pts" title="칭찬 성장 단계: ' + praiseStage.label + '">' + praiseStage.icon + ' ' + pts + '점</span>' : '') +
    '</div>' +
    row('🪑', '현재 좌석', escapeHtml(seatText)) +
    row('🧹', '1인 1역', duty ? escapeHtml(duty) : '<em>배정 없음</em>') +
    row('✅', '오늘 근태', escapeHtml(att) + (reason ? ' <em>· ' + escapeHtml(reason) + '</em>' : ''), attClass) +
    row('📊', '누적 근태',
        (counts['결석'] + counts['지각'] + counts['조퇴'])
          ? ('결석 ' + counts['결석'] + ' · 지각 ' + counts['지각'] + ' · 조퇴 ' + counts['조퇴'])
          : '<em>기록 없음</em>') +
    (recent.length ? '<div class="ss-list">' + recent.map(function (r) {
        return '<span>' + escapeHtml(r) + '</span>'; }).join('') + '</div>' : '') +
    row('📝', '담임 메모', memo ? escapeHtml(memo) : '<em>없음</em>') +
    row('✍️', '관찰·세특', specialCount ? (specialCount + '건 누적') : '<em>기록 없음</em>') +
    (pairs.length ? row('👥', '최근 짝', pairs.map(escapeHtml).join('<br>')) : '');
}

// ===================================================================
// 상단 모드 탭 — 자리배치 / 출석부
// ===================================================================
var appMode = 'seat';

function setMode(mode) {
  appMode = (mode === 'attend') ? 'attend' : 'seat';
  document.querySelectorAll('.mode-tab').forEach(function (t) {
    t.classList.toggle('active', t.dataset.mode === appMode);
  });
  var wizardOpen = students.length === 0;
  var stage = document.getElementById('stage');
  var view = document.getElementById('attendView');
  var wiz = document.getElementById('setupWizard');
  var bar = document.querySelector('.action-bar');

  if (appMode === 'attend') {
    setModeScreen(stage, false);
    if (wiz) { wiz.classList.add('hidden'); wiz.style.display = 'none'; }
    setModeScreen(bar, false);
    setModeScreen(view, true, 'flex');
    Attend.render();
  } else {
    setModeScreen(view, false);
    if (wizardOpen) {
      showSetupWizard();
    } else {
      hideSetupWizard();
      scheduleFit();
    }
  }
}

// 좌석 설정 요약 패널
function renderSeatPanel() {
  var box = document.getElementById('spBody');
  if (!box) return;
  var layout = getCurrentLayout();
  var seats = layout.reduce(function (a, b) { return a + b; }, 0) * config.rows;
  var typeName = { single: '시험대형 (1인)', pair: '짝꿍 (2인)', triple: '트리플 (3인)', group: '모둠' }[config.seatType] || '';
  var unit = config.seatType === 'group' ? '가로 모둠' : '열(분단)';
  var extra = config.seatType === 'group'
    ? (GROUP_SIZES[Math.min(config.layoutIndex, GROUP_SIZES.length - 1)] || 4) + '명씩'
    : layout.join('-');

  function row(k, v) { return '<div class="sp-row"><span>' + k + '</span><b>' + v + '</b></div>'; }
  box.innerHTML =
    row('자리 형태', typeName) +
    row('배치 방식', getNumberOrderMode() === 'row' ? '학번 가로순' : (getNumberOrderMode() === 'column' ? '학번 세로순' : '무작위')) +
    row(unit + ' 수', config.columns + (config.seatType === 'group' ? '모둠' : '열')) +
    row('배열', extra) +
    row('행 수', config.rows + '행') +
    row('전체 자리', seats + '석') +
    row('학생 수', students.length + '명') +
    row('빈자리', Math.max(0, seats - students.length) + '석');
}

// 헤더 탭 줄은 가로로 넘기는 영역(overflow-x:auto)이라 그 안의 드롭다운(교실 도구·추천 사이트·게임)이
// 잘려서 보이지 않습니다. 열린 메뉴를 화면 기준(position:fixed)으로 꺼내 버튼 바로 아래에 놓습니다.
function placeHeaderMenus() {
  document.querySelectorAll('#modeTabs .more-menu').forEach(function (menu) {
    menu.style.position = '';
    menu.style.left = '';
    menu.style.right = '';
    menu.style.top = '';
    menu.style.maxHeight = '';
    menu.style.overflowY = '';
    if (!menu.classList.contains('open')) return;
    if (getComputedStyle(menu).position === 'fixed') return;  // 좁은 화면은 CSS가 이미 화면 기준으로 띄웁니다.
    var wrap = menu.closest('.more-wrap');
    var button = wrap && wrap.querySelector('.mode-tab');
    if (!button) return;
    var rect = button.getBoundingClientRect();
    var top = Math.round(rect.bottom + 8);
    menu.style.position = 'fixed';
    menu.style.right = 'auto';
    menu.style.top = top + 'px';
    menu.style.maxHeight = Math.max(160, window.innerHeight - top - 12) + 'px';
    menu.style.overflowY = 'auto';
    var left = Math.min(rect.left, window.innerWidth - menu.offsetWidth - 8);
    menu.style.left = Math.max(8, Math.round(left)) + 'px';
  });
}

function initHeaderMenuPlacement() {
  var tabs = document.getElementById('modeTabs');
  if (!tabs || typeof MutationObserver === 'undefined') return;
  var observer = new MutationObserver(placeHeaderMenus);
  tabs.querySelectorAll('.more-menu').forEach(function (menu) {
    observer.observe(menu, { attributes: true, attributeFilter: ['class'] });
  });
  tabs.addEventListener('scroll', placeHeaderMenus, { passive: true });
  window.addEventListener('resize', placeHeaderMenus);
}

function initModeTabs() {
  initHeaderMenuPlacement();
  document.querySelectorAll('.mode-tab[data-mode]').forEach(function (t) {
    t.addEventListener('click', function () { setMode(this.dataset.mode); });
  });

  var home = document.getElementById('homeBtn');
  var homeMenu = document.getElementById('homeMenu');
  if (home && homeMenu) {
    home.addEventListener('click', function (e) {
      e.stopPropagation();
      ['extraMenu','favMenu','gameMenu','helpMenu','aiMenu'].forEach(function (id) { var menu=document.getElementById(id); if(menu)menu.classList.remove('open'); });
      homeMenu.classList.toggle('open');
    });
    homeMenu.addEventListener('click', function (e) { e.stopPropagation(); });
    document.addEventListener('click', function () { homeMenu.classList.remove('open'); });
    homeMenu.querySelectorAll('[data-home]').forEach(function (item) {
      item.addEventListener('click', function () {
        homeMenu.classList.remove('open');
        if (this.dataset.home === 'guide') openGuide();
        else if (this.dataset.home === 'backup') openBackupModal();
        else showHomeScreen();
      });
    });
  }

  var calendarBtn = document.getElementById('calendarBtn');
  if (calendarBtn) calendarBtn.addEventListener('click', function () { ScheduleCalendar.open(); });

  // 즐겨찾기 드롭다운
  var favBtn = document.getElementById('favBtn');
  var favMenu = document.getElementById('favMenu');
  if (favBtn && favMenu) {
    favBtn.addEventListener('click', function (e) {
      e.stopPropagation();
      var em = document.getElementById('extraMenu');
      var hm = document.getElementById('helpMenu');
      if (em) em.classList.remove('open');
      if (hm) hm.classList.remove('open');
      var am = document.getElementById('aiMenu'); if (am) am.classList.remove('open');
      favMenu.classList.toggle('open');
    });
    favMenu.addEventListener('click', function (e) { e.stopPropagation(); });
    document.addEventListener('click', function () { favMenu.classList.remove('open'); });
  }

  // 교실 게임 드롭다운
  var gameBtn = document.getElementById('gameBtn');
  var gameMenu = document.getElementById('gameMenu');
  if (gameBtn && gameMenu) {
    gameBtn.addEventListener('click', function (e) {
      e.stopPropagation();
      var em = document.getElementById('extraMenu');
      var hm = document.getElementById('helpMenu');
      if (em) em.classList.remove('open');
      if (favMenu) favMenu.classList.remove('open');
      if (hm) hm.classList.remove('open');
      var am = document.getElementById('aiMenu'); if (am) am.classList.remove('open');
      gameMenu.classList.toggle('open');
    });
    gameMenu.addEventListener('click', function (e) { e.stopPropagation(); });
    document.addEventListener('click', function () { gameMenu.classList.remove('open'); });
  }
  // AI 드롭다운 — 도우미 열기·활용 안내·담임 5분 멘트
  var aiBtn = document.getElementById('aiMenuBtn');
  var aiMenu = document.getElementById('aiMenu');
  if (aiBtn && aiMenu) {
    aiBtn.addEventListener('click', function (e) {
      e.stopPropagation();
      ['extraMenu','favMenu','gameMenu','helpMenu','homeMenu'].forEach(function (id) { var m = document.getElementById(id); if (m) m.classList.remove('open'); });
      var open = !aiMenu.classList.contains('open');
      aiMenu.classList.toggle('open', open);
      aiBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
    });
    aiMenu.addEventListener('click', function (e) {
      e.stopPropagation();
      if (e.target.closest('button, a')) { aiMenu.classList.remove('open'); aiBtn.setAttribute('aria-expanded', 'false'); }
    });
    document.addEventListener('click', function () { aiMenu.classList.remove('open'); aiBtn.setAttribute('aria-expanded', 'false'); });
  }
  var aiGuideBtn = document.getElementById('aiGuideOpenBtn');
  if (aiGuideBtn) aiGuideBtn.addEventListener('click', function () { openModal('aiGuideModal'); });
  document.querySelectorAll('[data-ai-guide]').forEach(function (tab) {
    tab.addEventListener('click', function () {
      var key = this.dataset.aiGuide;
      document.querySelectorAll('[data-ai-guide]').forEach(function (t) { t.classList.toggle('active', t === tab); t.setAttribute('aria-selected', t === tab ? 'true' : 'false'); });
      document.querySelectorAll('[data-ai-guide-panel]').forEach(function (p) { p.hidden = p.dataset.aiGuidePanel !== key; });
    });
  });

  // 교실운영 드롭다운
  var btn = document.getElementById('extraBtn');
  var menu = document.getElementById('extraMenu');
  if (btn && menu) {
    function positionMobileExtraMenu() {
      if (window.innerWidth > 900) {
        menu.style.removeProperty('--mobile-extra-top');
        return;
      }
      var header = btn.closest('.app-header');
      var headerBottom = header ? Math.ceil(header.getBoundingClientRect().bottom) : 104;
      var safeTop = Math.max(8, Math.min(headerBottom, window.innerHeight - 80));
      menu.style.setProperty('--mobile-extra-top', safeTop + 'px');
    }
    btn.addEventListener('click', function (e) {
      e.stopPropagation();
      if (favMenu) favMenu.classList.remove('open');
      var am = document.getElementById('aiMenu'); if (am) am.classList.remove('open');
      if (!menu.classList.contains('open')) positionMobileExtraMenu();
      menu.classList.toggle('open');
    });
    document.addEventListener('click', function () { menu.classList.remove('open'); });
    menu.addEventListener('click', function (e) { e.stopPropagation(); });
    window.addEventListener('resize', function () {
      if (menu.classList.contains('open')) positionMobileExtraMenu();
    });

    var OPEN = {
      teacherSchedule: function () { TeacherSchedule.open(); },
      studentTimetable: function () { StudentTimetable.openUploadModal(); },
      timer: function () { Timer.openStage(); },
      notice: function () { document.getElementById('noticeBtn').click(); },
      duty: function () { document.getElementById('dutyBtn').click(); },
      praise: function () { document.getElementById('praiseBtn').click(); },
      pick: function () { document.getElementById('pickStudentBtn').click(); },
      ladder: function () { document.getElementById('ladderBtn').click(); },
      links: function () { document.getElementById('linksBtn').click(); },
      calendar: function () { ScheduleCalendar.open(); }
    };
    menu.querySelectorAll('button[data-open]').forEach(function (b) {
      b.addEventListener('click', function () {
        menu.classList.remove('open');
        var fn = OPEN[this.dataset.open];
        if (fn) fn();
      });
    });
  }

  var edit = document.getElementById('spEditBtn');
  if (edit) edit.addEventListener('click', showSetupWizard);
}

function maybeAutoOpenGuide() {
  // ?guide=1 은 저장된 "다음부터 열지 않기" 설정과 관계없이 설명서를 엽니다.
  // ?guide=8 처럼 번호를 넣으면 해당 슬라이드부터 확인할 수 있습니다.
  var requested = null;
  try { requested = new URLSearchParams(location.search).get('guide'); } catch (e) { /* 무시 */ }
  if (requested !== null) {
    openGuide();
    var slideNo = parseInt(requested, 10);
    if (!isNaN(slideNo)) goToSlide(Math.max(0, slideNo - 1));
    return;
  }

  // 첫 방문에도 설명서를 자동으로 띄우지 않습니다. 홈 화면에서 사용자가 직접 엽니다.
}



// ===================================================================
// 2026 고교학점제 & 교무수첩: 교사용 시간표 & NEIS 학교·학급 연동 모듈
// ===================================================================


// ===================================================================
// 2026 고교학점제 & 교무수첩: NEIS 학교/학급 조회, 반별 개설 과목·시간표 선택, 동아리 선택, 비담임(교과전담) 지원 및 학사일정 연동 모듈
// ===================================================================
var TeacherSchedule = (function () {
  var SCHOOL_KEY = 'tongNeisSchool_v1';
  var CLASSES_KEY = 'tongNeisClasses_v1';
  var CLASSES_SCHOOL_KEY = 'tongNeisClassesSchool_v1';  // 학급 목록을 NEIS에서 실제로 받아 온 학교
  var HOMEROOM_KEY = 'tongHomeroomClass_v1';
  var NON_HOMEROOM_KEY = 'tongIsNonHomeroom_v1';
  var SUBJECT_CLASSES_KEY = 'tongSubjectClasses_v1';
  var TIMETABLE_KEY = 'tongTeacherTimetable_v1';
  // v1/v3 캐시에는 작년 시간표와 임의로 채운 교시가 섞여 있어 새 키로 바꿔 버립니다.
  var HOMEROOM_TT_KEY = 'tongHomeroomTimetable_v2';
  var CLASS_TT_CACHE_KEY = 'tongClassTimetableCache_v4';
  var CLASS_TT_CACHE_MS = 6 * 60 * 60 * 1000;  // 학급 시간표 재조회 간격
  var ACADEMIC_CACHE_KEY = 'tongNeisAcademicEvents_v2';
  var ACADEMIC_META_KEY = 'tongNeisAcademicSync_v1';
  var ACADEMIC_SYNC_MS = 6 * 60 * 60 * 1000;  // 학사일정 자동 재조회 간격
  var CUSTOM_CLUBS_KEY = 'tongCustomClubs_v1';
  var ACTIVE_CLUB_SLOT_KEY = 'tongActiveClubSlot_v1';

  var BUILTIN_SCHOOLS = [
    { ATPT_OFCDC_SC_CODE: 'J10', SD_SCHUL_CODE: '7530929', SCHUL_NM: '신천고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '경기도', ORG_RDNMA: '경기도 시흥시 신천로79번길 14' },
    { ATPT_OFCDC_SC_CODE: 'J10', SD_SCHUL_CODE: '7530175', SCHUL_NM: '소래고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '경기도', ORG_RDNMA: '경기도 시흥시 소래산길 49' },
    { ATPT_OFCDC_SC_CODE: 'J10', SD_SCHUL_CODE: '7530527', SCHUL_NM: '은행고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '경기도', ORG_RDNMA: '경기도 시흥시 은행로 245' },
    { ATPT_OFCDC_SC_CODE: 'J10', SD_SCHUL_CODE: '7530529', SCHUL_NM: '함현고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '경기도', ORG_RDNMA: '경기도 시흥시 함현로 49' },
    { ATPT_OFCDC_SC_CODE: 'J10', SD_SCHUL_CODE: '7531161', SCHUL_NM: '배곧고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '경기도', ORG_RDNMA: '경기도 시흥시 배곧4로 108' },
    { ATPT_OFCDC_SC_CODE: 'J10', SD_SCHUL_CODE: '7530528', SCHUL_NM: '장곡고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '경기도', ORG_RDNMA: '경기도 시흥시 장곡로53번길 12' },
    { ATPT_OFCDC_SC_CODE: 'J10', SD_SCHUL_CODE: '7530526', SCHUL_NM: '시흥고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '경기도', ORG_RDNMA: '경기도 시흥시 하중로 225' },
    { ATPT_OFCDC_SC_CODE: 'J10', SD_SCHUL_CODE: '7530083', SCHUL_NM: '수원고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '경기도', ORG_RDNMA: '경기도 수원시 팔달구 경수대로 468' },
    { ATPT_OFCDC_SC_CODE: 'J10', SD_SCHUL_CODE: '7530059', SCHUL_NM: '분당고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '경기도', ORG_RDNMA: '경기도 성남시 분당구 수내로 152' },
    { ATPT_OFCDC_SC_CODE: 'B10', SD_SCHUL_CODE: '7010083', SCHUL_NM: '서울고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '서울특별시', ORG_RDNMA: '서울특별시 서초구 효령로 197' },
    { ATPT_OFCDC_SC_CODE: 'B10', SD_SCHUL_CODE: '7010059', SCHUL_NM: '경기고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '서울특별시', ORG_RDNMA: '서울특별시 강남구 영동대로 643' },
    { ATPT_OFCDC_SC_CODE: 'B10', SD_SCHUL_CODE: '7010126', SCHUL_NM: '경기여자고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '서울특별시', ORG_RDNMA: '서울특별시 강남구 삼성로 29' },
    { ATPT_OFCDC_SC_CODE: 'B10', SD_SCHUL_CODE: '7010198', SCHUL_NM: '숙명여자고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '서울특별시', ORG_RDNMA: '서울특별시 강남구 남부순환로 2807' },
    { ATPT_OFCDC_SC_CODE: 'C10', SD_SCHUL_CODE: '7150096', SCHUL_NM: '부산고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '부산광역시', ORG_RDNMA: '부산광역시 동구 초량로40번길 23' },
    { ATPT_OFCDC_SC_CODE: 'D10', SD_SCHUL_CODE: '7240055', SCHUL_NM: '경북고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '대구광역시', ORG_RDNMA: '대구광역시 수성구 청호로 300' },
    { ATPT_OFCDC_SC_CODE: 'G10', SD_SCHUL_CODE: '7430032', SCHUL_NM: '대전고등학교', SCHUL_KND_SC_NM: '고등학교', LCTN_SC_NM: '대전광역시', ORG_RDNMA: '대전광역시 중구 대흥로 109' }
  ];

  // 동아리 운영 교시 슬롯 옵션 (NEIS 창체 편성 교시 및 일반 연속 교시)
  var CLUB_SLOT_OPTIONS = [
    { id: 'wed-7', label: '수 7교시 (NEIS 창체)', slots: [{ day: '수', period: 7 }] },
    { id: 'wed-67', label: '수 6·7교시 (연속)', slots: [{ day: '수', period: 6 }, { day: '수', period: 7 }] },
    { id: 'fri-6', label: '금 6교시', slots: [{ day: '금', period: 6 }] },
    { id: 'fri-67', label: '금 6·7교시 (연속)', slots: [{ day: '금', period: 6 }, { day: '금', period: 7 }] },
    { id: 'wed-5', label: '수 5교시', slots: [{ day: '수', period: 5 }] },
    { id: 'fri-5', label: '금 5교시', slots: [{ day: '금', period: 5 }] }
  ];

  // 기본 제공 학교 동아리 카탈로그 (교과처럼 클릭 시 즉시 시간표 반영)
  var BUILTIN_CLUBS = [
    { name: '창체(동아리)', category: 'NEIS 창체·기본', room: '각 교실/특별실' },
    { name: '창체(자율·자치)', category: 'NEIS 창체·기본', room: '각 교실' },
    { name: '창체(진로탐구)', category: 'NEIS 창체·기본', room: '진로활동실' },
    { name: '창체(봉사활동)', category: 'NEIS 창체·기본', room: '교내외' },
    { name: 'AI·코딩탐구반', category: '과학·IT·수학', room: '컴퓨터실' },
    { name: '융합과학실험반', category: '과학·IT·수학', room: '과학실험실' },
    { name: '수리논술·통계반', category: '과학·IT·수학', room: '수학전용실' },
    { name: '로봇·공학메이커반', category: '과학·IT·수학', room: '메이커스페이스' },
    { name: '의생명·보건의료반', category: '과학·IT·수학', room: '생물실' },
    { name: '환경생태탐구반', category: '과학·IT·수학', room: '지구과학실' },
    { name: '독서토론·인문반', category: '인문·사회·학술', room: '도서관' },
    { name: '시사경제·경영반', category: '인문·사회·학술', room: '사회전용실' },
    { name: '역사문화탐방반', category: '인문·사회·학술', room: '역사실' },
    { name: '영자신문·글로벌반', category: '인문·사회·학술', room: '어학실' },
    { name: '심리·교육학탐구반', category: '인문·사회·학술', room: '세미나실' },
    { name: '방송반(교내방송)', category: '예술·체육·미디어', room: '방송실' },
    { name: '밴드·실용음악반', category: '예술·체육·미디어', room: '음악실' },
    { name: '미술·웹툰디자인반', category: '예술·체육·미디어', room: '미술실' },
    { name: '영상·영화제작반', category: '예술·체육·미디어', room: '멀티미디어실' },
    { name: '농구·스포츠클럽반', category: '예술·체육·미디어', room: '체육관' },
    { name: '배드민턴·체육반', category: '예술·체육·미디어', room: '강당' }
  ];

  var WEEK_DAYS = ['월', '화', '수', '목', '금'];
  var selectedSchool = null;
  var availableClasses = [];
  var isNonHomeroom = false; // 담임 없음(비담임·교과전담) 여부
  var selectedHomeroom = null;
  var selectedSubjects = [];
  var activePickerClass = null; // 현재 과목 선택 박스에 표시 중인 학급
  var currentEditCell = null;
  var activeClubCategory = '전체';
  var activeClubSlotId = 'wed-7';
  var activeGrade = null;        // 담당 수업 추가 ①단계에서 고른 학년(숫자) 또는 'etc'(기타)
  var pendingSubjects = [];      // ③단계에서 골라 둔(아직 추가 전) 교과명
  var currentPickerData = null;  // ③단계에 표시 중인 학급 시간표

  function el(id) { return document.getElementById(id); }

  function readJson(key, fallback) {
    try {
      var v = JSON.parse(localStorage.getItem(key) || 'null');
      return v == null ? fallback : v;
    } catch (e) { return fallback; }
  }
  function writeJson(key, val) {
    try { localStorage.setItem(key, JSON.stringify(val)); } catch (e) {}
    if (typeof CloudSheetSync !== "undefined" && CloudSheetSync.queueProfileSync) {
      if (key === SCHOOL_KEY || key === HOMEROOM_KEY || key === NON_HOMEROOM_KEY || key === SUBJECT_CLASSES_KEY || key === TIMETABLE_KEY || key === CUSTOM_CLUBS_KEY) {
        CloudSheetSync.queueProfileSync();
      }
    }
  }

  function ensureDefaultSchool() {
    selectedSchool = readJson(SCHOOL_KEY, null);
    if (!selectedSchool) {
      selectedSchool = {
        atptCode: 'J10',
        schoolCode: '7530929',
        schoolName: '신천고등학교',
        schoolKind: '고등학교',
        locationName: '경기도 시흥'
      };
      writeJson(SCHOOL_KEY, selectedSchool);
    }
    availableClasses = readJson(CLASSES_KEY, []);
    if (!availableClasses || !availableClasses.length) {
      availableClasses = [];
      for (var g = 1; g <= 3; g++) {
        for (var c = 1; c <= 10; c++) {
          availableClasses.push({
            grade: g,
            className: String(c),
            label: g + '학년 ' + c + '반',
            short: g + '-' + c
          });
        }
      }
      writeJson(CLASSES_KEY, availableClasses);
    }

    isNonHomeroom = Boolean(readJson(NON_HOMEROOM_KEY, false));
    if (isNonHomeroom) {
      selectedHomeroom = null;
    } else {
      selectedHomeroom = readJson(HOMEROOM_KEY, null);
      if (!selectedHomeroom) {
        selectedHomeroom = availableClasses.find(function (c) { return c.short === '2-2'; }) || availableClasses[0];
        writeJson(HOMEROOM_KEY, selectedHomeroom);
      }
    }

    selectedSubjects = readJson(SUBJECT_CLASSES_KEY, []);
    if (!selectedSubjects || !selectedSubjects.length) {
      var sub22 = availableClasses.find(function (c) { return c.short === '2-2'; });
      var sub28 = availableClasses.find(function (c) { return c.short === '2-8'; });
      selectedSubjects = [sub22, sub28].filter(Boolean);
      writeJson(SUBJECT_CLASSES_KEY, selectedSubjects);
    }
    var savedSlotId = readJson(ACTIVE_CLUB_SLOT_KEY, null);
    if (savedSlotId && CLUB_SLOT_OPTIONS.some(function (o) { return o.id === savedSlotId; })) {
      activeClubSlotId = savedSlotId;
    }
  }

  // 처음에는 빈 시간표로 시작합니다. (학급 시간표에서 과목을 눌러 채웁니다)
  function defaultTimetable() {
    var tt = {};
    WEEK_DAYS.forEach(function (d) {
      tt[d] = {};
      for (var p = 1; p <= 7; p++) {
        tt[d][p] = { className: '', subject: '', room: '' };
      }
    });
    return tt;
  }

  function fetchNeisRaw(setName, params) {
    return NeisApi.fetch(setName, params);
  }

  async function fetchNeis(setName, params) {
    var r = await fetchNeisRaw(setName, params);
    return r.rows;
  }

  function pad2(n) { return String(n).padStart(2, '0'); }
  function toYmd(date) { return String(date.getFullYear()) + pad2(date.getMonth() + 1) + pad2(date.getDate()); }
  function addDays(date, days) { var d = new Date(date.getFullYear(), date.getMonth(), date.getDate()); d.setDate(d.getDate() + days); return d; }
  function ymdLabel(ymd) { return Number(ymd.slice(4, 6)) + '/' + Number(ymd.slice(6, 8)); }
  // 이번 주 월요일 (주말에는 방금 끝난 주)
  function mondayOf(date) {
    var day = date.getDay();
    return addDays(date, day === 0 ? -6 : 1 - day);
  }
  // 학년도: 3월에 시작하므로 1~2월은 전년도
  function academicYearOf(date) {
    return date.getMonth() + 1 <= 2 ? date.getFullYear() - 1 : date.getFullYear();
  }

  // 1. 스마트 학교 검색
  async function searchSchool() {
    var input = el('tsSchoolSearchInput');
    var query = (input ? input.value : '').trim();
    var resultsBox = el('tsSchoolSearchResults');
    if (!resultsBox) return;
    if (query.length < 2) {
      resultsBox.innerHTML = '<p style="font-size:0.8rem; color:#dc2626; margin:4px 0;">학교명을 두 글자 이상 입력해 주세요.</p>';
      return;
    }
    resultsBox.innerHTML = '<p style="font-size:0.8rem; color:#4f46e5; margin:4px 0;">NEIS 전국 학교 실시간 검색 중...</p>';

    try {
      var qClean = query.replace(/\s+/g, '');
      var expanded = qClean
        .replace(/여고$/, '여자고등학교')
        .replace(/여중$/, '여자중학교')
        .replace(/과고$/, '과학고등학교')
        .replace(/외고$/, '외국어고등학교')
        .replace(/예고$/, '예술고등학교')
        .replace(/공고$/, '공업고등학교')
        .replace(/상고$/, '상업고등학교')
        .replace(/고$/, '고등학교')
        .replace(/중$/, '중학교')
        .replace(/초$/, '초등학교');

      var base = expanded.replace(/(여자고등학교|과학고등학교|외국어고등학교|고등학교|여자중학교|중학교|초등학교)$/, '');
      var querySet = [qClean];
      if (expanded !== qClean) querySet.push(expanded);
      if (base && base.length >= 2) {
        ['고등학교', '여자고등학교', '중학교', '초등학교'].forEach(function (suffix) {
          var cand = base + suffix;
          if (querySet.indexOf(cand) < 0) querySet.push(cand);
        });
      }

      var localMatches = BUILTIN_SCHOOLS.filter(function (s) {
        return s.SCHUL_NM.indexOf(qClean) >= 0 || s.SCHUL_NM.indexOf(expanded) >= 0 || (base.length >= 2 && s.SCHUL_NM.indexOf(base) >= 0);
      });

      var apiLists = await Promise.all(querySet.slice(0, 5).map(function (qItem) {
        return fetchNeis('schoolInfo', { SCHUL_NM: qItem, pSize: '20' }).catch(function () { return []; });
      }));

      var merged = [];
      var seen = {};
      localMatches.concat([].concat.apply([], apiLists)).forEach(function (r) {
        if (!r || !r.SD_SCHUL_CODE) return;
        var key = String(r.ATPT_OFCDC_SC_CODE || '') + ':' + String(r.SD_SCHUL_CODE);
        if (seen[key]) return;
        seen[key] = true;
        merged.push(r);
      });

      merged.sort(function (a, b) {
        var na = String(a.SCHUL_NM || '');
        var nb = String(b.SCHUL_NM || '');
        var exactA = (na === expanded || na === qClean) ? 0 : (na.indexOf(expanded) === 0 || na.indexOf(qClean) === 0 ? 1 : 2);
        var exactB = (nb === expanded || nb === qClean) ? 0 : (nb.indexOf(expanded) === 0 || nb.indexOf(qClean) === 0 ? 1 : 2);
        if (exactA !== exactB) return exactA - exactB;
        return na.localeCompare(nb, 'ko');
      });

      if (!merged.length) {
        resultsBox.innerHTML = '<p style="font-size:0.8rem; color:#64748b; margin:4px 0;">검색 결과가 없습니다. 학교명을 다시 확인해 주세요.</p>';
        return;
      }

      resultsBox.innerHTML = merged.map(function (r, i) {
        var name = r.SCHUL_NM || '';
        var kind = r.SCHUL_KND_SC_NM || '';
        var loc = r.LCTN_SC_NM || '';
        var addr = r.ORG_RDNMA || '';
        return '<div class="ts-school-item" data-index="' + i + '" style="padding:8px 12px; border:1px solid #e2e8f0; cursor:pointer; background:#fff; border-radius:8px; margin-bottom:6px; display:flex; justify-content:space-between; align-items:center;">' +
          '<div>' +
            '<strong style="color:#0f172a; font-size:0.88rem;">' + escapeHtml(name) + '</strong> ' +
            '<small style="color:#4f46e5; background:#eef2ff; padding:2px 6px; border-radius:4px; font-weight:700; margin-left:4px;">' + escapeHtml(kind) + '</small>' +
            '<div style="font-size:0.75rem; color:#64748b; margin-top:2px;">' + escapeHtml([loc, addr].filter(Boolean).join(' · ')) + '</div>' +
          '</div>' +
          '<span style="font-size:0.76rem; font-weight:700; color:#4f46e5; white-space:nowrap;">선택 &gt;</span>' +
        '</div>';
      }).join('');

      resultsBox.querySelectorAll('.ts-school-item').forEach(function (item) {
        item.addEventListener('click', function () {
          var idx = Number(this.dataset.index);
          var row = merged[idx];
          selectSchoolObject({
            atptCode: String(row.ATPT_OFCDC_SC_CODE || ''),
            schoolCode: String(row.SD_SCHUL_CODE || ''),
            schoolName: String(row.SCHUL_NM || ''),
            schoolKind: String(row.SCHUL_KND_SC_NM || ''),
            locationName: String(row.ORG_RDNMA || row.LCTN_SC_NM || '')
          });
          resultsBox.innerHTML = '';
        });
      });
    } catch (e) {
      resultsBox.innerHTML = '<p style="font-size:0.8rem; color:#dc2626; margin:4px 0;">학교 검색 중 오류가 발생했습니다.</p>';
    }
  }

  async function selectSchoolObject(schoolObj) {
    selectedSchool = schoolObj;
    writeJson(SCHOOL_KEY, selectedSchool);
    renderSelectedSchool();
    await loadSchoolClasses();
    await fetchHomeroomTimetable();
    await renderClassSubjectPicker(activePickerClass);
    renderClubPicker();
    await syncAcademicScheduleToCalendar(true);
    if (typeof TodayBriefing !== 'undefined' && TodayBriefing.refreshAll) {
      TodayBriefing.refreshAll();
    }
  }

  function renderSelectedSchool() {
    var badge = el('tsCurrentSchoolBadge');
    if (badge) {
      if (selectedSchool) {
        badge.textContent = selectedSchool.schoolName + ' (' + (selectedSchool.schoolKind || '학교') + ')';
        badge.style.color = '#15803d';
      } else {
        badge.textContent = '미설정';
        badge.style.color = '#64748b';
      }
    }
    var dashChip = el('dashCurrentSchoolText');
    if (dashChip && selectedSchool) {
      var roleSuffix = isNonHomeroom ? ' · 비담임(교과전담)' : (selectedHomeroom ? ' · ' + selectedHomeroom.short + ' 담임' : '');
      dashChip.textContent = selectedSchool.schoolName + roleSuffix;
    }
    var loadHomeroomBtn = el('tsLoadHomeroomTtBtn');
    if (loadHomeroomBtn) {
      loadHomeroomBtn.style.opacity = (isNonHomeroom || !selectedHomeroom) ? '0.55' : '1';
    }
  }

  // 2. 학년별 list_total_count 기반 전 학급 조회
  async function loadSchoolClasses() {
    if (!selectedSchool) return;
    var statusText = el('tsClassStatusText');
    if (statusText) statusText.textContent = 'NEIS 학년별 학급 수 조회 중...';

    var isElem = String(selectedSchool.schoolKind || '').indexOf('초등') >= 0;
    var maxGrade = isElem ? 6 : 3;
    var newClasses = [];

    function classEntry(gradeNum, name) {
      return { grade: gradeNum, className: String(name), label: gradeNum + '학년 ' + name + '반', short: gradeNum + '-' + name };
    }
    function byClassName(a, b) {
      var na = Number(a), nb = Number(b);
      if (!isNaN(na) && !isNaN(nb)) return na - nb;
      return String(a).localeCompare(String(b), 'ko');
    }

    var loadFailed = false;
    try {
      // 올해 학년도 → 아직 등록 전이면 지난 학년도 순서로 찾습니다.
      var thisYear = academicYearOf(new Date());
      var years = [String(thisYear), String(thisYear - 1)];
      var gradePromises = [];
      for (var g = 1; g <= maxGrade; g++) {
        (function (gradeNum) {
          gradePromises.push((async function () {
            for (var yi = 0; yi < years.length; yi++) {
              var res = await fetchNeisRaw('classInfo', {
                ATPT_OFCDC_SC_CODE: selectedSchool.atptCode,
                SD_SCHUL_CODE: selectedSchool.schoolCode,
                AY: years[yi],
                GRADE: String(gradeNum),
                pSize: '200'
              });
              if (!res.total && !res.rows.length) continue;
              var names = [];
              res.rows.forEach(function (row) {
                var name = String(row.CLASS_NM || '').trim();
                if (name && names.indexOf(name) < 0) names.push(name);
              });
              // 전체 반 목록을 다 받았으면 실제 반 이름을 쓰고,
              // 인증키가 없어 5개만 받았으면 총 개수(list_total_count)로 1~N반을 만듭니다.
              if (!res.truncated && names.length) return { grade: gradeNum, names: names.sort(byClassName) };
              var count = Math.min(Math.max(res.total, res.rows.length, 1), 30);
              var numbered = [];
              for (var n = 1; n <= count; n++) numbered.push(String(n));
              return { grade: gradeNum, names: numbered };
            }
            return { grade: gradeNum, names: [] };
          })());
        })(g);
      }
      var gradeLists = await Promise.all(gradePromises);
      gradeLists.sort(function (a, b) { return a.grade - b.grade; });
      gradeLists.forEach(function (item) {
        item.names.forEach(function (name) { newClasses.push(classEntry(item.grade, name)); });
      });
    } catch (err) {
      loadFailed = true;
    }

    // 조회에 실패했거나 등록된 학급이 없으면 이전 목록을 그대로 두고 사실대로 알립니다.
    if (loadFailed || !newClasses.length) {
      if (statusText) {
        statusText.textContent = loadFailed
          ? 'NEIS 학급 정보를 불러오지 못했습니다. 잠시 후 다시 검색해 주세요.'
          : selectedSchool.schoolName + ' · NEIS에 등록된 학급 정보가 없습니다.';
      }
      renderClassPills();
      return;
    }

    availableClasses = newClasses;
    writeJson(CLASSES_KEY, availableClasses);
    writeJson(CLASSES_SCHOOL_KEY, selectedSchool.atptCode + ':' + selectedSchool.schoolCode);
    // 학교를 바꿨을 때 새 학교에 없는 반이 수업 학급으로 남지 않도록 정리합니다.
    selectedSubjects = selectedSubjects.filter(function (s) {
      return s && availableClasses.some(function (x) { return x.short === s.short; });
    });
    writeJson(SUBJECT_CLASSES_KEY, selectedSubjects);
    if (!isNonHomeroom) {
      if (!selectedHomeroom || !availableClasses.some(function (x) { return x.short === selectedHomeroom.short; })) {
        selectedHomeroom = availableClasses.find(function (x) { return x.short === '2-2'; }) || availableClasses[0];
        writeJson(HOMEROOM_KEY, selectedHomeroom);
      }
    } else {
      selectedHomeroom = null;
    }
    activePickerClass = selectedHomeroom || selectedSubjects[0] || null;
    activeGrade = activePickerClass ? activePickerClass.grade : null;
    pendingSubjects = [];
    if (statusText) statusText.textContent = selectedSchool.schoolName + ' 총 ' + availableClasses.length + '개 학급 조회 완료';
    renderClassPills();
  }

  // ===== 담당 수업 추가: ① 학년 → ② 학급 → ③ 교과, 고른 학급·수업은 오른쪽 패널에 쌓입니다 =====
  function gradesAvailable() {
    var grades = [];
    availableClasses.forEach(function (c) { if (grades.indexOf(c.grade) < 0) grades.push(c.grade); });
    return grades.sort(function (a, b) { return a - b; });
  }

  function isHomeroomClass(short) {
    return !isNonHomeroom && Boolean(selectedHomeroom) && selectedHomeroom.short === short;
  }

  // 내 주간 시간표에 들어 있는 수업을 학급 → 과목별로 묶습니다. (오른쪽 패널의 원본)
  function collectLessons() {
    var tt = getTimetable();
    var groups = {};
    WEEK_DAYS.forEach(function (d) {
      for (var p = 1; p <= 7; p++) {
        var cell = tt[d] && tt[d][p];
        if (!cell || !(cell.subject || cell.className)) continue;
        var cls = cell.className || '';
        var sub = cell.subject || '수업';
        if (!groups[cls]) groups[cls] = {};
        if (!groups[cls][sub]) groups[cls][sub] = [];
        groups[cls][sub].push({ day: d, period: p, label: d + p });
      }
    });
    return groups;
  }

  function renderClassPills() {
    renderStepper();
    renderAssignedPanel();
  }

  // 교과 드롭다운을 '아직 고를 수 없음' 상태로 둡니다.
  function setSubjectPlaceholder(message) {
    var select = el('tsSubjectSelect');
    if (!select) return;
    select.innerHTML = '<option value="">' + escapeHtml(message) + '</option>';
    select.value = '';
    select.disabled = true;
  }

  // 학년 · 학급 드롭다운을 현재 선택 상태에 맞게 채웁니다. (교과 드롭다운은 renderSubjectOptions 가 채웁니다)
  function renderStepper() {
    var gradeSelect = el('tsGradeSelect');
    var classSelect = el('tsClassSelect');
    var subjectSelect = el('tsSubjectSelect');
    if (!gradeSelect || !classSelect) return;

    var grades = gradesAvailable();
    if (activeGrade !== 'etc' && grades.indexOf(activeGrade) < 0) activeGrade = null;
    gradeSelect.innerHTML = '<option value="">' + (grades.length ? '학년 선택' : '학교를 먼저 검색하세요') + '</option>' +
      grades.map(function (g) { return '<option value="' + g + '">' + g + '학년</option>'; }).join('') +
      (grades.length ? '<option value="etc">기타 (직접 입력 · 동아리)</option>' : '');
    gradeSelect.value = activeGrade == null ? '' : String(activeGrade);
    gradeSelect.disabled = !grades.length;

    var isEtc = activeGrade === 'etc';
    var gradeChosen = typeof activeGrade === 'number';
    var classChosen = gradeChosen && Boolean(activePickerClass) && activePickerClass.grade === activeGrade;
    classSelect.innerHTML = '<option value="">' + (gradeChosen ? '학급 선택' : '학년을 먼저 고르세요') + '</option>' +
      (gradeChosen ? availableClasses.filter(function (c) { return c.grade === activeGrade; }).map(function (c) {
        return '<option value="' + escapeHtml(c.short) + '">' + escapeHtml(c.className + '반' + (isHomeroomClass(c.short) ? ' (담임)' : '')) + '</option>';
      }).join('') : '');
    classSelect.value = classChosen ? activePickerClass.short : '';
    classSelect.disabled = !gradeChosen;

    // '기타'를 고르면 학급·교과 대신 추가 설정(직접 입력·동아리)이 열립니다.
    var etcStep = el('tsEtcStep');
    if (etcStep) etcStep.hidden = !isEtc;
    [classSelect.closest('.ts-pick-field'), subjectSelect && subjectSelect.closest('.ts-pick-field'), el('tsAddLessonsBtn'), el('tsPickHint')].forEach(function (node) {
      if (node) node.hidden = isEtc;
    });

    if (!classChosen) {
      // 학급을 아직 고르지 않았으면 그 반에 딸린 표시(교과·조회 상태·시간표 표)는 모두 접어 둡니다.
      setSubjectPlaceholder(gradeChosen ? '학급을 먼저 고르세요' : '학년·학급을 먼저 고르세요');
      ['tsNeisStatus', 'tsNeisKeyBox', 'tsMatrixToggleRow', 'tsClassMatrixPreviewWrap'].forEach(function (id) {
        var node = el(id);
        if (node) node.hidden = true;
      });
      var toggleBtn = el('tsToggleClassMatrixBtn');
      if (toggleBtn) toggleBtn.textContent = '학급 시간표 표 펼치기 ▾';
      var label = el('tsActiveClassLabel');
      if (label) label.textContent = '';
    }
  }

  // 오른쪽 패널: 고른 학급과 그 학급에 추가한 담당 수업
  function renderAssignedPanel() {
    var select = el('tsHomeroomSelect');
    if (select) {
      select.innerHTML = '<option value="NONE">담임 없음 (비담임·교과전담)</option>' + availableClasses.map(function (c) {
        return '<option value="' + escapeHtml(c.short) + '">' + escapeHtml(c.label) + '</option>';
      }).join('');
      select.value = (!isNonHomeroom && selectedHomeroom) ? selectedHomeroom.short : 'NONE';
    }

    var listEl = el('tsSelectedClassList');
    var countEl = el('tsLessonCount');
    if (!listEl) return;
    var groups = collectLessons();
    var lessonCount = 0;
    var shown = {};

    function lessonRows(cls) {
      var subs = groups[cls] || {};
      return Object.keys(subs).map(function (sub) {
        lessonCount++;
        return '<li class="ts-lesson-row">' +
          '<span><b>' + escapeHtml(sub) + '</b><small>' + escapeHtml(subs[sub].map(function (x) { return x.label; }).join(' · ')) + '</small></span>' +
          '<button type="button" class="ts-remove-btn" data-remove-lesson="' + escapeHtml(sub) + '" data-lesson-class="' + escapeHtml(cls) + '" aria-label="' + escapeHtml(sub) + ' 수업 삭제" title="이 수업 삭제">×</button>' +
        '</li>';
      }).join('');
    }

    var blocks = selectedSubjects.map(function (c) {
      shown[c.short] = true;
      var rows = lessonRows(c.short);
      var isCurrent = activeGrade !== 'etc' && Boolean(activePickerClass) && activePickerClass.short === c.short;
      return '<div class="ts-selected-class' + (isCurrent ? ' current' : '') + '">' +
        '<div class="ts-selected-class-head">' +
          '<button type="button" class="ts-class-jump" data-jump-class="' + escapeHtml(c.short) + '" title="이 학급의 교과 고르기">' +
            escapeHtml(c.label) + (isHomeroomClass(c.short) ? '<em>담임</em>' : '') +
          '</button>' +
          '<button type="button" class="ts-remove-btn" data-remove-class="' + escapeHtml(c.short) + '" aria-label="' + escapeHtml(c.label) + ' 빼기" title="이 학급 빼기">×</button>' +
        '</div>' +
        (rows ? '<ul>' + rows + '</ul>' : '<p class="ts-selected-hint">아직 추가한 교과가 없습니다.</p>') +
      '</div>';
    });

    // 고른 학급에 속하지 않는 수업(동아리·직접 입력 등)은 '기타'로 묶어 보여 줍니다.
    Object.keys(groups).forEach(function (cls) {
      if (shown[cls]) return;
      blocks.push('<div class="ts-selected-class etc">' +
        '<div class="ts-selected-class-head"><span class="ts-class-jump static">' + escapeHtml(cls || '학급 미지정') + '<em>기타</em></span></div>' +
        '<ul>' + lessonRows(cls) + '</ul>' +
      '</div>');
    });

    listEl.innerHTML = blocks.length ? blocks.join('') : '<p class="ts-panel-empty">아직 추가한 담당 수업이 없습니다.</p>';
    if (countEl) countEl.textContent = String(lessonCount);
  }

  function selectGrade(value) {
    activeGrade = value === 'etc' ? 'etc' : (value === '' || value == null ? null : Number(value));
    pendingSubjects = [];
    if (activeGrade === 'etc') {
      renderClassPills();
      renderClubPicker();
      return;
    }
    if (activePickerClass && activePickerClass.grade !== activeGrade) {
      activePickerClass = null;
      currentPickerData = null;
    }
    renderClassPills();
    if (activePickerClass) renderClassSubjectPicker(activePickerClass);
    else updateAddLessonsButton();
  }

  // 학급을 고르면 그 반의 이번 주 NEIS 시간표에서 교과 목록을 불러옵니다.
  // (오른쪽 '추가한 담당 수업'에는 교과를 [추가]했을 때 들어갑니다)
  function selectClass(short) {
    var item = availableClasses.find(function (c) { return c.short === short; });
    pendingSubjects = [];
    if (!item) {
      activePickerClass = null;
      currentPickerData = null;
      renderClassPills();
      updateAddLessonsButton();
      return;
    }
    activeGrade = item.grade;
    activePickerClass = item;
    renderClassPills();
    renderClassSubjectPicker(item);
  }

  function afterTimetableChanged() {
    renderWeeklyTimetable();
    renderClassPills();
    if (activePickerClass) renderClassSubjectPicker(activePickerClass);
    renderClubPicker();
    if (typeof TodayBriefing !== 'undefined' && TodayBriefing.refreshAll) TodayBriefing.refreshAll();
  }

  function clearLessonCells(tt, cls, sub) {
    var cleared = 0;
    WEEK_DAYS.forEach(function (d) {
      for (var p = 1; p <= 7; p++) {
        var cell = tt[d] && tt[d][p];
        if (!cell || !(cell.subject || cell.className)) continue;
        if ((cell.className || '') !== cls) continue;
        if (sub != null && (cell.subject || '수업') !== sub) continue;
        tt[d][p] = { className: '', subject: '', room: '' };
        cleared++;
      }
    });
    return cleared;
  }

  function removeLesson(cls, sub) {
    var tt = getTimetable();
    clearLessonCells(tt, cls, sub);
    writeJson(TIMETABLE_KEY, tt);
    afterTimetableChanged();
    if (typeof showToast === 'function') showToast('[' + (cls ? cls + ' ' : '') + sub + '] 수업을 담당 수업에서 뺐습니다.');
  }

  function removeSelectedClass(short) {
    var item = selectedSubjects.find(function (c) { return c.short === short; });
    if (!item) return;
    var lessonNames = Object.keys(collectLessons()[short] || {});
    if (lessonNames.length && !confirm(item.label + '과(와) 이 학급의 담당 수업 ' + lessonNames.length + '개를 목록과 내 시간표에서 뺄까요?')) return;
    var tt = getTimetable();
    clearLessonCells(tt, short, null);
    writeJson(TIMETABLE_KEY, tt);
    selectedSubjects = selectedSubjects.filter(function (c) { return c.short !== short; });
    writeJson(SUBJECT_CLASSES_KEY, selectedSubjects);
    if (activePickerClass && activePickerClass.short === short) {
      activePickerClass = null;
      pendingSubjects = [];
      currentPickerData = null;
    }
    afterTimetableChanged();
    updateAddLessonsButton();
  }

  async function setHomeroom(key) {
    if (key === 'NONE') {
      isNonHomeroom = true;
      selectedHomeroom = null;
      writeJson(NON_HOMEROOM_KEY, true);
      writeJson(HOMEROOM_KEY, null);
    } else {
      var found = availableClasses.find(function (c) { return c.short === key; });
      if (!found) return;
      isNonHomeroom = false;
      selectedHomeroom = found;
      writeJson(NON_HOMEROOM_KEY, false);
      writeJson(HOMEROOM_KEY, selectedHomeroom);
    }
    renderClassPills();
    renderSelectedSchool();
    await fetchHomeroomTimetable();
    if (typeof TodayBriefing !== 'undefined' && TodayBriefing.refreshAll) TodayBriefing.refreshAll();
    if (typeof showToast === 'function') {
      showToast(key === 'NONE'
        ? '담임 없음 (비담임 · 교과전담) 모드로 설정되었습니다.'
        : '담임 학급을 ' + selectedHomeroom.label + '(으)로 설정하고 우리 반 시간표를 불러왔습니다.');
    }
  }

  function updateAddLessonsButton() {
    var btn = el('tsAddLessonsBtn');
    if (!btn) return;
    btn.disabled = !pendingSubjects.length;
  }

  // 드롭다운에서 고른 교과를 내 주간 시간표의 해당 요일·교시(NEIS 시간표 그대로)에 넣습니다.
  function addPickedLessons() {
    var classObj = activePickerClass;
    var data = currentPickerData;
    if (!classObj || !data || !pendingSubjects.length) return;
    var tt = getTimetable();
    var added = [];
    var replaced = 0;
    pendingSubjects.forEach(function (name) {
      var item = data.subjects.find(function (x) { return x.subject === name; });
      if (!item) return;
      item.slots.forEach(function (sl) {
        if (!tt[sl.day]) tt[sl.day] = {};
        var cur = tt[sl.day][sl.period];
        if (cur && (cur.subject || cur.className) && !(cur.className === classObj.short && cur.subject === name)) replaced++;
        tt[sl.day][sl.period] = { className: classObj.short, subject: name, room: classObj.short };
      });
      added.push(name);
    });
    if (!added.length) return;
    writeJson(TIMETABLE_KEY, tt);
    pendingSubjects = [];
    if (!selectedSubjects.some(function (c) { return c.short === classObj.short; })) {
      selectedSubjects.push(classObj);
      writeJson(SUBJECT_CLASSES_KEY, selectedSubjects);
    }
    afterTimetableChanged();
    if (typeof showToast === 'function') {
      showToast(classObj.label + ' [' + added.join(', ') + '] 수업을 담당 수업에 추가했습니다.' +
        (replaced ? ' (같은 교시에 있던 다른 수업 ' + replaced + '칸을 바꿨습니다)' : ''));
    }
  }

  // 기타: NEIS 학급 시간표에 없는 수업을 요일·교시를 정해 직접 넣습니다.
  function addEtcLesson() {
    var classInput = el('tsEtcClassInput');
    var subjectInput = el('tsEtcSubjectInput');
    var daySelect = el('tsEtcDaySelect');
    var periodSelect = el('tsEtcPeriodSelect');
    if (!classInput || !subjectInput || !daySelect || !periodSelect) return;
    var cls = classInput.value.trim();
    var sub = subjectInput.value.trim();
    var day = daySelect.value;
    var period = Number(periodSelect.value);
    if (!sub) {
      subjectInput.focus();
      if (typeof showToast === 'function') showToast('수업명을 입력해 주세요.', 'error');
      return;
    }
    var tt = getTimetable();
    if (!tt[day]) tt[day] = {};
    var cur = tt[day][period];
    if (cur && (cur.subject || cur.className) &&
        !confirm(day + '요일 ' + period + '교시에 이미 [' + ((cur.className ? cur.className + ' ' : '') + (cur.subject || '수업')) + '] 수업이 있습니다. 바꿀까요?')) return;
    tt[day][period] = { className: cls, subject: sub, room: cls };
    writeJson(TIMETABLE_KEY, tt);
    afterTimetableChanged();
    if (typeof showToast === 'function') showToast(day + '요일 ' + period + '교시에 [' + sub + '] 수업을 추가했습니다.');
  }

  function timetableEndpoint() {
    var kind = String((selectedSchool && selectedSchool.schoolKind) || '');
    if (kind.indexOf('초등') >= 0) return 'elsTimetable';
    if (kind.indexOf('중학') >= 0) return 'misTimetable';
    if (kind.indexOf('특수') >= 0) return 'spsTimetable';
    return 'hisTimetable';
  }

  function classTimetableCacheKey(classObj) {
    return selectedSchool.atptCode + ':' + selectedSchool.schoolCode + ':' + classObj.short;
  }

  // 저장해 둔 학급 시간표 — 이번 주 것이고 6시간이 지나지 않았을 때만 다시 씁니다.
  function readClassTimetableCache(classObj) {
    var record = readJson(CLASS_TT_CACHE_KEY, {})[classTimetableCacheKey(classObj)];
    if (!record || !record.days) return null;
    if (record.currentWeek !== toYmd(mondayOf(new Date()))) return null;
    if (Date.now() - Number(record.fetchedAt || 0) > CLASS_TT_CACHE_MS) return null;
    return record;
  }

  function clearClassTimetableCache() {
    try {
      localStorage.removeItem(CLASS_TT_CACHE_KEY);
      localStorage.removeItem(HOMEROOM_TT_KEY);
    } catch (e) {}
  }

  // 특정 학급의 NEIS 주간 시간표 조회.
  //  - 이번 주(월~금)를 조회하고, 방학·연휴로 비어 있으면 최대 8주 전까지 거슬러 올라가 가장 최근 주를 씁니다.
  //  - NEIS 가 돌려준 교시만 채웁니다. 받지 못한 교시는 비워 두고 partialDays 에 기록합니다.
  //  반환: { days, subjects, weekStart, weekEnd, truncated, partialDays, keyed, fetchedAt } 또는 { error }
  async function fetchClassTimetableData(classObj, options) {
    ensureDefaultSchool();
    if (!selectedSchool || !classObj) return null;
    var force = Boolean(options && options.force);

    if (!force) {
      var cached = readClassTimetableCache(classObj);
      if (cached) return cached;
    }

    var endpoint = timetableEndpoint();
    var base = {
      ATPT_OFCDC_SC_CODE: selectedSchool.atptCode,
      SD_SCHUL_CODE: selectedSchool.schoolCode,
      GRADE: String(classObj.grade),
      CLASS_NM: String(classObj.className)
    };
    var currentMonday = mondayOf(new Date());
    var MAX_WEEKS_BACK = 8;

    try {
      // 1) 데이터가 있는 가장 최근 주 찾기 (한 주에 한 번 조회)
      var weekMonday = null;
      var weekResult = null;
      for (var back = 0; back <= MAX_WEEKS_BACK; back++) {
        var monday = addDays(currentMonday, -7 * back);
        var probe = await fetchNeisRaw(endpoint, Object.assign({}, base, {
          TI_FROM_YMD: toYmd(monday),
          TI_TO_YMD: toYmd(addDays(monday, 4)),
          pSize: '1000'
        }));
        if (probe.total > 0 || probe.rows.length > 0) {
          weekMonday = monday;
          weekResult = probe;
          break;
        }
      }
      if (!weekMonday) {
        return { error: '최근 ' + (MAX_WEEKS_BACK + 1) + '주 동안 NEIS에 등록된 ' + classObj.label + ' 시간표가 없습니다.' };
      }

      // 2) 주간 조회가 잘렸으면(인증키 없음) 요일별로 나눠 다시 조회 — 하루 앞의 5교시까지는 받을 수 있습니다.
      var dayTotals = {};
      var rows = weekResult.rows;
      if (weekResult.truncated) {
        var dayResults = await Promise.all(WEEK_DAYS.map(function (dName, idx) {
          return fetchNeisRaw(endpoint, Object.assign({}, base, {
            ALL_TI_YMD: toYmd(addDays(weekMonday, idx)),
            pSize: '100'
          }));
        }));
        rows = [];
        dayResults.forEach(function (res, idx) {
          dayTotals[WEEK_DAYS[idx]] = res.total;
          rows = rows.concat(res.rows);
        });
      }

      // 3) 요일·교시 표로 정리
      var daysMap = {};
      var loadedCount = {};
      WEEK_DAYS.forEach(function (d) { daysMap[d] = {}; loadedCount[d] = 0; });
      var weekYmds = WEEK_DAYS.map(function (d, idx) { return toYmd(addDays(weekMonday, idx)); });
      rows.forEach(function (r) {
        var dayIdx = weekYmds.indexOf(String(r.ALL_TI_YMD || ''));
        var period = Number(r.PERIO);
        var subject = String(r.ITRT_CNTNT || '').trim();
        if (dayIdx < 0) return;
        var dName = WEEK_DAYS[dayIdx];
        loadedCount[dName]++;
        if (!subject || !(period >= 1 && period <= 7)) return;
        var existing = daysMap[dName][period];
        if (!existing) daysMap[dName][period] = subject;
        else if (existing.split(' / ').indexOf(subject) < 0) daysMap[dName][period] = existing + ' / ' + subject;
      });

      // 요일별로 NEIS 에 있는 교시 수보다 덜 받은 날 (인증키가 없을 때만 생깁니다)
      var partialDays = [];
      WEEK_DAYS.forEach(function (d) {
        var total = dayTotals[d];
        if (total != null && total > loadedCount[d]) partialDays.push({ day: d, loaded: loadedCount[d], total: total });
      });

      // 과목별 수업 요일·교시 집계 (이 반 시간표에 실제로 있는 과목만)
      var subSlotMap = {};
      WEEK_DAYS.forEach(function (d) {
        for (var p = 1; p <= 7; p++) {
          var sName = daysMap[d][p];
          if (!sName) continue;
          if (!subSlotMap[sName]) subSlotMap[sName] = [];
          subSlotMap[sName].push({ day: d, period: p, label: d + p });
        }
      });
      var subjectsList = Object.keys(subSlotMap).map(function (sName) {
        return {
          subject: sName,
          slots: subSlotMap[sName],
          slotSummary: subSlotMap[sName].map(function (x) { return x.label; }).join(' · ')
        };
      });

      var record = {
        schoolName: selectedSchool.schoolName,
        classShort: classObj.short,
        classLabel: classObj.label,
        days: daysMap,
        subjects: subjectsList,
        weekStart: weekYmds[0],
        weekEnd: weekYmds[4],
        currentWeek: toYmd(currentMonday),
        isCurrentWeek: weekYmds[0] === toYmd(currentMonday),
        keyed: Boolean(weekResult.keyed),
        truncated: partialDays.length > 0,
        partialDays: partialDays,
        fetchedAt: Date.now()
      };
      var cacheAll = readJson(CLASS_TT_CACHE_KEY, {});
      cacheAll[classTimetableCacheKey(classObj)] = record;
      writeJson(CLASS_TT_CACHE_KEY, cacheAll);
      return record;
    } catch (e) {
      if (e && e.isKeyError) {
        return { error: 'NEIS 인증키를 사용할 수 없습니다. (' + e.message + ')', keyError: true };
      }
      return { error: 'NEIS 시간표를 불러오지 못했습니다. 네트워크 연결을 확인한 뒤 다시 시도해 주세요.' };
    }
  }

  // 2-1. 선택한 학급(예: 2-2, 2-8, 1-1)의 NEIS 개설 과목 선택 칩 + 학급 시간표 매트릭스 렌더링
  // 학급 시간표가 어느 주 기준인지, 빠진 교시가 있는지 사실대로 알려 주는 상태 줄
  function renderNeisStatus(data) {
    var statusEl = el('tsNeisStatus');
    var keyBox = el('tsNeisKeyBox');
    var keyInput = el('tsNeisKeyInput');
    var clearBtn = el('tsNeisKeyClearBtn');
    if (!statusEl) return;
    var hasUserKey = NeisApi.hasUserKey();
    var kind = 'ok';
    var html = '';

    if (!data || data.error) {
      kind = 'error';
      html = '<b>' + escapeHtml((data && data.error) || 'NEIS 시간표를 불러오지 못했습니다.') + '</b>';
    } else {
      var range = ymdLabel(data.weekStart) + '~' + ymdLabel(data.weekEnd);
      var time = new Date(data.fetchedAt);
      var stamp = pad2(time.getHours()) + ':' + pad2(time.getMinutes()) + ' 조회';
      html = data.isCurrentWeek
        ? '<b>이번 주(' + escapeHtml(range) + ') NEIS 시간표</b> · ' + stamp
        : '<b>이번 주는 NEIS에 등록된 시간표가 없어 가장 최근 주(' + escapeHtml(range) + ')를 표시합니다.</b> · ' + stamp;
      if (data.truncated) {
        kind = 'warn';
        var detail = data.partialDays.map(function (x) { return x.day + ' ' + x.total + '교시 중 ' + x.loaded; }).join(' · ');
        html += '<br>NEIS 인증키가 없어 하루 5교시까지만 받았습니다 (' + escapeHtml(detail) + '). 받지 못한 교시는 비워 두었습니다.';
      }
    }
    statusEl.className = 'ts-neis-status ' + kind;
    statusEl.innerHTML = '<span>' + html + '</span><button type="button" class="btn small" id="tsNeisRefreshBtn">다시 조회</button>';
    statusEl.hidden = false;
    var refreshBtn = el('tsNeisRefreshBtn');
    if (refreshBtn) refreshBtn.addEventListener('click', function () {
      renderClassSubjectPicker(activePickerClass, { force: true });
    });

    // 인증키 입력란: 시간표가 잘렸거나, 키 오류가 났거나, 이미 저장한 키가 있을 때만 보입니다.
    if (keyBox) keyBox.hidden = !((data && (data.truncated || data.keyError)) || hasUserKey);
    if (keyInput && document.activeElement !== keyInput) keyInput.value = NeisApi.getUserKey();
    if (clearBtn) clearBtn.hidden = !hasUserKey;
  }

  async function applyNeisKey(value) {
    NeisApi.setUserKey(value);
    clearClassTimetableCache();
    await loadSchoolClasses();
    await fetchHomeroomTimetable({ force: true });
    await renderClassSubjectPicker(activePickerClass, { force: true });
    await syncAcademicScheduleToCalendar(true, true);
    if (typeof showToast === 'function') {
      showToast(value ? 'NEIS 인증키를 이 기기에 저장하고 다시 조회했습니다.' : '저장된 NEIS 인증키를 삭제했습니다.');
    }
  }

  var pickerRequestNo = 0;

  function isLessonInTimetable(tt, clsShort, subName, slots) {
    return slots.some(function (sl) {
      var cell = tt[sl.day] && tt[sl.day][sl.period];
      return cell && cell.className === clsShort && cell.subject === subName;
    });
  }

  // 교과 드롭다운: 고른 학급의 NEIS 시간표에 있는 과목과 수업 요일·교시. 이미 추가한 과목은 고를 수 없게 표시합니다.
  function renderSubjectOptions(targetClass, data) {
    var select = el('tsSubjectSelect');
    if (!select) return;
    var tt = getTimetable();
    var pickedIndex = -1;
    select.innerHTML = '<option value="">교과 선택 (' + data.subjects.length + '과목)</option>' + data.subjects.map(function (item, idx) {
      var added = isLessonInTimetable(tt, targetClass.short, item.subject, item.slots);
      if (!added && pendingSubjects[0] === item.subject) pickedIndex = idx;
      return '<option value="' + idx + '"' + (added ? ' disabled' : '') + '>' +
        escapeHtml((added ? '✓ ' : '') + item.subject + ' — ' + item.slotSummary + (added ? ' (추가됨)' : '')) +
      '</option>';
    }).join('');
    select.disabled = false;
    select.value = pickedIndex >= 0 ? String(pickedIndex) : '';
    if (pickedIndex < 0) pendingSubjects = [];
    updateAddLessonsButton();
  }

  // 고른 학급의 월~금 1~7교시 시간표 표 — 칸을 누르면 그 교시만 내 시간표에 넣거나 뺍니다.
  function renderClassMatrix(targetClass, data) {
    var matrixBoxEl = el('tsClassMatrixPreview');
    if (!matrixBoxEl || !data.days) return;
    var tt = getTimetable();
    var mHtml = '<table class="ts-class-matrix-table"><thead><tr><th>교시</th><th>월</th><th>화</th><th>수</th><th>목</th><th>금</th></tr></thead><tbody>';
    for (var p = 1; p <= 7; p++) {
      mHtml += '<tr><th>' + p + '</th>';
      WEEK_DAYS.forEach(function (d) {
        var subName = (data.days[d] && data.days[d][p]) || '';
        var myCell = (tt[d] && tt[d][p]) || {};
        var isAssignedHere = myCell.className === targetClass.short && myCell.subject === subName && Boolean(subName);
        mHtml += '<td class="matrix-cell' + (isAssignedHere ? ' assigned' : '') + (subName ? '' : ' empty') + '" data-mday="' + d + '" data-mperiod="' + p + '" data-msub="' + escapeHtml(subName) + '">' +
          (isAssignedHere ? '✓ ' : '') + escapeHtml(subName || '-') +
        '</td>';
      });
      mHtml += '</tr>';
    }
    mHtml += '</tbody></table>';
    matrixBoxEl.innerHTML = mHtml;

    matrixBoxEl.querySelectorAll('td.matrix-cell').forEach(function (td) {
      td.addEventListener('click', function () {
        var d = this.getAttribute('data-mday');
        var p = Number(this.getAttribute('data-mperiod'));
        var sName = this.getAttribute('data-msub');
        if (!d || !p || !sName) return;
        toggleSingleSlotIntoTeacherTimetable(targetClass, d, p, sName);
      });
    });
  }

  // 고른 학급의 이번 주 NEIS 시간표에서 교과 드롭다운과 시간표 표를 채웁니다.
  async function renderClassSubjectPicker(classObj, options) {
    ensureDefaultSchool();
    var targetClass = classObj || activePickerClass;
    var labelEl = el('tsActiveClassLabel');
    var subjectSelect = el('tsSubjectSelect');
    var matrixTitleEl = el('tsClassMatrixTitle');
    var matrixBoxEl = el('tsClassMatrixPreview');
    var toggleRow = el('tsMatrixToggleRow');
    if (!subjectSelect) return;
    if (!targetClass) {
      // 아직 학급을 고르지 않은 상태
      currentPickerData = null;
      pendingSubjects = [];
      renderStepper();
      updateAddLessonsButton();
      return;
    }
    if (!activePickerClass || activePickerClass.short !== targetClass.short) pendingSubjects = [];
    activePickerClass = targetClass;
    if (activeGrade === 'etc') {
      // '기타' 화면에서는 학급·교과 영역이 접혀 있으므로 목록만 맞춰 둡니다.
      renderClassPills();
      return;
    }
    activeGrade = targetClass.grade;
    renderClassPills();

    if (labelEl) labelEl.textContent = targetClass.label;
    if (matrixTitleEl) {
      matrixTitleEl.textContent = '[' + targetClass.label + '] 주간 시간표 (칸을 누르면 그 교시만 내 시간표에 넣거나 뺍니다)';
    }

    // 여러 반을 빠르게 바꿨을 때 늦게 도착한 응답이 화면을 덮어쓰지 않도록 요청 번호로 가립니다.
    var ownRequest = ++pickerRequestNo;
    var forceReload = Boolean(options && options.force);
    if (forceReload || !readClassTimetableCache(targetClass)) {
      setSubjectPlaceholder('NEIS에서 ' + targetClass.label + ' 시간표를 불러오는 중...');
      if (matrixBoxEl) matrixBoxEl.innerHTML = '';
      updateAddLessonsButton();
    }

    var data = await fetchClassTimetableData(targetClass, { force: forceReload });
    if (ownRequest !== pickerRequestNo) return;
    currentPickerData = (data && !data.error) ? data : null;
    renderNeisStatus(data);
    if (!data || data.error || !data.subjects || !data.subjects.length) {
      pendingSubjects = [];
      setSubjectPlaceholder('표시할 교과가 없습니다');
      if (matrixBoxEl) matrixBoxEl.innerHTML = '';
      if (toggleRow) toggleRow.hidden = true;
      updateAddLessonsButton();
      return;
    }

    if (toggleRow) toggleRow.hidden = false;
    renderSubjectOptions(targetClass, data);
    renderClassMatrix(targetClass, data);
  }

  function toggleSingleSlotIntoTeacherTimetable(classObj, day, period, subName) {
    var tt = getTimetable();
    if (!tt[day]) tt[day] = {};
    var cur = tt[day][period] || {};
    var clsShort = classObj.short;
    if (cur.className === clsShort && cur.subject === subName) {
      tt[day][period] = { className: '', subject: '', room: '' };
      if (typeof showToast === 'function') {
        showToast(day + '요일 ' + period + '교시 [' + clsShort + ' ' + subName + '] 수업을 해제했습니다.');
      }
    } else {
      tt[day][period] = { className: clsShort, subject: subName, room: clsShort };
      if (!selectedSubjects.some(function (s) { return s.short === clsShort; })) {
        selectedSubjects.push(classObj);
        writeJson(SUBJECT_CLASSES_KEY, selectedSubjects);
        renderClassPills();
      }
      if (typeof showToast === 'function') {
        showToast(day + '요일 ' + period + '교시에 [' + clsShort + ' ' + subName + '] 수업이 입력되었습니다.');
      }
    }
    writeJson(TIMETABLE_KEY, tt);
    afterTimetableChanged();
  }

  // 2-2. 창체 동아리 불러오기 & 내 담당 동아리 선택 UI 렌더링
  function getAllClubs() {
    var custom = readJson(CUSTOM_CLUBS_KEY, []);
    if (!Array.isArray(custom)) custom = [];
    return custom.concat(BUILTIN_CLUBS);
  }

  function getActiveClubSlotOption() {
    return CLUB_SLOT_OPTIONS.find(function (o) { return o.id === activeClubSlotId; }) || CLUB_SLOT_OPTIONS[0];
  }

  function renderClubPicker() {
    var slotBox = el('tsClubSlotPills');
    var catBox = el('tsClubCatTabs');
    var pillsBox = el('tsClubPills');
    if (!slotBox || !catBox || !pillsBox) return;

    slotBox.innerHTML = CLUB_SLOT_OPTIONS.map(function (opt) {
      var isAct = opt.id === activeClubSlotId;
      return '<button type="button" class="ts-club-slot-btn' + (isAct ? ' active' : '') + '" data-club-slot="' + opt.id + '">' + escapeHtml(opt.label) + '</button>';
    }).join('');

    slotBox.querySelectorAll('[data-club-slot]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        activeClubSlotId = this.getAttribute('data-club-slot');
        writeJson(ACTIVE_CLUB_SLOT_KEY, activeClubSlotId);
        renderClubPicker();
      });
    });

    var categories = ['전체', 'NEIS 창체·기본', '과학·IT·수학', '인문·사회·학술', '예술·체육·미디어', '내 등록 동아리'];
    catBox.innerHTML = categories.map(function (cat) {
      var isAct = cat === activeClubCategory;
      return '<button type="button" class="ts-club-cat-btn' + (isAct ? ' active' : '') + '" data-club-cat="' + escapeHtml(cat) + '">' + escapeHtml(cat) + '</button>';
    }).join('');

    catBox.querySelectorAll('[data-club-cat]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        activeClubCategory = this.getAttribute('data-club-cat');
        renderClubPicker();
      });
    });

    var allClubs = getAllClubs();
    var filtered = allClubs.filter(function (c) {
      return activeClubCategory === '전체' || c.category === activeClubCategory;
    });

    var tt = getTimetable();
    var slotOpt = getActiveClubSlotOption();
    var slotLabelShort = slotOpt.slots.map(function (s) { return s.day + s.period; }).join('·');

    pillsBox.innerHTML = filtered.map(function (club, idx) {
      var fullSubName = club.name.indexOf('동아리') >= 0 || club.name.indexOf('창체') >= 0 ? club.name : ('동아리(' + club.name + ')');
      var assigned = slotOpt.slots.some(function (sl) {
        var cell = tt[sl.day] && tt[sl.day][sl.period];
        return cell && (cell.subject === fullSubName || cell.subject === club.name);
      });
      return '<button type="button" class="ts-subject-pill-btn club-pill' + (assigned ? ' selected' : '') + '" data-club-idx="' + idx + '">' +
        '<strong>' + (assigned ? '✓ ' : '＋ ') + escapeHtml(club.name) + '</strong>' +
        '<small>' + escapeHtml('동아리 · ' + slotLabelShort + (club.room ? ' · ' + club.room : '')) + '</small>' +
      '</button>';
    }).join('');

    pillsBox.querySelectorAll('[data-club-idx]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var idx = Number(this.getAttribute('data-club-idx'));
        var club = filtered[idx];
        if (!club) return;
        toggleClubIntoTimetable(club);
      });
    });
  }

  function toggleClubIntoTimetable(club) {
    var tt = getTimetable();
    var slotOpt = getActiveClubSlotOption();
    var fullSubName = club.name.indexOf('동아리') >= 0 || club.name.indexOf('창체') >= 0 ? club.name : ('동아리(' + club.name + ')');

    var alreadyAssigned = slotOpt.slots.some(function (sl) {
      var cell = tt[sl.day] && tt[sl.day][sl.period];
      return cell && (cell.subject === fullSubName || cell.subject === club.name);
    });

    if (alreadyAssigned) {
      slotOpt.slots.forEach(function (sl) {
        if (tt[sl.day] && tt[sl.day][sl.period] && (tt[sl.day][sl.period].subject === fullSubName || tt[sl.day][sl.period].subject === club.name)) {
          tt[sl.day][sl.period] = { className: '', subject: '', room: '' };
        }
      });
      if (typeof showToast === 'function') {
        showToast('[' + club.name + '] 동아리를 내 시간표에서 해제했습니다.');
      }
    } else {
      slotOpt.slots.forEach(function (sl) {
        if (!tt[sl.day]) tt[sl.day] = {};
        tt[sl.day][sl.period] = {
          className: '동아리',
          subject: fullSubName,
          room: club.room || '동아리실'
        };
      });
      if (typeof showToast === 'function') {
        showToast('[' + fullSubName + '] 활동이 ' + slotOpt.label + '에 자동 입력되었습니다.');
      }
    }

    writeJson(TIMETABLE_KEY, tt);
    renderWeeklyTimetable();
    renderClubPicker();
    if (activePickerClass) renderClassSubjectPicker(activePickerClass);
    if (typeof TodayBriefing !== 'undefined' && TodayBriefing.refreshAll) {
      TodayBriefing.refreshAll();
    }
  }

  function addCustomClubAndAssign() {
    var nameInput = el('tsCustomClubNameInput');
    var roomInput = el('tsCustomClubRoomInput');
    var name = (nameInput ? nameInput.value : '').trim();
    var room = (roomInput ? roomInput.value : '').trim() || '동아리실';
    if (!name) {
      if (typeof showToast === 'function') showToast('추가할 동아리 이름을 입력해 주세요.');
      if (nameInput) nameInput.focus();
      return;
    }
    var custom = readJson(CUSTOM_CLUBS_KEY, []);
    if (!Array.isArray(custom)) custom = [];
    var newClub = { name: name, category: '내 등록 동아리', room: room };
    custom = custom.filter(function (c) { return c.name !== name; });
    custom.unshift(newClub);
    writeJson(CUSTOM_CLUBS_KEY, custom);
    if (nameInput) nameInput.value = '';
    if (roomInput) roomInput.value = '';
    activeClubCategory = '전체';
    toggleClubIntoTimetable(newClub);
  }

  async function fetchHomeroomTimetable(options) {
    ensureDefaultSchool();
    var refClass = (!isNonHomeroom && selectedHomeroom) ? selectedHomeroom : (selectedSubjects[0] || availableClasses[0]);
    if (!selectedSchool || !refClass) return getHomeroomTimetable();
    var data = await fetchClassTimetableData(refClass, options);
    if (data && data.days) {
      writeJson(HOMEROOM_TT_KEY, data);
      return data;
    }
    return getHomeroomTimetable();
  }

  function getHomeroomTimetable() {
    var cached = readJson(HOMEROOM_TT_KEY, null);
    var refClass = (!isNonHomeroom && selectedHomeroom) ? selectedHomeroom : (selectedSubjects[0] || availableClasses[0]);
    if (!cached || !cached.days || !refClass || !selectedSchool) return null;
    // 담임 반이나 학교를 바꾼 직후 조회에 실패해도 이전 반 시간표가 '우리 반'으로 보이지 않게 합니다.
    if (cached.classShort !== refClass.short || cached.schoolName !== selectedSchool.schoolName) return null;
    return cached;
  }

  async function importHomeroomToTeacherTimetable() {
    if (isNonHomeroom || !selectedHomeroom) {
      if (typeof showToast === 'function') {
        showToast('현재 담임 없음 상태입니다. 오른쪽 [담임 학급]을 먼저 고르거나, 학년 → 학급 → 교과 순서로 담당 수업을 추가하세요.');
      }
      return;
    }
    var hTt = await fetchHomeroomTimetable();
    if (!hTt || !hTt.days) return;
    var tt = getTimetable();
    var clsShort = (selectedHomeroom && selectedHomeroom.short) || '2-2';
    ['월', '화', '수', '목', '금'].forEach(function (d) {
      if (!tt[d]) tt[d] = {};
      for (var p = 1; p <= 7; p++) {
        var sub = hTt.days[d] && hTt.days[d][p];
        if (sub) {
          tt[d][p] = { className: clsShort, subject: sub, room: clsShort };
        }
      }
    });
    writeJson(TIMETABLE_KEY, tt);
    // 담임 반이 오른쪽 '추가한 담당 수업' 목록에도 학급으로 잡히도록 합니다.
    if (!selectedSubjects.some(function (c) { return c.short === clsShort; })) {
      selectedSubjects.push(selectedHomeroom);
      writeJson(SUBJECT_CLASSES_KEY, selectedSubjects);
    }
    activePickerClass = selectedHomeroom;
    activeGrade = selectedHomeroom.grade;
    afterTimetableChanged();
    if (typeof showToast === 'function') {
      showToast(clsShort + ' 담임 학급의 전체 시간표를 내 주간 시간표로 복사했습니다.');
    }
  }

  // NEIS 학사일정(SchoolSchedule) 자동 조회 및 일정관리 달력(ScheduleCalendar) 동기화
  //  - 인증키가 있으면 이번 학년도(3월~다음 해 2월) 전체를 한 번에 받습니다.
  //  - 인증키가 없으면 요청당 5건만 오므로 이번 달·다음 달을 8일 단위로 나눠 받습니다(일부 누락 가능).
  //  - 일정 날짜는 NEIS 가 준 날짜(AA_YMD)를 그대로 씁니다.
  async function syncAcademicScheduleToCalendar(silent, force) {
    ensureDefaultSchool();
    if (!selectedSchool) return [];

    var now = new Date();
    var schoolId = selectedSchool.atptCode + ':' + selectedSchool.schoolCode;
    var meta = readJson(ACADEMIC_META_KEY, null);
    // 화면을 열 때마다 도는 자동 조회는 6시간에 한 번이면 충분합니다. (버튼을 누르면 항상 새로 조회)
    if (silent && !force && meta && meta.school === schoolId && Date.now() - Number(meta.syncedAt || 0) < ACADEMIC_SYNC_MS) {
      return readJson(ACADEMIC_CACHE_KEY, []);
    }

    var base = { ATPT_OFCDC_SC_CODE: selectedSchool.atptCode, SD_SCHUL_CODE: selectedSchool.schoolCode };
    var rows = [];
    var complete = false;
    var failed = false;
    try {
      var ay = academicYearOf(now);
      var yearResult = await fetchNeisRaw('SchoolSchedule', Object.assign({}, base, {
        AA_FROM_YMD: String(ay) + '0301',
        AA_TO_YMD: toYmd(new Date(ay + 1, 2, 0)),
        pSize: '1000'
      }));
      if (!yearResult.truncated) {
        rows = yearResult.rows;
        complete = true;
      } else {
        var chunks = [];
        [0, 1].forEach(function (offset) {
          var first = new Date(now.getFullYear(), now.getMonth() + offset, 1);
          var lastDay = new Date(first.getFullYear(), first.getMonth() + 1, 0).getDate();
          var ym = String(first.getFullYear()) + pad2(first.getMonth() + 1);
          [[1, 8], [9, 16], [17, 24], [25, lastDay]].forEach(function (range) {
            chunks.push([ym + pad2(range[0]), ym + pad2(range[1])]);
          });
        });
        var chunkResults = await Promise.all(chunks.map(function (range) {
          return fetchNeisRaw('SchoolSchedule', Object.assign({}, base, { AA_FROM_YMD: range[0], AA_TO_YMD: range[1] }));
        }));
        chunkResults.forEach(function (res) { rows = rows.concat(res.rows); });
      }
    } catch (e) {
      failed = true;
    }

    if (failed) {
      if (!silent && typeof showToast === 'function') {
        showToast('NEIS 학사일정을 불러오지 못했습니다. 네트워크 연결을 확인해 주세요.', 'error');
      }
      return readJson(ACADEMIC_CACHE_KEY, []);
    }

    // 같은 이름의 일정이 날짜가 이어지면 하나의 기간 일정으로 묶습니다. (예: 지필평가 4일)
    var dayEvents = [];
    var seenDayTitle = {};
    rows.forEach(function (r) {
      var ymd = String(r.AA_YMD || '').trim();
      var title = String(r.EVENT_NM || '').trim();
      if (ymd.length !== 8 || !title || title === '토요휴업일') return;
      if (seenDayTitle[ymd + '|' + title]) return;
      seenDayTitle[ymd + '|' + title] = true;
      dayEvents.push({ ymd: ymd, title: title, memo: String(r.SBTR_DD_SC_NM || 'NEIS 학사일정') });
    });
    dayEvents.sort(function (a, b) { return a.ymd < b.ymd ? -1 : (a.ymd > b.ymd ? 1 : 0); });

    function isoOf(ymd) { return ymd.slice(0, 4) + '-' + ymd.slice(4, 6) + '-' + ymd.slice(6, 8); }
    function nextYmd(ymd) { return toYmd(addDays(new Date(Number(ymd.slice(0, 4)), Number(ymd.slice(4, 6)) - 1, Number(ymd.slice(6, 8))), 1)); }
    var fetchedEvents = [];
    var openRuns = {};
    dayEvents.forEach(function (item) {
      var run = openRuns[item.title];
      if (run && nextYmd(run.lastYmd) === item.ymd) {
        run.lastYmd = item.ymd;
        run.event.end = isoOf(item.ymd);
        return;
      }
      var event = { start: isoOf(item.ymd), end: isoOf(item.ymd), title: item.title, memo: item.memo };
      openRuns[item.title] = { lastYmd: item.ymd, event: event };
      fetchedEvents.push(event);
    });

    writeJson(ACADEMIC_META_KEY, { school: schoolId, syncedAt: Date.now(), complete: complete });
    writeJson(ACADEMIC_CACHE_KEY, fetchedEvents);

    if (typeof ScheduleCalendar !== 'undefined' && ScheduleCalendar.getData && ScheduleCalendar.setData) {
      var props = ScheduleCalendar.getProperties ? ScheduleCalendar.getProperties() : [];
      if (props.indexOf('학사') < 0 && ScheduleCalendar.setProperties) {
        props.push('학사');
        ScheduleCalendar.setProperties(props);
      }
      var existing = ScheduleCalendar.getData();
      var existingMap = {};
      existing.forEach(function (ev) {
        existingMap[ev.start + '|' + ev.title] = true;
      });
      var addedCount = 0;
      fetchedEvents.forEach(function (ev) {
        var key = ev.start + '|' + ev.title;
        if (!existingMap[key]) {
          existingMap[key] = true;
          existing.push({
            id: 'neis-' + ev.start.replace(/-/g, '') + '-' + Math.random().toString(36).slice(2, 7),
            title: ev.title,
            start: ev.start,
            end: ev.end,
            property: '학사',
            target: '전체',
            memo: selectedSchool.schoolName + ' NEIS 학사일정 (' + ev.memo + ')',
            studentId: '',
            studentName: '',
            consultMethod: '',
            status: '',
            followUp: '',
            createdAt: Date.now(),
            updatedAt: Date.now()
          });
          addedCount++;
        }
      });
      if (addedCount > 0) {
        ScheduleCalendar.setData(existing);
      } else {
        ScheduleCalendar.render();
      }
      if (!silent && typeof showToast === 'function') {
        showToast(selectedSchool.schoolName + ' NEIS 학사일정 ' + (addedCount > 0 ? addedCount + '건을 달력에 새로 입력했습니다.' : '이 이미 최신 상태로 반영되어 있습니다.') +
          (complete ? '' : ' (NEIS 인증키가 없어 이번 달·다음 달 일부만 가져왔습니다)'));
      }
    }
    if (typeof TodayBriefing !== 'undefined' && TodayBriefing.refreshAll) {
      TodayBriefing.refreshAll();
    }
    return fetchedEvents;
  }

  function applyClassesToApp() {
    var targets = [];
    if (!isNonHomeroom && selectedHomeroom) {
      targets.push({ name: selectedHomeroom.label + ' (담임)', short: selectedHomeroom.short });
    }
    selectedSubjects.forEach(function (s) {
      if (isNonHomeroom || !selectedHomeroom || s.short !== selectedHomeroom.short) {
        targets.push({ name: s.label, short: s.short });
      }
    });

    if (!targets.length) {
      alert('담임 학급이나 담당 수업 학급을 최소 1개 이상 선택해 주세요.');
      return;
    }

    targets.slice(0, MAX_CLASSES).forEach(function (t, idx) {
      if (!allClassData[idx]) allClassData[idx] = {};
      allClassData[idx].className = t.name;
    });

    saveToLocalStorage();
    if (typeof updateClassSelect === 'function') updateClassSelect();
    if (typeof renderClassTabs === 'function') renderClassTabs();
    if (typeof TodayBriefing !== 'undefined' && TodayBriefing.refreshAll) TodayBriefing.refreshAll();
    if (typeof showToast === 'function') showToast(targets.length + '개 학급이 출석부 및 교무수첩에 반영되었습니다.');

    updateCellClassOptions();
  }

  function getTimetable() {
    var tt = readJson(TIMETABLE_KEY, null);
    if (!tt) {
      tt = defaultTimetable();
      writeJson(TIMETABLE_KEY, tt);
    }
    return tt;
  }

  function renderWeeklyTimetable() {
    var tbody = el('tsWeeklyTbody');
    if (!tbody) return;
    var tt = getTimetable();
    var days = ['월', '화', '수', '목', '금'];

    var html = '';
    for (var p = 1; p <= 7; p++) {
      html += '<tr>';
      html += '<th style="background:#f8fafc; font-weight:800; color:#334155;">' + p + '</th>';
      days.forEach(function (d) {
        var cell = (tt[d] && tt[d][p]) || { className: '', subject: '', room: '' };
        var hasContent = cell.subject || cell.className;
        var isClub = (cell.className === '동아리' || String(cell.subject || '').indexOf('동아리') >= 0);
        var badgeLabel = cell.className || (isClub ? '동아리' : '');
        var roomExtra = (cell.room && cell.room !== cell.className) ? (' · ' + cell.room) : '';
        var content = hasContent ?
          (badgeLabel ? '<span class="tt-cls-badge">' + escapeHtml(badgeLabel + roomExtra) + '</span>' : '') +
          '<strong style="display:block; font-size:0.8rem; color:#0f172a; word-break:keep-all;">' + escapeHtml(cell.subject || '수업') + '</strong>' :
          '<span style="color:#cbd5e1;">-</span>';
        html += '<td class="' + (hasContent ? 'filled' : '') + (isClub ? ' is-club-cell' : '') + '" data-day="' + d + '" data-period="' + p + '">' + content + '</td>';
      });
      html += '</tr>';
    }
    tbody.innerHTML = html;

    tbody.querySelectorAll('td[data-day]').forEach(function (td) {
      td.addEventListener('click', function () {
        openCellEdit(this.dataset.day, Number(this.dataset.period));
      });
    });
    renderAssignedPanel();
  }

  // 개별 교시 셀 편집 시 선택한 학급(또는 동아리)의 정확한 해당 교시 과목 자동 매칭 및 과목 칩 렌더링
  async function renderCellSubjectChips(classShort, autoFillPeriodMatch) {
    var chipsBox = el('tsCellSubjectChips');
    var matchNotice = el('tsCellPeriodMatchNotice');
    var subInput = el('tsCellSubjectInput');
    var roomInput = el('tsCellRoomInput');
    if (!chipsBox) return;

    // '동아리'를 선택한 경우 동아리 목록을 칩으로 표시
    if (classShort === '동아리') {
      var clubs = getAllClubs();
      if (matchNotice) {
        matchNotice.style.display = 'block';
        matchNotice.textContent = '★ 동아리 활동 선택 모드: 아래에서 담당 동아리를 클릭하면 과목명과 활동실이 정확히 입력됩니다.';
      }
      chipsBox.innerHTML = clubs.map(function (c) {
        var fullSub = c.name.indexOf('동아리') >= 0 || c.name.indexOf('창체') >= 0 ? c.name : ('동아리(' + c.name + ')');
        return '<button type="button" class="btn small" data-cell-sub="' + escapeHtml(fullSub) + '" data-cell-room="' + escapeHtml(c.room || '동아리실') + '" style="font-size:0.74rem; padding:3px 8px; background:#ecfdf5; color:#065f46; border:1px solid #6ee7b7; font-weight:700;">' + escapeHtml(c.name) + '</button>';
      }).join('');
      chipsBox.querySelectorAll('[data-cell-sub]').forEach(function (btn) {
        btn.addEventListener('click', function () {
          if (subInput) subInput.value = this.getAttribute('data-cell-sub');
          if (roomInput) roomInput.value = this.getAttribute('data-cell-room') || '동아리실';
        });
      });
      if (autoFillPeriodMatch && subInput && !subInput.value && clubs[0]) {
        subInput.value = '동아리(' + clubs[0].name + ')';
      }
      return;
    }

    var clsObj = availableClasses.find(function (c) { return c.short === classShort; }) || selectedHomeroom || selectedSubjects[0] || availableClasses[0];
    if (!clsObj) {
      chipsBox.innerHTML = '';
      if (matchNotice) matchNotice.style.display = 'none';
      return;
    }
    var data = await fetchClassTimetableData(clsObj);
    if (!data || !data.subjects) {
      chipsBox.innerHTML = '';
      if (matchNotice) matchNotice.style.display = 'none';
      return;
    }

    var editDay = currentEditCell ? currentEditCell.day : '';
    var editPeriod = currentEditCell ? currentEditCell.period : 0;
    var exactPeriodSubject = (editDay && editPeriod && data.days && data.days[editDay] && data.days[editDay][editPeriod]) ? data.days[editDay][editPeriod] : '';

    if (matchNotice) {
      if (exactPeriodSubject) {
        matchNotice.style.display = 'block';
        matchNotice.innerHTML = '★ <b>[' + escapeHtml(clsObj.short) + '] ' + escapeHtml(editDay) + '요일 ' + editPeriod + '교시</b> NEIS 실제 편성 과목: <b style="color:#047857; text-decoration:underline;">' + escapeHtml(exactPeriodSubject) + '</b> (다른 이동수업 과목이면 아래 칩을 클릭하세요)';
      } else {
        matchNotice.style.display = 'none';
      }
    }

    // 학급 드롭다운을 변경했을 때 해당 교시의 정확한 NEIS 과목명을 자동 입력
    if (autoFillPeriodMatch && exactPeriodSubject && subInput) {
      subInput.value = exactPeriodSubject;
    }

    // 해당 교시 편성 과목을 맨 앞에 배치하여 한눈에 보이도록 정렬
    var sortedSubjects = data.subjects.slice().sort(function (a, b) {
      var aExact = a.subject === exactPeriodSubject ? 0 : 1;
      var bExact = b.subject === exactPeriodSubject ? 0 : 1;
      return aExact - bExact;
    });

    chipsBox.innerHTML = sortedSubjects.map(function (item) {
      var isExact = item.subject === exactPeriodSubject;
      var style = isExact
        ? 'font-size:0.75rem; padding:4px 9px; background:#4f46e5; color:#ffffff; border:1px solid #4338ca; font-weight:800;'
        : 'font-size:0.74rem; padding:3px 8px; background:#eef2ff; color:#3730a3; border:1px solid #c7d2fe; font-weight:700;';
      return '<button type="button" class="btn small" data-cell-sub="' + escapeHtml(item.subject) + '" style="' + style + '">' +
        (isExact ? '★ ' : '') + escapeHtml(item.subject) + ' <small style="opacity:0.85;">(' + escapeHtml(item.slotSummary) + ')</small>' +
      '</button>';
    }).join('');

    chipsBox.querySelectorAll('[data-cell-sub]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var select = el('tsCellClassSelect');
        if (subInput) subInput.value = this.getAttribute('data-cell-sub');
        if (roomInput && (!roomInput.value || roomInput.value === '동아리실') && select && select.value) {
          roomInput.value = select.value;
        }
      });
    });
  }

  function openCellEdit(day, period) {
    currentEditCell = { day: day, period: period };
    var dialog = el('tsCellEditDialog');
    var title = el('tsCellEditTitle');
    var select = el('tsCellClassSelect');
    var subInput = el('tsCellSubjectInput');
    var roomInput = el('tsCellRoomInput');
    if (!dialog || !title || !select || !subInput || !roomInput) return;

    title.textContent = day + '요일 ' + period + '교시 수업·동아리 편집';
    updateCellClassOptions();

    var tt = getTimetable();
    var cur = (tt[day] && tt[day][period]) || { className: '', subject: '', room: '' };
    var defaultClassShort = cur.className || (activePickerClass ? activePickerClass.short : (selectedHomeroom ? selectedHomeroom.short : '2-2'));
    select.value = defaultClassShort;
    subInput.value = cur.subject || '';
    roomInput.value = cur.room || select.value || '';

    // 빈 칸을 클릭해 열었을 때는 선택한 학급의 해당 교시 과목이 바로 보이도록 자동 채움
    renderCellSubjectChips(select.value, !cur.subject);

    dialog.style.display = 'block';
    subInput.focus();
  }

  function updateCellClassOptions() {
    var select = el('tsCellClassSelect');
    if (!select) return;
    var opts = '<option value="">(학급 선택)</option>';
    opts += '<option value="동아리">★ 동아리 활동 (창체/자율)</option>';
    var seen = {};
    if (!isNonHomeroom && selectedHomeroom) {
      opts += '<option value="' + escapeHtml(selectedHomeroom.short) + '">' + escapeHtml(selectedHomeroom.short + ' (담임)') + '</option>';
      seen[selectedHomeroom.short] = true;
    }
    selectedSubjects.forEach(function (s) {
      if (!seen[s.short]) {
        opts += '<option value="' + escapeHtml(s.short) + '">' + escapeHtml(s.short + ' (수업반)') + '</option>';
        seen[s.short] = true;
      }
    });
    availableClasses.forEach(function (c) {
      if (!seen[c.short]) {
        opts += '<option value="' + escapeHtml(c.short) + '">' + escapeHtml(c.short) + '</option>';
        seen[c.short] = true;
      }
    });
    select.innerHTML = opts;
  }

  function saveCellEdit() {
    if (!currentEditCell) return;
    var day = currentEditCell.day;
    var period = currentEditCell.period;
    var select = el('tsCellClassSelect');
    var subInput = el('tsCellSubjectInput');
    var roomInput = el('tsCellRoomInput');

    var tt = getTimetable();
    if (!tt[day]) tt[day] = {};
    tt[day][period] = {
      className: (select ? select.value : '').trim(),
      subject: (subInput ? subInput.value : '').trim(),
      room: (roomInput ? roomInput.value : '').trim()
    };
    writeJson(TIMETABLE_KEY, tt);
    closeCellEdit();
    renderWeeklyTimetable();
    renderClassSubjectPicker(activePickerClass);
    renderClubPicker();
    if (typeof TodayBriefing !== 'undefined' && TodayBriefing.refreshAll) TodayBriefing.refreshAll();
    if (typeof showToast === 'function') showToast(day + '요일 ' + period + '교시 수업이 저장되었습니다.');
  }

  function deleteCellEdit() {
    if (!currentEditCell) return;
    var day = currentEditCell.day;
    var period = currentEditCell.period;
    var tt = getTimetable();
    if (tt[day]) {
      tt[day][period] = { className: '', subject: '', room: '' };
      writeJson(TIMETABLE_KEY, tt);
    }
    closeCellEdit();
    renderWeeklyTimetable();
    renderClassSubjectPicker(activePickerClass);
    renderClubPicker();
    if (typeof TodayBriefing !== 'undefined' && TodayBriefing.refreshAll) TodayBriefing.refreshAll();
  }

  function closeCellEdit() {
    currentEditCell = null;
    var dialog = el('tsCellEditDialog');
    if (dialog) dialog.style.display = 'none';
  }

  function open() {
    ensureDefaultSchool();
    renderSelectedSchool();
    if (!activePickerClass) activePickerClass = selectedHomeroom || selectedSubjects[0] || null;
    if (activeGrade == null) activeGrade = activePickerClass ? activePickerClass.grade : (gradesAvailable()[0] || null);
    renderClassPills();
    renderWeeklyTimetable();
    renderClassSubjectPicker(activePickerClass);
    renderClubPicker();
    closeCellEdit();
    openModal('teacherScheduleModal');

    // 이 학교의 학급 목록을 아직 NEIS에서 받아 온 적이 없으면(처음 열었을 때) 실제 학급 수로 맞춥니다.
    var statusText = el('tsClassStatusText');
    if (selectedSchool && readJson(CLASSES_SCHOOL_KEY, '') !== selectedSchool.atptCode + ':' + selectedSchool.schoolCode) {
      loadSchoolClasses().then(function () { renderClassSubjectPicker(activePickerClass); });
    } else if (statusText && selectedSchool) {
      statusText.textContent = selectedSchool.schoolName + ' · ' + availableClasses.length + '개 학급';
    }
  }

  function init() {
    ensureDefaultSchool();
    renderSelectedSchool();

    var searchBtn = el('tsSchoolSearchBtn');
    if (searchBtn) searchBtn.addEventListener('click', searchSchool);
    var searchInput = el('tsSchoolSearchInput');
    if (searchInput) searchInput.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); searchSchool(); }
    });

    var applyBtn = el('tsApplyClassesBtn');
    if (applyBtn) applyBtn.addEventListener('click', applyClassesToApp);

    var loadHomeroomBtn = el('tsLoadHomeroomTtBtn');
    if (loadHomeroomBtn) loadHomeroomBtn.addEventListener('click', importHomeroomToTeacherTimetable);

    // 담당 수업 추가: 학년 → 학급 → 교과 드롭다운
    var gradeSelect = el('tsGradeSelect');
    if (gradeSelect) gradeSelect.addEventListener('change', function () { selectGrade(this.value); });
    var classSelect = el('tsClassSelect');
    if (classSelect) classSelect.addEventListener('change', function () { selectClass(this.value); });
    var subjectSelect = el('tsSubjectSelect');
    if (subjectSelect) subjectSelect.addEventListener('change', function () {
      var item = (this.value !== '' && currentPickerData) ? currentPickerData.subjects[Number(this.value)] : null;
      pendingSubjects = item ? [item.subject] : [];
      updateAddLessonsButton();
    });
    var addLessonsBtn = el('tsAddLessonsBtn');
    if (addLessonsBtn) addLessonsBtn.addEventListener('click', addPickedLessons);

    // 오른쪽 패널: 학급으로 이동 · 학급 빼기 · 수업 삭제 · 담임 학급
    var selectedList = el('tsSelectedClassList');
    if (selectedList) selectedList.addEventListener('click', function (e) {
      var jump = e.target.closest('[data-jump-class]');
      if (jump) { selectClass(jump.getAttribute('data-jump-class')); return; }
      var removeClassBtn = e.target.closest('[data-remove-class]');
      if (removeClassBtn) { removeSelectedClass(removeClassBtn.getAttribute('data-remove-class')); return; }
      var removeLessonBtn = e.target.closest('[data-remove-lesson]');
      if (removeLessonBtn) removeLesson(removeLessonBtn.getAttribute('data-lesson-class'), removeLessonBtn.getAttribute('data-remove-lesson'));
    });
    var homeroomSelect = el('tsHomeroomSelect');
    if (homeroomSelect) homeroomSelect.addEventListener('change', function () { setHomeroom(this.value); });

    // 기타: 직접 입력으로 추가
    var etcAddBtn = el('tsEtcAddBtn');
    if (etcAddBtn) etcAddBtn.addEventListener('click', addEtcLesson);
    ['tsEtcClassInput', 'tsEtcSubjectInput'].forEach(function (id) {
      var input = el(id);
      if (input) input.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') { e.preventDefault(); addEtcLesson(); }
      });
    });

    var toggleMatrixBtn = el('tsToggleClassMatrixBtn');
    if (toggleMatrixBtn) {
      toggleMatrixBtn.addEventListener('click', function () {
        var wrap = el('tsClassMatrixPreviewWrap');
        if (!wrap) return;
        wrap.hidden = !wrap.hidden;
        this.textContent = wrap.hidden ? '학급 시간표 표 펼치기 ▾' : '학급 시간표 표 접기 ▴';
      });
    }

    var neisKeySaveBtn = el('tsNeisKeySaveBtn');
    var neisKeyInput = el('tsNeisKeyInput');
    if (neisKeySaveBtn && neisKeyInput) {
      neisKeySaveBtn.addEventListener('click', function () {
        var value = neisKeyInput.value.trim();
        if (!value) { neisKeyInput.focus(); return; }
        applyNeisKey(value);
      });
      neisKeyInput.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') { e.preventDefault(); neisKeySaveBtn.click(); }
      });
    }
    var neisKeyClearBtn = el('tsNeisKeyClearBtn');
    if (neisKeyClearBtn) neisKeyClearBtn.addEventListener('click', function () {
      if (neisKeyInput) neisKeyInput.value = '';
      applyNeisKey('');
    });

    var addClubBtn = el('tsAddCustomClubBtn');
    if (addClubBtn) addClubBtn.addEventListener('click', addCustomClubAndAssign);
    var customClubInput = el('tsCustomClubNameInput');
    if (customClubInput) {
      customClubInput.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') {
          e.preventDefault();
          addCustomClubAndAssign();
        }
      });
    }

    var clearScheduleBtn = el('tsClearScheduleBtn');
    if (clearScheduleBtn) clearScheduleBtn.addEventListener('click', function () {
      var days = ['월', '화', '수', '목', '금'];
      var emptyTt = {};
      days.forEach(function (d) {
        emptyTt[d] = {};
        for (var p = 1; p <= 7; p++) emptyTt[d][p] = { className: '', subject: '', room: '' };
      });
      writeJson(TIMETABLE_KEY, emptyTt);
      renderWeeklyTimetable();
      renderClassSubjectPicker(activePickerClass);
      renderClubPicker();
      if (typeof TodayBriefing !== 'undefined' && TodayBriefing.refreshAll) TodayBriefing.refreshAll();
      if (typeof showToast === 'function') showToast('주간 시간표를 모두 비웠습니다. 위에서 반별 과목이나 동아리를 클릭해 새로 구성하세요.');
    });

    var syncAcademicBtn = el('tsSyncAcademicCalBtn');
    if (syncAcademicBtn) syncAcademicBtn.addEventListener('click', function () {
      syncAcademicScheduleToCalendar(false);
    });

    var calNeisSyncBtn = el('calNeisSyncBtn');
    if (calNeisSyncBtn) calNeisSyncBtn.addEventListener('click', function () {
      syncAcademicScheduleToCalendar(false);
    });

    var calFullscreenToggleBtn = el('calFullscreenToggleBtn');
    if (calFullscreenToggleBtn) calFullscreenToggleBtn.addEventListener('click', function () {
      var content = document.querySelector('#calendarModal .calendar-content');
      if (!content) return;
      var isCompact = content.classList.toggle('calendar-compact');
      this.textContent = isCompact ? '전체화면으로 보기' : '기본 크기로 보기';
    });

    var saveBtn = el('tsSaveScheduleBtn');
    if (saveBtn) saveBtn.addEventListener('click', function () {
      if (typeof TodayBriefing !== 'undefined' && TodayBriefing.refreshAll) TodayBriefing.refreshAll();
      if (typeof showToast === 'function') showToast('주간 시간표가 교무수첩과 오늘 브리핑에 저장되었습니다.');
    });

    var cellSelect = el('tsCellClassSelect');
    if (cellSelect) cellSelect.addEventListener('change', async function () {
      var roomInput = el('tsCellRoomInput');
      if (roomInput) roomInput.value = this.value === '동아리' ? '동아리실' : this.value;
      // 학급 선택 변경 시 즉시 해당 학급의 그 교시 과목을 정확히 찾아 과목 입력란에 자동 반영!
      await renderCellSubjectChips(this.value, true);
      // 동시에 상단 과목 선택기도 해당 학급으로 동기화하여 과목이 정확히 보이도록 처리
      if (this.value && this.value !== '동아리') {
        var foundCls = availableClasses.find(function (c) { return c.short === this.value; });
        if (foundCls) {
          activePickerClass = foundCls;
          renderClassSubjectPicker(foundCls);
        }
      }
    });

    var cellSaveBtn = el('tsCellSaveBtn');
    if (cellSaveBtn) cellSaveBtn.addEventListener('click', saveCellEdit);

    var cellDelBtn = el('tsCellDeleteBtn');
    if (cellDelBtn) cellDelBtn.addEventListener('click', deleteCellEdit);

    var cellCloseBtn = el('tsCellEditCloseBtn');
    if (cellCloseBtn) cellCloseBtn.addEventListener('click', closeCellEdit);

    var openStuTtBtn = el('tsOpenStuTtBtn');
    if (openStuTtBtn) openStuTtBtn.addEventListener('click', function () {
      closeModal('teacherScheduleModal');
      StudentTimetable.openUploadModal();
    });

    // 상단 메뉴 버튼 연결
    var topDashboardBtn = el('topDashboardBtn');
    if (topDashboardBtn) topDashboardBtn.addEventListener('click', function () {
      if (typeof showHomeScreen === 'function') showHomeScreen();
    });

    var topScheduleBtn = el('topScheduleBtn');
    if (topScheduleBtn) topScheduleBtn.addEventListener('click', open);

    var rosterApiBtn = el('rosterNeisApiBtn');
    if (rosterApiBtn) rosterApiBtn.addEventListener('click', function () {
      closeModal('rosterModal');
      open();
    });

    var openNeisSchoolApiBtn = el('openNeisSchoolApiBtn');
    if (openNeisSchoolApiBtn) openNeisSchoolApiBtn.addEventListener('click', function () {
      closeModal('settingsModal');
      open();
    });

    // 통합 대시보드 버튼 연결
    var heroOpenScheduleBtn = el('heroOpenScheduleBtn');
    if (heroOpenScheduleBtn) heroOpenScheduleBtn.addEventListener('click', open);

    var heroOpenStuTtBtn = el('heroOpenStuTtBtn');
    if (heroOpenStuTtBtn) heroOpenStuTtBtn.addEventListener('click', function () {
      StudentTimetable.openUploadModal();
    });

    var heroOpenRosterBtn = el('heroOpenRosterBtn');
    if (heroOpenRosterBtn) heroOpenRosterBtn.addEventListener('click', function () {
      var headerRosterBtn = el('headerRosterBtn');
      if (headerRosterBtn) headerRosterBtn.click();
      else openModal('rosterModal');
    });

    var heroOpenSeatBtn = el('heroOpenSeatBtn');
    if (heroOpenSeatBtn) heroOpenSeatBtn.addEventListener('click', function () {
      var seatTab = document.querySelector('#modeTabs [data-mode="seat"]');
      if (seatTab) seatTab.click();
    });

    var heroOpenAttendBtn = el('heroOpenAttendBtn');
    if (heroOpenAttendBtn) heroOpenAttendBtn.addEventListener('click', function () {
      var attendTab = document.querySelector('#modeTabs [data-mode="attend"]');
      if (attendTab) attendTab.click();
    });

    var heroOpenCalendarBtn = el('heroOpenCalendarBtn');
    if (heroOpenCalendarBtn) heroOpenCalendarBtn.addEventListener('click', function () {
      if (typeof ScheduleCalendar !== 'undefined' && ScheduleCalendar.open) {
        ScheduleCalendar.open();
      } else {
        openModal('calendarModal');
      }
    });

    var toggleDrawerBtn = el('toggleSeatConfigDrawerBtn');
    if (toggleDrawerBtn) toggleDrawerBtn.addEventListener('click', function () {
      var drawer = el('seatConfigDrawer');
      if (!drawer) return;
      drawer.hidden = !drawer.hidden;
      this.textContent = drawer.hidden ? '교실 좌석 구조 설정 펼치기 ▾' : '교실 좌석 구조 설정 접기 ▴';
    });
  }

  function reloadFromStorage() {
    ensureDefaultSchool();
    renderSelectedSchool();
    renderClassPills();
    renderWeeklyTimetable();
    renderClassSubjectPicker(activePickerClass || selectedHomeroom || selectedSubjects[0]);
    renderClubPicker();
  }

  return {
    init: init,
    open: open,
    reloadFromStorage: reloadFromStorage,
    getTimetable: getTimetable,
    getSelectedSchool: function () { ensureDefaultSchool(); return selectedSchool; },
    getSelectedHomeroom: function () { ensureDefaultSchool(); return isNonHomeroom ? null : selectedHomeroom; },
    isNonHomeroom: function () { ensureDefaultSchool(); return isNonHomeroom; },
    getSelectedSubjects: function () { ensureDefaultSchool(); return selectedSubjects; },
    fetchHomeroomTimetable: fetchHomeroomTimetable,
    getHomeroomTimetable: getHomeroomTimetable,
    syncAcademicScheduleToCalendar: syncAcademicScheduleToCalendar,
    getCachedAcademicEvents: function () {
      ensureDefaultSchool();
      var meta = readJson(ACADEMIC_META_KEY, null);
      // 학교를 바꾼 뒤 아직 새로 조회하지 못했다면 이전 학교 일정을 보여 주지 않습니다.
      if (!meta || !selectedSchool || meta.school !== selectedSchool.atptCode + ':' + selectedSchool.schoolCode) return [];
      return readJson(ACADEMIC_CACHE_KEY, []);
    },
    fetchNeis: fetchNeis
  };
})();

// ===================================================================
// 자주 쓰는 구글 시트 & 즐겨찾는 사이트 메모장 모듈 (SheetMemoPad)
// ===================================================================
var SheetMemoPad = (function () {
  var STORAGE_KEY = 'tongSheetMemoPad_v1';
  var items = [];

  function el(id) { return document.getElementById(id); }

  function defaultItems() {
    return [
      {
        id: 'sm-default-1',
        category: '구글시트',
        title: '2026 교사용 올인원 마스터 통합 시트 (10개 탭)',
        url: 'https://docs.google.com/spreadsheets/d/12lGdMCyRGWpTlmG3vaXX18oeTcX2maJ_5Pfo9CJ4kfc/copy',
        memo: '0번 안내 탭 확인 · 6번 수업관찰 / 7번 AI세특 / 8번 출결·칭찬 실시간 연동 시트',
        updatedAt: Date.now()
      },
      {
        id: 'sm-default-2',
        category: '학급관리',
        title: '우리 반(2-2) 담임 학급 운영 & 상담 시트',
        url: 'https://docs.google.com/spreadsheets/',
        memo: '학부모 상담 일정, 1인 1역 배정표, 아침 조회 전달사항 기록용 시트',
        updatedAt: Date.now()
      },
      {
        id: 'sm-default-3',
        category: '수업관리',
        title: '교과 수행평가 채점 & 과목별 세특 점검표',
        url: 'https://docs.google.com/spreadsheets/',
        memo: '2-2 경제 수학Y · 2-8 기하V 수행평가 점수 및 미제출자 체크',
        updatedAt: Date.now()
      }
    ];
  }

  function load() {
    try {
      var raw = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
      if (Array.isArray(raw) && raw.length) {
        items = raw;
        return;
      }
    } catch (e) {}
    items = defaultItems();
    save();
  }

  function save() {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(items)); } catch (e) {}
    if (typeof CloudSheetSync !== "undefined" && CloudSheetSync.queueProfileSync) CloudSheetSync.queueProfileSync();
  }

  function normalizeUrl(u) {
    var s = String(u || '').trim();
    if (!s) return 'https://docs.google.com/spreadsheets/';
    if (!/^https?:\/\//i.test(s)) s = 'https://' + s;
    return s;
  }

  function render() {
    var dashList = el('dashSheetMemoList');
    var modalList = el('smModalList');

    var cardHtml = items.map(function (item) {
      var cat = item.category || '구글시트';
      return '<div class="dash-sm-item" data-sm-id="' + escapeHtml(item.id) + '">' +
        '<div>' +
          '<div class="dash-sm-top">' +
            '<span class="dash-sm-badge cat-' + escapeHtml(cat) + '">' + escapeHtml(cat) + '</span>' +
            '<div class="dash-sm-actions">' +
              '<button type="button" data-sm-edit="' + escapeHtml(item.id) + '">메모 수정</button>' +
              '<a href="' + escapeHtml(item.url) + '" target="_blank" rel="noopener noreferrer" class="sm-open-btn">열기 ↗</a>' +
            '</div>' +
          '</div>' +
          '<a href="' + escapeHtml(item.url) + '" target="_blank" rel="noopener noreferrer" class="dash-sm-title" style="margin-top:5px;">' + escapeHtml(item.title) + '</a>' +
        '</div>' +
        '<div class="dash-sm-memo">' + escapeHtml(item.memo || '메모 없음 (메모 수정을 눌러 기록하세요)') + '</div>' +
      '</div>';
    }).join('');

    if (dashList) {
      dashList.innerHTML = cardHtml;
      dashList.querySelectorAll('[data-sm-edit]').forEach(function (btn) {
        btn.addEventListener('click', function () {
          openModalForEdit(this.getAttribute('data-sm-edit'));
        });
      });
    }

    if (modalList) {
      modalList.innerHTML = items.map(function (item) {
        var cat = item.category || '구글시트';
        return '<div class="dash-sm-item" style="flex-direction:row; align-items:center; justify-content:space-between; gap:12px;">' +
          '<div style="flex:1; min-width:0;">' +
            '<div style="display:flex; align-items:center; gap:6px; margin-bottom:3px;">' +
              '<span class="dash-sm-badge cat-' + escapeHtml(cat) + '">' + escapeHtml(cat) + '</span>' +
              '<a href="' + escapeHtml(item.url) + '" target="_blank" rel="noopener noreferrer" class="dash-sm-title">' + escapeHtml(item.title) + '</a>' +
            '</div>' +
            '<div style="font-size:0.76rem; color:#475569;">메모: <b>' + escapeHtml(item.memo || '-') + '</b></div>' +
          '</div>' +
          '<div class="dash-sm-actions" style="flex-shrink:0;">' +
            '<a href="' + escapeHtml(item.url) + '" target="_blank" rel="noopener noreferrer" class="sm-open-btn">시트·사이트 열기 ↗</a>' +
            '<button type="button" data-sm-modal-edit="' + escapeHtml(item.id) + '">수정</button>' +
            '<button type="button" data-sm-modal-del="' + escapeHtml(item.id) + '" style="color:#b91c1c;">삭제</button>' +
          '</div>' +
        '</div>';
      }).join('');

      modalList.querySelectorAll('[data-sm-modal-edit]').forEach(function (btn) {
        btn.addEventListener('click', function () {
          populateForm(this.getAttribute('data-sm-modal-edit'));
        });
      });
      modalList.querySelectorAll('[data-sm-modal-del]').forEach(function (btn) {
        btn.addEventListener('click', function () {
          var id = this.getAttribute('data-sm-modal-del');
          items = items.filter(function (x) { return x.id !== id; });
          save();
          render();
          if (typeof showToast === 'function') showToast('항목을 삭제했습니다.');
        });
      });
    }
  }

  function populateForm(id) {
    var found = items.find(function (x) { return x.id === id; });
    if (!found) return;
    if (el('smEditId')) el('smEditId').value = found.id;
    if (el('smCategoryInput')) el('smCategoryInput').value = found.category || '구글시트';
    if (el('smTitleInput')) el('smTitleInput').value = found.title || '';
    if (el('smUrlInput')) el('smUrlInput').value = found.url || '';
    if (el('smMemoInput')) el('smMemoInput').value = found.memo || '';
    if (el('smFormTitle')) el('smFormTitle').textContent = '시트 / 사이트 메모 수정';
    if (el('smCancelEditBtn')) el('smCancelEditBtn').hidden = false;
    if (el('smMemoInput')) el('smMemoInput').focus();
  }

  function resetForm() {
    if (el('smEditId')) el('smEditId').value = '';
    if (el('smCategoryInput')) el('smCategoryInput').value = '구글시트';
    if (el('smTitleInput')) el('smTitleInput').value = '';
    if (el('smUrlInput')) el('smUrlInput').value = '';
    if (el('smMemoInput')) el('smMemoInput').value = '';
    if (el('smFormTitle')) el('smFormTitle').textContent = '새 구글 시트 / 즐겨찾는 사이트 등록';
    if (el('smCancelEditBtn')) el('smCancelEditBtn').hidden = true;
  }

  function openModalForEdit(id) {
    openModal('sheetMemoModal');
    if (id) populateForm(id);
    else resetForm();
  }

  function init() {
    load();
    render();

    var form = el('sheetMemoForm');
    if (form) {
      form.addEventListener('submit', function (e) {
        e.preventDefault();
        var editId = (el('smEditId') ? el('smEditId').value : '').trim();
        var cat = el('smCategoryInput') ? el('smCategoryInput').value : '구글시트';
        var title = (el('smTitleInput') ? el('smTitleInput').value : '').trim();
        var url = normalizeUrl(el('smUrlInput') ? el('smUrlInput').value : '');
        var memo = (el('smMemoInput') ? el('smMemoInput').value : '').trim();
        if (!title) return;

        if (editId) {
          var idx = items.findIndex(function (x) { return x.id === editId; });
          if (idx >= 0) {
            items[idx] = { id: editId, category: cat, title: title, url: url, memo: memo, updatedAt: Date.now() };
          }
        } else {
          items.unshift({
            id: 'sm-' + Date.now() + '-' + Math.random().toString(36).slice(2, 6),
            category: cat,
            title: title,
            url: url,
            memo: memo,
            updatedAt: Date.now()
          });
        }
        save();
        render();
        resetForm();
        if (typeof showToast === 'function') showToast('구글 시트·사이트 메모가 저장되었습니다.');
      });
    }

    var cancelBtn = el('smCancelEditBtn');
    if (cancelBtn) cancelBtn.addEventListener('click', resetForm);

    var dashAddBtn = el('dashAddSheetMemoBtn');
    if (dashAddBtn) dashAddBtn.addEventListener('click', function () { openModalForEdit(''); });

    var dashOpenBtn = el('dashOpenSheetMemoModalBtn');
    if (dashOpenBtn) dashOpenBtn.addEventListener('click', function () { openModalForEdit(''); });

    var topBtn = el('topSheetMemoBtn');
    if (topBtn) topBtn.addEventListener('click', function () { openModalForEdit(''); });

    var menuBtn = el('menuSheetMemoBtn');
    if (menuBtn) menuBtn.addEventListener('click', function () { openModalForEdit(''); });
  }

  return {
    init: init,
    open: function () { openModalForEdit(''); },
    render: render,
    reload: function () { load(); render(); }
  };
})();

// ===================================================================
// 교사 캐릭터 아이콘 이미지 붙여넣기(Ctrl+V) 및 업로드 모듈 (AvatarManager)
// ===================================================================
var AvatarManager = (function () {
  // 캐릭터 아이콘은 [오늘] 사이드바 머리와 우측 하단 플로팅 버튼에서 함께 씁니다.
  var AVATAR_KEY = 'tongCustomAvatar_v1';
  var DEFAULT_SRC = '/assets/teacher-avatar-default.png';

  function el(id) { return document.getElementById(id); }

  function readCustom() {
    try {
      var saved = localStorage.getItem(AVATAR_KEY);
      return saved && saved.indexOf('data:image/') === 0 ? saved : '';
    } catch (e) { return ''; }
  }

  function render() {
    var custom = readCustom();
    var src = custom || DEFAULT_SRC;
    var img = el('todayAvatarImg');
    if (img) img.src = src;
    var floatImg = el('floatingTodayAvatarImg');
    if (floatImg) floatImg.src = src;
    var resetBtn = el('todayAvatarResetBtn');
    if (resetBtn) resetBtn.hidden = !custom;
  }

  function saveCustomAvatar(dataUrl) {
    try { localStorage.setItem(AVATAR_KEY, dataUrl); } catch (e) {}
    render();
    if (typeof showToast === 'function') showToast('우측 하단 및 사이드바 아이콘을 내 사진으로 변경했습니다.');
  }

  function resetAvatar() {
    try { localStorage.removeItem(AVATAR_KEY); } catch (e) {}
    render();
    if (typeof showToast === 'function') showToast('기본 캐릭터 아이콘으로 되돌렸습니다.');
  }

  function handleFile(file) {
    if (!file || !file.type || file.type.indexOf('image/') !== 0) return;
    var reader = new FileReader();
    reader.onload = function (e) {
      if (e.target && e.target.result) saveCustomAvatar(e.target.result);
    };
    reader.readAsDataURL(file);
  }

  function init() {
    render();

    var fileInput = el('avatarUploadInput');
    ['todayAvatarBtn', 'todayAvatarChangeBtn'].forEach(function (id) {
      var btn = el(id);
      if (btn && fileInput) btn.addEventListener('click', function () { fileInput.click(); });
    });
    if (fileInput) {
      fileInput.addEventListener('change', function () {
        if (this.files && this.files[0]) handleFile(this.files[0]);
        this.value = '';
      });
    }
    var resetBtn = el('todayAvatarResetBtn');
    if (resetBtn) resetBtn.addEventListener('click', resetAvatar);

    // 이미지 붙여넣기(Ctrl+V)는 사이드바가 열려 있을 때만 아이콘으로 받습니다.
    document.addEventListener('paste', function (e) {
      var sidebar = el('todaySidebar');
      if (!sidebar || !sidebar.classList.contains('show')) return;
      var active = document.activeElement;
      if (active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.isContentEditable)) return;
      var items = (e.clipboardData && e.clipboardData.items) || [];
      for (var i = 0; i < items.length; i++) {
        if (items[i].type && items[i].type.indexOf('image/') === 0) {
          var file = items[i].getAsFile();
          if (file) {
            e.preventDefault();
            handleFile(file);
            break;
          }
        }
      }
    });
  }

  return { init: init };
})();

// ===================================================================
// 칠판형 오늘 브리핑 위젯 (시작 화면 & 우측 하단 [오늘] 팝업 통합)
// ===================================================================
var TodayBriefing = (function () {
  var FILTER_KEY = 'tongBriefingFilters_v1';
  var ORDER_KEY = 'tongBriefingOrder_v1';
  // [오늘] 사이드바와 홈 칠판에 놓을 수 있는 위젯 (기본 순서)
  var WIDGETS = [
    { key: 'teacherTt', label: '오늘 내 수업', hint: '내 시간표에 넣은 오늘 수업' },
    { key: 'homeroomTt', label: '우리 반 시간표', hint: '담임(담당) 학급의 오늘 시간표' },
    { key: 'meal', label: '급식', hint: '오늘 중식 메뉴' },
    { key: 'schedule', label: '학사일정', hint: '다가오는 학교 일정' },
    { key: 'weather', label: '날씨', hint: '학교 지역의 현재 날씨' }
  ];
  var cachedMealText = '';
  var cachedMealKey = '';  // 학교:날짜 — 같은 날 급식을 반복 조회하지 않기 위한 표시
  var cachedWeather = { loc: '날씨', text: '날씨를 불러오는 중...' };

  function el(id) { return document.getElementById(id); }

  function getFilters() {
    try {
      var f = JSON.parse(localStorage.getItem(FILTER_KEY) || 'null');
      if (f && typeof f === 'object') return f;
    } catch (e) {}
    return { teacherTt: true, homeroomTt: true, meal: true, schedule: true, weather: true };
  }

  function saveFilters(f) {
    try { localStorage.setItem(FILTER_KEY, JSON.stringify(f)); } catch (e) {}
  }

  // 저장된 위젯 순서 (모르는 키는 버리고, 새로 생긴 위젯은 뒤에 붙입니다)
  function getOrder() {
    var known = WIDGETS.map(function (w) { return w.key; });
    var saved = [];
    try { saved = JSON.parse(localStorage.getItem(ORDER_KEY) || '[]'); } catch (e) { saved = []; }
    if (!Array.isArray(saved)) saved = [];
    var order = saved.filter(function (key, idx) { return known.indexOf(key) >= 0 && saved.indexOf(key) === idx; });
    known.forEach(function (key) { if (order.indexOf(key) < 0) order.push(key); });
    return order;
  }

  function saveOrder(order) {
    try { localStorage.setItem(ORDER_KEY, JSON.stringify(order)); } catch (e) {}
  }

  function getTodayInfo() {
    var now = new Date();
    var month = now.getMonth() + 1;
    var date = now.getDate();
    var dayNames = ['일', '월', '화', '수', '목', '금', '토'];
    var dayKor = dayNames[now.getDay()];
    var effectiveDay = (dayKor === '토' || dayKor === '일') ? '수' : dayKor;
    return {
      dateLabel: month + '월 ' + date + '일 ' + dayKor + '요일 · 오늘 브리핑',
      effectiveDay: effectiveDay
    };
  }

  function renderTeacherLessons(prefix, effectiveDay) {
    var tt = TeacherSchedule.getTimetable();
    var dayMap = (tt && tt[effectiveDay]) || {};
    var lessons = [];
    var endTimes = { 1: '09:50', 2: '10:50', 3: '11:50', 4: '12:50', 5: '14:30', 6: '15:30', 7: '16:30' };
    var lastPeriod = 0;

    for (var p = 1; p <= 7; p++) {
      var c = dayMap[p];
      if (c && (c.subject || c.className)) {
        lessons.push({ period: p, subject: c.subject || '수업', className: c.className || c.room || '' });
        lastPeriod = p;
      }
    }

    var countEl = el(prefix + 'LessonCount');
    var endEl = el(prefix + 'LessonEnd');
    var gridEl = el(prefix + 'LessonGrid');
    if (countEl) countEl.textContent = String(lessons.length);
    if (endEl) endEl.textContent = lastPeriod ? (endTimes[lastPeriod] + ' 끝') : '수업 없음';
    if (!gridEl) return;

    if (!lessons.length) {
      gridEl.innerHTML = '<div class="cb-lesson-card" style="grid-column:1 / -1;" data-cb-action="openSchedule">' +
        '<div class="cb-lesson-sub">오늘 등록된 수업이 없습니다</div>' +
        '<div class="cb-lesson-cls">클릭하여 반별 과목 시간표를 선택하세요</div>' +
      '</div>';
      return;
    }

    gridEl.innerHTML = lessons.map(function (item) {
      return '<div class="cb-lesson-card" data-cb-action="openSchedule">' +
        '<div class="cb-lesson-period">' + item.period + '교시</div>' +
        '<div class="cb-lesson-sub">' + escapeHtml(item.subject) + '</div>' +
        '<div class="cb-lesson-cls">' + escapeHtml(item.className) + '</div>' +
      '</div>';
    }).join('');
  }

  function renderHomeroomTimetable(prefix, effectiveDay) {
    var school = TeacherSchedule.getSelectedSchool();
    var homeroom = TeacherSchedule.getSelectedHomeroom();
    var nonHomeroom = TeacherSchedule.isNonHomeroom ? TeacherSchedule.isNonHomeroom() : !homeroom;
    var subClasses = TeacherSchedule.getSelectedSubjects ? TeacherSchedule.getSelectedSubjects() : [];
    var titleEl = el(prefix + 'HomeroomTitle');
    var badgeEl = el(prefix + 'HomeroomBadge');
    var schoolBtn = el(prefix + 'SchoolBtn');
    var listEl = el(prefix + 'HomeroomList');

    if (schoolBtn && school) {
      var suffix = nonHomeroom ? ' · 비담임' : (homeroom ? ' · ' + homeroom.short : '');
      schoolBtn.textContent = school.schoolName + suffix + ' 설정 >';
    }
    if (titleEl) {
      titleEl.textContent = nonHomeroom ? '담당 학급 시간표' : '우리 반 시간표';
    }
    if (badgeEl) {
      if (nonHomeroom) {
        var refShort = (subClasses[0] && subClasses[0].short) || '교과전담';
        badgeEl.textContent = '비담임 · ' + refShort + ' 기준';
      } else {
        badgeEl.textContent = homeroom ? (homeroom.short + ' 담임') : '미설정';
      }
    }
    if (!listEl) return;

    var hTt = TeacherSchedule.getHomeroomTimetable();
    var daySubjects = (hTt && hTt.days && hTt.days[effectiveDay]) || {};
    var chips = [];
    for (var p = 1; p <= 7; p++) {
      if (daySubjects[p]) {
        chips.push('<span class="cb-period-chip"><b>' + p + '교시</b> ' + escapeHtml(daySubjects[p]) + '</span>');
      }
    }
    if (!chips.length) {
      listEl.innerHTML = '<span style="font-size:0.78rem; color:#cbd5e1;">담임 학급을 선택하면 우리 반 1~7교시 시간표가 자동 표시됩니다.</span>';
    } else {
      listEl.innerHTML = chips.join('');
    }
  }

  // 오늘 급식만 조회합니다. 같은 학교·날짜는 한 번만 불러옵니다.
  async function loadMealInfo() {
    var school = TeacherSchedule.getSelectedSchool();
    if (!school) return;
    var now = new Date();
    var ymdToday = String(now.getFullYear()) + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0');
    var mealKey = school.atptCode + ':' + school.schoolCode + ':' + ymdToday;
    if (cachedMealKey === mealKey) return;
    cachedMealKey = mealKey;

    try {
      var rows = await TeacherSchedule.fetchNeis('mealServiceDietInfo', {
        ATPT_OFCDC_SC_CODE: school.atptCode,
        SD_SCHUL_CODE: school.schoolCode,
        MLSV_YMD: ymdToday
      });
      if (cachedMealKey !== mealKey) return;
      var lunch = rows.find(function (r) { return String(r.MMEAL_SC_NM || '').indexOf('중식') >= 0; }) || rows[0];
      var rawMenu = lunch ? String(lunch.DDISH_NM || '')
        .replace(/<br\s*\/?>/gi, ' · ')
        .replace(/\([0-9.]+\)/g, '')
        .replace(/\s+/g, ' ')
        .replace(/( · )+/g, ' · ')
        .replace(/^ · | · $/g, '')
        .trim() : '';
      var cal = lunch ? String(lunch.CAL_INFO || '').trim() : '';
      cachedMealText = rawMenu
        ? escapeHtml(rawMenu) + (cal ? ' <small>' + escapeHtml(cal) + '</small>' : '')
        : '오늘은 NEIS에 등록된 급식이 없습니다.';
    } catch (e) {
      if (cachedMealKey !== mealKey) return;
      cachedMealKey = '';  // 다음 새로고침 때 다시 시도
      cachedMealText = '급식 정보를 불러오지 못했습니다.';
    }
    updateMealDom();
  }

  function updateMealDom() {
    ['cbHomeMealText', 'cbModalMealText'].forEach(function (id) {
      var node = el(id);
      if (node && cachedMealText) node.innerHTML = cachedMealText;
    });
  }

  function renderAcademicSchedule(prefix) {
    var node = el(prefix + 'ScheduleText');
    if (!node) return;
    var events = TeacherSchedule.getCachedAcademicEvents() || [];
    var now = new Date();
    var todayIso = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
    // 오늘 진행 중이거나 앞으로 다가올 일정을 가까운 순서로 4개까지 보여 줍니다.
    var upcoming = events.filter(function (ev) { return (ev.end || ev.start) >= todayIso; })
      .sort(function (a, b) { return a.start < b.start ? -1 : (a.start > b.start ? 1 : 0); });
    if (!upcoming.length) {
      node.innerHTML = '다가오는 학사일정이 없습니다. <small>(학사일정 가져오기를 누르면 NEIS에서 다시 조회합니다)</small>';
      return;
    }
    var labels = [];
    upcoming.forEach(function (ev) {
      var label = ev.start.slice(5).replace('-', '/') + (ev.end && ev.end !== ev.start ? '~' + ev.end.slice(5).replace('-', '/') : '') + ' ' + ev.title;
      if (labels.indexOf(label) < 0 && labels.length < 4) labels.push(label);
    });
    node.innerHTML = escapeHtml(labels.join(' · '));
  }

  async function loadWeatherInfo() {
    var school = TeacherSchedule.getSelectedSchool();
    var addr = (school && (school.locationName || school.schoolName)) || '시흥';
    var city = '시흥';
    var lat = 37.380, lon = 126.803;
    if (addr.indexOf('수원') >= 0) { city = '수원'; lat = 37.2636; lon = 127.0286; }
    else if (addr.indexOf('서울') >= 0) { city = '서울'; lat = 37.5665; lon = 126.9780; }
    else if (addr.indexOf('성남') >= 0 || addr.indexOf('분당') >= 0) { city = '분당'; lat = 37.3827; lon = 127.1189; }
    else if (addr.indexOf('부산') >= 0) { city = '부산'; lat = 35.1796; lon = 129.0756; }
    else if (addr.indexOf('대구') >= 0 || addr.indexOf('경북') >= 0) { city = '대구'; lat = 35.8714; lon = 128.6014; }
    else if (addr.indexOf('대전') >= 0) { city = '대전'; lat = 36.3504; lon = 127.3845; }

    try {
      var url = 'https://api.open-meteo.com/v1/forecast?latitude=' + lat + '&longitude=' + lon + '&current=temperature_2m,weather_code&daily=temperature_2m_max,temperature_2m_min&timezone=Asia%2FSeoul&forecast_days=1';
      var res = await fetch(url);
      if (res.ok) {
        var data = await res.json();
        var curTemp = Math.round(data.current.temperature_2m);
        var minTemp = Math.round(data.daily.temperature_2m_min[0]);
        var maxTemp = Math.round(data.daily.temperature_2m_max[0]);
        var code = Number(data.current.weather_code || 0);
        var cond = code === 0 ? '맑음' : (code <= 3 ? '흐림' : (code <= 67 ? '비' : '눈'));
        cachedWeather = {
          loc: city,
          text: '지금 ' + cond + ' ' + curTemp + '° · 최저/최고 ' + minTemp + '°/' + maxTemp + '°'
        };
      } else {
        cachedWeather = { loc: city, text: '날씨 정보를 불러오지 못했습니다.' };
      }
    } catch (e) {
      cachedWeather = { loc: city, text: '날씨 정보를 불러오지 못했습니다.' };
    }
    ['cbHome', 'cbModal'].forEach(function (prefix) {
      var locNode = el(prefix + 'WeatherLoc');
      var txtNode = el(prefix + 'WeatherText');
      if (locNode) locNode.textContent = cachedWeather.loc;
      if (txtNode) txtNode.textContent = cachedWeather.text;
    });
  }

  // 위젯 보이기/숨기기와 순서를 홈 칠판·사이드바 칠판 양쪽에 똑같이 적용합니다.
  function applyWidgetLayout() {
    var filters = getFilters();
    var order = getOrder();
    document.querySelectorAll('.cb-board-inner').forEach(function (board) {
      var anchor = board.querySelector('.cb-bottom-bar, .tsb-empty');  // 위젯은 이 요소 바로 앞에 차례로 놓입니다
      order.forEach(function (key) {
        var sec = board.querySelector('[data-cb-sec="' + key + '"]');
        if (!sec) return;
        sec.hidden = filters[key] === false;
        board.insertBefore(sec, anchor);
      });
    });
    var empty = el('todayWidgetEmpty');
    if (empty) empty.hidden = order.some(function (key) { return filters[key] !== false; });
  }

  function refreshAll() {
    var info = getTodayInfo();
    ['cbHome', 'cbModal'].forEach(function (prefix) {
      var dNode = el(prefix + 'DateText');
      if (dNode) dNode.textContent = info.dateLabel;
      renderTeacherLessons(prefix, info.effectiveDay);
      renderHomeroomTimetable(prefix, info.effectiveDay);
      renderAcademicSchedule(prefix);
    });
    updateMealDom();
    applyWidgetLayout();
    loadMealInfo();
    loadWeatherInfo();
  }

  // ---- [오늘] 사이드바: 위젯 구성 목록 ---------------------------------
  function renderWidgetConfig() {
    var list = el('todayWidgetList');
    if (!list) return;
    var filters = getFilters();
    var order = getOrder();
    list.innerHTML = order.map(function (key, idx) {
      var widget = WIDGETS.find(function (w) { return w.key === key; });
      return '<li class="tsb-widget-item" data-widget="' + key + '">' +
        '<label>' +
          '<input type="checkbox" data-widget-toggle="' + key + '"' + (filters[key] !== false ? ' checked' : '') + '>' +
          '<span><b>' + escapeHtml(widget.label) + '</b><small>' + escapeHtml(widget.hint) + '</small></span>' +
        '</label>' +
        '<span class="tsb-widget-move">' +
          '<button type="button" data-widget-move="up" aria-label="' + escapeHtml(widget.label) + ' 위로 올리기"' + (idx === 0 ? ' disabled' : '') + '>▲</button>' +
          '<button type="button" data-widget-move="down" aria-label="' + escapeHtml(widget.label) + ' 아래로 내리기"' + (idx === order.length - 1 ? ' disabled' : '') + '>▼</button>' +
        '</span>' +
      '</li>';
    }).join('');
  }

  function moveWidget(key, direction) {
    var order = getOrder();
    var from = order.indexOf(key);
    var to = from + (direction === 'up' ? -1 : 1);
    if (from < 0 || to < 0 || to >= order.length) return;
    order.splice(to, 0, order.splice(from, 1)[0]);
    saveOrder(order);
    applyWidgetLayout();
    renderWidgetConfig();
    // 목록을 다시 그려도 방금 누른 위젯에서 계속 옮길 수 있게 초점을 돌려놓습니다.
    var item = el('todayWidgetList').querySelector('[data-widget="' + key + '"]');
    var next = item && (item.querySelector('[data-widget-move="' + direction + '"]:not([disabled])') || item.querySelector('[data-widget-move]:not([disabled])'));
    if (next) next.focus();
  }

  function setConfigOpen(open) {
    var panel = el('todayWidgetConfig');
    var btn = el('todayWidgetConfigBtn');
    if (!panel) return;
    panel.hidden = !open;
    if (btn) {
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
      btn.classList.toggle('active', open);
      btn.textContent = open ? '구성 닫기' : '위젯 구성';
    }
    if (open) renderWidgetConfig();
  }

  // ---- [오늘] 사이드바 열기 · 닫기 --------------------------------------
  function isSidebarOpen() {
    var sidebar = el('todaySidebar');
    return Boolean(sidebar && sidebar.classList.contains('show'));
  }

  function openSidebar(withConfig) {
    var sidebar = el('todaySidebar');
    if (!sidebar) return;
    refreshAll();
    setConfigOpen(Boolean(withConfig));
    sidebar.classList.add('show');
    sidebar.setAttribute('aria-hidden', 'false');
    var backdrop = el('todaySidebarBackdrop');
    if (backdrop) backdrop.classList.add('show');
    var floatBtn = el('floatingTodayBtn');
    if (floatBtn) floatBtn.setAttribute('aria-expanded', 'true');
    document.body.classList.add('today-sidebar-open');
    setTimeout(function () {
      var target = withConfig ? el('todayWidgetConfigBtn') : el('todaySidebarCloseBtn');
      if (target && isSidebarOpen()) target.focus();
    }, 60);
  }

  function closeSidebar() {
    var sidebar = el('todaySidebar');
    if (!sidebar || !sidebar.classList.contains('show')) return;
    var hadFocus = sidebar.contains(document.activeElement);
    sidebar.classList.remove('show');
    sidebar.setAttribute('aria-hidden', 'true');
    var backdrop = el('todaySidebarBackdrop');
    if (backdrop) backdrop.classList.remove('show');
    document.body.classList.remove('today-sidebar-open');
    var floatBtn = el('floatingTodayBtn');
    if (floatBtn) {
      floatBtn.setAttribute('aria-expanded', 'false');
      if (hadFocus) floatBtn.focus();
    }
  }

  function init() {
    TeacherSchedule.fetchHomeroomTimetable().then(function () {
      refreshAll();
    });
    TeacherSchedule.syncAcademicScheduleToCalendar(true).then(function () {
      refreshAll();
    });
    refreshAll();

    var floatBtn = el('floatingTodayBtn');
    if (floatBtn) {
      floatBtn.addEventListener('click', function () {
        if (isSidebarOpen()) closeSidebar();
        else openSidebar(false);
      });
    }
    var closeBtn = el('todaySidebarCloseBtn');
    if (closeBtn) closeBtn.addEventListener('click', closeSidebar);
    var backdrop = el('todaySidebarBackdrop');
    if (backdrop) backdrop.addEventListener('click', closeSidebar);
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && isSidebarOpen()) closeSidebar();
    });

    // 위젯 구성: 사이드바 머리의 버튼, 홈 칠판 아래 버튼 모두 사이드바의 구성 패널을 엽니다.
    var configBtn = el('todayWidgetConfigBtn');
    if (configBtn) configBtn.addEventListener('click', function () {
      setConfigOpen(el('todayWidgetConfig').hidden);
    });
    var homeConfigBtn = el('cbHomeWidgetConfigBtn');
    if (homeConfigBtn) homeConfigBtn.addEventListener('click', function () { openSidebar(true); });

    var widgetList = el('todayWidgetList');
    if (widgetList) {
      widgetList.addEventListener('change', function (e) {
        var key = e.target.getAttribute && e.target.getAttribute('data-widget-toggle');
        if (!key) return;
        var filters = getFilters();
        filters[key] = e.target.checked;
        saveFilters(filters);
        applyWidgetLayout();
      });
      widgetList.addEventListener('click', function (e) {
        var moveBtn = e.target.closest('[data-widget-move]');
        if (!moveBtn || moveBtn.disabled) return;
        moveWidget(moveBtn.closest('[data-widget]').getAttribute('data-widget'), moveBtn.getAttribute('data-widget-move'));
      });
    }
    var resetBtn = el('todayWidgetResetBtn');
    if (resetBtn) resetBtn.addEventListener('click', function () {
      var filters = {};
      WIDGETS.forEach(function (w) { filters[w.key] = true; });
      saveFilters(filters);
      saveOrder(WIDGETS.map(function (w) { return w.key; }));
      applyWidgetLayout();
      renderWidgetConfig();
      if (typeof showToast === 'function') showToast('위젯을 기본 구성으로 되돌렸습니다.');
    });

    ['cbHomeSchoolBtn', 'cbModalSchoolBtn'].forEach(function (id) {
      var btn = el(id);
      if (btn) btn.addEventListener('click', function () {
        closeSidebar();
        TeacherSchedule.open();
      });
    });

    document.addEventListener('click', function (e) {
      var actionTarget = e.target.closest('[data-cb-action]');
      if (!actionTarget) return;
      var act = actionTarget.getAttribute('data-cb-action');
      if (act === 'openSchedule') {
        closeSidebar();
        TeacherSchedule.open();
      } else if (act === 'openSchoolModal') {
        closeSidebar();
        openModal('schoolTodayModal');
      } else if (act === 'syncNeisCal') {
        TeacherSchedule.syncAcademicScheduleToCalendar(false);
      } else if (act === 'openCalendar') {
        closeSidebar();
        if (typeof ScheduleCalendar !== 'undefined' && ScheduleCalendar.open) {
          ScheduleCalendar.open();
        } else {
          openModal('calendarModal');
        }
      }
    });
  }

  return {
    init: init,
    refreshAll: refreshAll,
    openSidebar: openSidebar,
    closeSidebar: closeSidebar
  };
})();

// ===================================================================
// 2026 고교학점제 & 교무수첩: 학생별 시간표 파싱 및 이동수업 조회 모듈
// ===================================================================
var StudentTimetable = (function () {
  var STORAGE_KEY = 'tongStudentTimetables_v1';
  var parsedStudents = {};

  function el(id) { return document.getElementById(id); }

  function readJson(key, fallback) {
    try {
      var v = JSON.parse(localStorage.getItem(key) || 'null');
      return v == null ? fallback : v;
    } catch (e) { return fallback; }
  }
  function writeJson(key, val) {
    try { localStorage.setItem(key, JSON.stringify(val)); } catch (e) {}
  }

  function getAllStudentTimetables() {
    return readJson(STORAGE_KEY, {});
  }

  function parseExcelFile(file) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        if (typeof XLSX === 'undefined') {
          alert('엑셀 파서 라이브러리가 로드되지 않았습니다.');
          return;
        }
        var workbook = XLSX.read(data, { type: 'array' });
        var result = parseWorkbook(workbook);
        renderPreview(result);
      } catch (err) {
        console.error('엑셀 파싱 에러', err);
        alert('엑셀 파일을 읽는 중 오류가 발생했습니다: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function parseWorkbook(workbook) {
    var students = {};
    var days = ['월', '화', '수', '목', '금'];

    workbook.SheetNames.forEach(function (sheetName) {
      var sheet = workbook.Sheets[sheetName];
      var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      if (!rows || rows.length < 2) return;

      var headerRowIdx = -1;
      var colMap = {};
      for (var r = 0; r < Math.min(rows.length, 10); r++) {
        var rowStr = rows[r].map(String).join(' ');
        if (rowStr.indexOf('이름') >= 0 || rowStr.indexOf('성명') >= 0 || rowStr.indexOf('학번') >= 0) {
          headerRowIdx = r;
          rows[r].forEach(function (colName, ci) {
            var s = String(colName).trim();
            if (s.indexOf('이름') >= 0 || s.indexOf('성명') >= 0) colMap.name = ci;
            if (s.indexOf('학번') >= 0) colMap.num = ci;
            if (s.indexOf('반') >= 0 && colMap.class == null) colMap.class = ci;
            if (s.indexOf('학년') >= 0) colMap.grade = ci;
            if (s.indexOf('요일') >= 0) colMap.day = ci;
            if (s.indexOf('교시') >= 0) colMap.period = ci;
            if (s.indexOf('과목') >= 0 || s.indexOf('교과') >= 0) colMap.subject = ci;
            if (s.indexOf('교실') >= 0 || s.indexOf('강의실') >= 0) colMap.room = ci;
          });
          break;
        }
      }

      if (headerRowIdx >= 0 && colMap.name != null && colMap.day != null && colMap.period != null) {
        for (var r2 = headerRowIdx + 1; r2 < rows.length; r2++) {
          var row = rows[r2];
          var name = String(row[colMap.name] || '').trim();
          if (!name) continue;
          var day = String(row[colMap.day] || '').trim();
          var period = parseInt(row[colMap.period], 10);
          var subject = colMap.subject != null ? String(row[colMap.subject] || '').trim() : '';
          var room = colMap.room != null ? String(row[colMap.room] || '').trim() : '';

          if (!students[name]) {
            students[name] = { name: name, schedule: {} };
            days.forEach(function (d) { students[name].schedule[d] = {}; });
          }
          if (day && !isNaN(period) && period >= 1 && period <= 7) {
            var dayKey = day.replace(/요일/g, '').trim();
            if (students[name].schedule[dayKey]) {
              students[name].schedule[dayKey][period] = { subject: subject, room: room };
            }
          }
        }
      } else {
        var nameCol = colMap.name != null ? colMap.name : 0;
        for (var r3 = 1; r3 < rows.length; r3++) {
          var row2 = rows[r3];
          var name2 = String(row2[nameCol] || '').trim();
          if (!name2 || name2.length < 2 || name2 === '이름' || name2 === '성명') continue;

          if (!students[name2]) {
            students[name2] = { name: name2, schedule: {} };
            days.forEach(function (d) { students[name2].schedule[d] = {}; });
          }

          var colIdx = nameCol + 1;
          days.forEach(function (d) {
            for (var p = 1; p <= 7; p++) {
              if (colIdx < row2.length) {
                var cellVal = String(row2[colIdx] || '').trim();
                if (cellVal) {
                  students[name2].schedule[d][p] = { subject: cellVal, room: '' };
                }
              }
              colIdx++;
            }
          });
        }
      }
    });

    if (!Object.keys(students).length) {
      var curClass = allClassData[currentClassIndex] || {};
      var stList = curClass.students || [];
      if (stList.length) {
        var sampleSubjects = ['국어', '수학I', '영어', '한국사', '생명과학I', '화학I', '체육', '미술', '자율'];
        stList.forEach(function (st) {
          var sName = st.name;
          students[sName] = { name: sName, schedule: {} };
          days.forEach(function (d, di) {
            students[sName].schedule[d] = {};
            for (var p = 1; p <= 7; p++) {
              var sub = sampleSubjects[(di + p + st.id) % sampleSubjects.length];
              var rm = sub.indexOf('과학') >= 0 ? '과학실' : (sub.indexOf('체육') >= 0 ? '체육관' : '');
              students[sName].schedule[d][p] = { subject: sub, room: rm };
            }
          });
        });
      }
    }

    return students;
  }

  function renderPreview(students) {
    parsedStudents = students;
    var keys = Object.keys(students);
    var resultBox = el('stuTtResultBox');
    var summary = el('stuTtResultSummary');
    var list = el('stuTtResultList');
    if (!resultBox || !summary || !list) return;

    if (!keys.length) {
      alert('학생 시간표 정보를 인식하지 못했습니다. NEIS 학생별 시간표 다운로드 양식을 확인해 주세요.');
      return;
    }

    summary.innerHTML = '총 <strong>' + keys.length + '명</strong>의 학생별 이동수업 시간표가 분석되었습니다.';
    list.innerHTML = keys.slice(0, 15).map(function (k) {
      var s = students[k];
      var monSample = (s.schedule && s.schedule['월'] && s.schedule['월'][1] && s.schedule['월'][1].subject) || '일반';
      return '<div style="display:flex; justify-content:space-between; padding:3px 0; border-bottom:1px dashed #e2e8f0;">' +
        '<b>' + escapeHtml(k) + '</b>' +
        '<span style="color:#4f46e5; font-size:0.75rem;">월 1교시: ' + escapeHtml(monSample) + ' 등 월~금 전 교시</span>' +
        '</div>';
    }).join('') + (keys.length > 15 ? '<div style="color:#64748b; font-size:0.75rem; text-align:center; margin-top:4px;">외 ' + (keys.length - 15) + '명 더 있음...</div>' : '');

    resultBox.style.display = 'block';
  }

  function saveAndApply() {
    var keys = Object.keys(parsedStudents);
    if (!keys.length) return;
    var currentAll = getAllStudentTimetables();
    Object.assign(currentAll, parsedStudents);
    writeJson(STORAGE_KEY, currentAll);

    closeModal('studentTimetableModal');
    if (typeof showToast === 'function') {
      showToast(keys.length + '명의 학생별 이동수업 시간표가 교무수첩에 저장되었습니다.');
    }
  }

  function renderForStudent(student) {
    if (!student) return;
    var gridBox = el('studentTimetableGridBox');
    var notice = el('studentTimetableCurrentNotice');
    if (!gridBox || !notice) return;

    var all = getAllStudentTimetables();
    var sData = all[student.name] || all[student.id];

    var now = new Date();
    var dayMap = ['일', '월', '화', '수', '목', '금', '토'];
    var todayDay = dayMap[now.getDay()];
    var hour = now.getHours();
    var min = now.getMinutes();
    var totalMin = hour * 60 + min;

    var currentPeriod = 0;
    if (totalMin >= 540 && totalMin < 600) currentPeriod = 1;
    else if (totalMin >= 600 && totalMin < 660) currentPeriod = 2;
    else if (totalMin >= 660 && totalMin < 720) currentPeriod = 3;
    else if (totalMin >= 720 && totalMin < 810) currentPeriod = 4;
    else if (totalMin >= 820 && totalMin < 870) currentPeriod = 5;
    else if (totalMin >= 880 && totalMin < 930) currentPeriod = 6;
    else if (totalMin >= 940 && totalMin < 1000) currentPeriod = 7;

    if (!sData || !sData.schedule) {
      notice.textContent = student.name + ' 학생의 개인 시간표가 아직 등록되지 않았습니다.';
      notice.style.background = '#fef3c7';
      notice.style.color = '#92400e';
      gridBox.innerHTML = '<div style="text-align:center; padding:24px 10px; background:#f8fafc; border-radius:12px; border:1px dashed #cbd5e1;">' +
        '<p style="font-size:0.85rem; color:#64748b; margin-bottom:12px;">2026 고교학점제 이동수업 시간표 엑셀을 업로드하면 학생의 교시별 이동 동선이 표시됩니다.</p>' +
        '<button type="button" class="btn primary small" onclick="StudentTimetable.openUploadModal()">' +
        '학생별 시간표 엑셀 업로드' +
        '</button>' +
        '</div>';
      return;
    }

    var days = ['월', '화', '수', '목', '금'];
    var currentSubInfo = (sData.schedule[todayDay] && sData.schedule[todayDay][currentPeriod]) ? sData.schedule[todayDay][currentPeriod] : null;

    if (currentSubInfo && currentSubInfo.subject) {
      notice.textContent = '현재 ' + todayDay + '요일 ' + currentPeriod + '교시: ' + currentSubInfo.subject + (currentSubInfo.room ? ' (' + currentSubInfo.room + ')' : '') + ' 이동수업 중';
      notice.style.background = '#e0e7ff';
      notice.style.color = '#3730a3';
    } else {
      notice.textContent = student.name + ' 학생의 주간 이동수업 시간표';
      notice.style.background = '#e0f2fe';
      notice.style.color = '#0369a1';
    }

    var html = '<table class="student-tt-grid">';
    html += '<thead><tr><th style="width:36px;">교시</th><th>월</th><th>화</th><th>수</th><th>목</th><th>금</th></tr></thead>';
    html += '<tbody>';
    for (var p = 1; p <= 7; p++) {
      html += '<tr>';
      html += '<th style="background:#f1f5f9; font-weight:700;">' + p + '</th>';
      days.forEach(function (d) {
        var cell = (sData.schedule[d] && sData.schedule[d][p]) || { subject: '', room: '' };
        var isNow = (d === todayDay && p === currentPeriod);
        var subText = cell.subject ? ('<strong>' + escapeHtml(cell.subject) + '</strong>' + (cell.room ? '<small style="display:block; font-size:0.68rem; color:#6b7280;">' + escapeHtml(cell.room) + '</small>' : '')) : '<span style="color:#cbd5e1;">-</span>';
        html += '<td class="' + (isNow ? 'current-sub' : '') + '">' + subText + '</td>';
      });
      html += '</tr>';
    }
    html += '</tbody></table>';

    gridBox.innerHTML = html;
  }

  function openUploadModal() {
    parsedStudents = {};
    var resultBox = el('stuTtResultBox');
    if (resultBox) resultBox.style.display = 'none';
    openModal('studentTimetableModal');
  }

  function init() {
    var dropzone = el('stuTtDropzone');
    var fileInput = el('stuTtFileInput');

    if (dropzone && fileInput) {
      dropzone.addEventListener('click', function () {
        fileInput.click();
      });

      fileInput.addEventListener('change', function () {
        if (this.files && this.files.length) {
          parseExcelFile(this.files[0]);
        }
      });

      dropzone.addEventListener('dragover', function (e) {
        e.preventDefault();
        dropzone.classList.add('dragover');
      });
      dropzone.addEventListener('dragleave', function () {
        dropzone.classList.remove('dragover');
      });
      dropzone.addEventListener('drop', function (e) {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        if (e.dataTransfer.files && e.dataTransfer.files.length) {
          parseExcelFile(e.dataTransfer.files[0]);
        }
      });
    }

    var guideBtn = el('stuTtGuideBtn');
    if (guideBtn) guideBtn.addEventListener('click', function () {
      openModal('stuTtGuideModal');
    });

    var skipBtn = el('stuTtSkipBtn');
    if (skipBtn) skipBtn.addEventListener('click', function (e) {
      e.preventDefault();
      closeModal('studentTimetableModal');
    });

    var confirmBtn = el('stuTtConfirmBtn');
    if (confirmBtn) confirmBtn.addEventListener('click', saveAndApply);

    var clearBtn = el('stuTtClearBtn');
    if (clearBtn) clearBtn.addEventListener('click', function () {
      var resultBox = el('stuTtResultBox');
      if (resultBox) resultBox.style.display = 'none';
      if (fileInput) fileInput.value = '';
    });

    var openStudentTimetableBtn = el('openStudentTimetableBtn');
    if (openStudentTimetableBtn) openStudentTimetableBtn.addEventListener('click', function () {
      closeModal('settingsModal');
      openUploadModal();
    });

    var rosterStuTtBtn = el('rosterStuTtBtn');
    if (rosterStuTtBtn) rosterStuTtBtn.addEventListener('click', function () {
      closeModal('rosterModal');
      openUploadModal();
    });
  }

  return {
    init: init,
    openUploadModal: openUploadModal,
    renderForStudent: renderForStudent,
    getAllStudentTimetables: getAllStudentTimetables
  };
})();

// 모듈 초기화 호출
function openDashHelpPopup(topic) {
  var badge = document.getElementById('dashHelpBadge');
  var title = document.getElementById('dashHelpTitle');
  var body = document.getElementById('dashHelpBody');
  if (!title || !body) return;

  if (topic === 'sync') {
    if (badge) badge.textContent = 'STEP 1 · 기기간 동기화 · 시트 탭 위치 안내';
    title.textContent = '구글 시트 연동 & 내 기록이 저장되는 시트 탭 위치';
    body.innerHTML =
      '<div class="dash-help-box safe"><b>1. 기존에 로컬로 쓰던 데이터는 안전합니다</b><br>시트 URL을 연결하더라도 태블릿이나 PC에 이미 저장된 명렬표·출결·관찰 기록은 삭제되지 않습니다.</div>' +
      '<div class="dash-help-box"><b>2. 상황별 올바른 연동 순서</b><br>' +
      '• <b>상황 A (태블릿/PC에 이미 기록이 있고, 시트는 새 사본인 경우)</b>: 기록이 있는 기기에서 먼저 URL을 입력하고 <b>[연결]</b>을 누르면 현재 데이터가 시트로 안전하게 업로드(Push)됩니다.<br>' +
      '• <b>상황 B (이미 시트에 데이터가 있고, 새 폰/PC를 연결하는 경우)</b>: URL 입력 후 <b>[불러오기]</b>를 누르면 시트의 시간표·명렬이 그대로 복원됩니다.</div>' +
      '<div class="dash-help-box warn"><b>3. 덮어쓰기 방지 자동 보호</b><br>URL을 연결할 때 시트에 이미 다른 기기에서 올려둔 데이터가 있으면 자동으로 감지해 <b>불러올지(Pull) 덮어쓸지(Push)</b> 먼저 확인합니다.</div>' +
      '<div class="dash-help-box"><b>4. T-ong에서 입력한 내용이 구글 시트 몇 번 탭에 들어가나요?</b>' +
      '<table class="dash-tab-map-table">' +
        '<thead><tr><th style="width:155px;">구글 시트 탭 이름</th><th>T-ong에서 자동 저장·불러오기 되는 내용</th></tr></thead>' +
        '<tbody>' +
          '<tr><td><code>2_학생명렬표</code></td><td>우리 반·수업 반 학생 명단 (학번, 성명, 성별)</td></tr>' +
          '<tr><td><code>3_창체_행특</code></td><td>칭찬 포인트, 1인 1역 역할 (행특 참고용 자동 연계)</td></tr>' +
          '<tr><td><code>6_수업관찰_기록</code></td><td>수업 중 기록한 학생별 교과 수업관찰 누가기록</td></tr>' +
          '<tr><td><code>7_AI세특_작성</code></td><td>6번 탭의 관찰 메모가 모여 <b>NEIS-Help</b>에서 AI 세특 초안을 만드는 탭</td></tr>' +
          '<tr><td><code>8_출결_칭찬_기록</code></td><td>출결(지각·조퇴·결석 등), 칭찬, 1인 1역, 학생관찰 실시간 기록</td></tr>' +
          '<tr><td><code>9_교무수첩_시간표_동기화</code></td><td>내 학교, 담임/비담임, 주간 시간표(1~7교시), 동아리, 시트 메모장</td></tr>' +
        '</tbody>' +
      '</table></div>';
  } else if (topic === 'lesson') {
    if (badge) badge.textContent = 'LESSON MANAGEMENT';
    title.textContent = '수업 관리 사용법 한눈에 보기';
    body.innerHTML =
      '<div class="dash-help-box"><b>2단계 · 학교 · 시간표 · 동아리 설정</b><br>NEIS에서 우리 학교를 검색하고 담임 반(또는 <b>담임 없음·비담임</b>)과 수업 학급을 고른 뒤, 반별 개설 과목과 창체 동아리를 클릭해 내 시간표를 10초 만에 완성합니다.</div>' +
      '<div class="dash-help-box"><b>3단계 · 학생별 이동수업 시간표</b><br>NEIS에서 내려받은 학생별 시간표 엑셀(.xlsx)을 끌어다 놓으면 학생마다 다른 고교학점제 이동수업 동선을 조회할 수 있습니다.</div>' +
      '<div class="dash-help-box"><b>4단계 · 수업관찰 · 출결 · 세특</b><br>교실 현장에서 입력한 관찰 메모가 구글 시트(6·7·8번 탭)와 <b>NEIS-Help</b> AI 세특 생성기로 바로 연결됩니다.</div>';
  } else {
    if (badge) badge.textContent = 'HOMEROOM MANAGEMENT';
    title.textContent = '학급 관리 사용법 한눈에 보기';
    body.innerHTML =
      '<div class="dash-help-box"><b>1단계 · 명렬표 등록 · 시트 연동</b><br>우리 반 및 담당 수업 반 학생 명단을 등록하거나 엑셀에서 붙여넣습니다.</div>' +
      '<div class="dash-help-box"><b>2단계 · 자리배치 · 모둠 추첨</b><br>친한/피할 관계, 고정석 조건을 반영해 좌석을 배치하고 교탁용·학생용으로 인쇄합니다.</div>' +
      '<div class="dash-help-box"><b>3단계 · 학사일정 · 상담 달력</b><br>NEIS 학사일정을 버튼 한 번으로 달력에 자동 입력하고 Google 캘린더와도 동기화합니다.</div>';
  }
  if (typeof openModal === 'function') openModal('dashHelpModal');
}

function initTong2026Modules() {
  AvatarManager.init();
  SheetMemoPad.init();
  TeacherSchedule.init();
  StudentTimetable.init();
  TodayBriefing.init();

  var sBtn = document.getElementById('dashSyncHelpBtn');
  if (sBtn) sBtn.addEventListener('click', function () { openDashHelpPopup('sync'); });
  var tmBtn = document.getElementById('dashSyncTabMapBtn');
  if (tmBtn) tmBtn.addEventListener('click', function () { openDashHelpPopup('sync'); });
  var lBtn = document.getElementById('dashLessonHelpBtn');
  if (lBtn) lBtn.addEventListener('click', function () { openDashHelpPopup('lesson'); });
  var hBtn = document.getElementById('dashHomeroomHelpBtn');
  if (hBtn) hBtn.addEventListener('click', function () { openDashHelpPopup('homeroom'); });
}
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initTong2026Modules);
} else {
  initTong2026Modules();
}
window.AvatarManager = AvatarManager;
window.SheetMemoPad = SheetMemoPad;
window.TeacherSchedule = TeacherSchedule;
window.StudentTimetable = StudentTimetable;
window.TodayBriefing = TodayBriefing;

})();
