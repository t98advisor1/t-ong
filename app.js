(function() {
'use strict';

var APP_VERSION = '5.17.0-preview.72';
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

function showHomeScreen() {
  appMode = 'home';
  document.querySelectorAll('.mode-tab[data-mode]').forEach(function (tab) { tab.classList.remove('active'); });
  var wizard = document.getElementById('setupWizard');
  var stage = document.getElementById('stage');
  var attend = document.getElementById('attendView');
  var bar = document.querySelector('.action-bar');
  if (wizard) { wizard.classList.remove('hidden'); wizard.classList.add('home-only'); wizard.style.display = 'grid'; }
  setModeScreen(stage, false);
  setModeScreen(attend, false);
  setModeScreen(bar, false);
  window.scrollTo({ top:0, behavior:'auto' });
}

function showSetupWizard() {
  var wizard = document.getElementById('setupWizard');
  if (wizard) {
    wizard.classList.remove('hidden');
    wizard.classList.remove('home-only');
    wizard.style.display = 'grid';
  }
  setModeScreen(document.getElementById('stage'), false);
  if (appMode === 'seat') setModeScreen(document.querySelector('.action-bar'), false);
  document.getElementById('totalDesks').value = config.totalDesks;
  document.getElementById('studentCount').value = students.length || 28;
  initColumnOptions();
  updateLayoutOptions();
  updatePreview();
  updateEmptySeatInfo();
  if (typeof showWizardStep === 'function') showWizardStep(1);
}

function hideSetupWizard() {
  var wizard = document.getElementById('setupWizard');
  if (wizard) {
    wizard.classList.remove('home-only');
    wizard.classList.add('hidden');
    wizard.style.display = 'none';
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
function openModal(id) {
  document.getElementById(id).classList.add('show');
}

function closeModal(id) {
  document.getElementById(id).classList.remove('show');
}

function saveToLocalStorage() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({
    settings: settings,
    allClassData: allClassData,
    currentClassIndex: currentClassIndex
  }));
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

function showToast(msg, type, actionLabel, action) {
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
  }, actionLabel ? 5200 : 2500);
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
  {
    img: 'guide/00-layout.png',
    title: '한 화면에서 수업과 학급 운영을 관리합니다',
    desc: '<ul>' +
      '<li><strong>① 상단 메뉴</strong> — 의자 홈·자리배치·출결관리·교실운영·일정관리·즐겨찾기·게임을 전환합니다.</li>' +
      '<li><strong>② 좌측 메뉴</strong> — 설정·사용법·저장 기록·백업처럼 가끔 쓰는 기능을 모았습니다.</li>' +
      '<li><strong>③ 좌석표</strong> — 교탁을 기준으로 복도(왼쪽)와 창가(오른쪽)가 표시됩니다.</li>' +
      '<li><strong>④ 우측 패널</strong> — 명렬표·좌석 설정·추첨·저장·출력을 바로 실행합니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '페이지 스크롤을 열어 두었습니다. 화면이 작거나 좌석이 많아도 아래로 내려가며 확인할 수 있습니다.' }
  },
  {
    section: 'start',
    img: 'guide/v67-home-backup.png',
    title: '상단 탭에서 필요한 화면을 고릅니다',
    desc: '<ul>' +
      '<li><strong>🏠 홈</strong> — 하위메뉴에서 첫 화면·사용 설명서·<strong>① 📦 데이터 백업</strong>을 엽니다.</li>' +
      '<li><strong>자리배치 · ✅ 출결관리</strong> — 좌석표와 날짜별 출결 화면을 오갑니다.</li>' +
      '<li><strong>➕ 교실운영</strong> — 컴시간알리미(교사용)·담임 5분 멘트·타이머·칭찬·사다리·1인 1역·룰렛·알림장을 엽니다.</li>' +
      '<li><strong>📅 일정관리</strong> — 전체화면 달력과 상담 일정을 엽니다.</li>' +
      '<li><strong>🔖 즐겨찾기</strong> — 업무별 기본 사이트와 직접 추가한 내 사이트를 엽니다.</li><li><strong>🎮 게임</strong> — 수업 전환과 참여에 활용할 교실 게임을 모았습니다.</li>' +
      '<li><strong>❓ 도움말</strong> — 버전·업데이트·오늘의 학교(실험)·오류 문의를 확인합니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '홈과 출결관리는 입력 중인 학급 자료를 그대로 유지합니다.' }
  },
  {
    section: 'classroom',
    img: 'guide/0B-dock.png',
    title: '➕ 교실운영 — 수업과 학급 운영 도구를 한곳에',
    desc: '<ul>' +
      '<li><strong>🕒 컴시간알리미(교사용)</strong> — 학교 코드를 설정해 교사용 개인 시간표를 확인합니다.</li>' +
      '<li><strong>💬 담임 5분 멘트</strong> — 조회 시간에 사용할 뉴스·명언·감동 이야기를 빠르게 준비합니다.</li>' +
      '<li><strong>⏱️ 간편 타이머</strong> — 3분·5분·10분 또는 직접 입력한 시간을 전체화면으로 보여줍니다.</li>' +
      '<li><strong>🏆 칭찬 포인트</strong> — +1·−1로 점수를 조절하고 알·부화 중·병아리·닭 성장 단계를 확인합니다.</li>' +
      '<li><strong>🪜 사다리타기</strong> — 발표 뽑기는 사다리타기로! 이름과 결과를 재미있게 연결합니다.</li>' +
      '<li><strong>🧹 1인 1역</strong> — 학생의 역할을 메모해 두세요. 엑셀로 불러오거나 무작위로 배정할 수 있습니다.</li>' +
      '<li><strong>🎯 룰렛 뽑기</strong> — 발표자나 활동 인원을 한 명 또는 여러 명 무작위로 뽑습니다.</li>' +
      '<li><strong>📢 알림장</strong> — 준비물·과제·공지사항을 작성하고 교실 TV에 크게 표시합니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '교실운영 메뉴는 현재 좌석표를 떠나지 않고 열리며, 입력한 자료는 학급별로 저장됩니다.' }
  },  {
    section: 'calendar',
    img: 'guide/v67-calendar.png',
    title: '📅 일정관리 ① — 월간·주간·오늘 보기',
    desc: '<ul>' +
      '<li><strong>월간일정 · 주간일정 · 오늘일정</strong> 탭으로 필요한 범위를 바로 바꿉니다.</li>' +
      '<li><strong>휠·트랙패드</strong>를 움직이거나 달력을 좌우로 밀면 이전·다음 달, 주, 날짜로 이동합니다.</li>' +
      '<li>월이 바뀌는 첫날은 <strong>8/1</strong>처럼 월과 날짜를 함께 표시합니다.</li>' +
      '<li><strong>날짜 더블클릭</strong> 또는 ＋ 일정 추가로 입력 팝업을 엽니다.</li>' +
      '<li><strong>수정·삭제</strong> — 월간 일정 카드를 한 번 누르면 수정 창이 열리고, 아래 🗑️ 일정 삭제 버튼으로 지웁니다.</li>' +
      '<li><strong>일정 이동</strong> — 월간·주간 일정 카드를 원하는 날짜로 끌어 놓습니다. 모바일에서는 일정을 약 0.4초 길게 누른 뒤 이동합니다.</li>' +
      '<li><strong>여러 날 일정</strong>은 전체 기간을 그대로 유지해 이동하며, 이동 직후 <strong>실행 취소</strong>할 수 있습니다.</li>' +
      '<li>일정은 <strong>현재 브라우저에 저장</strong>되고 백업 코드에 포함됩니다. Google 계정 연결과 자동 동기화가 유효하면 전송 대상 일정의 이동도 본인 계정에 반영됩니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '드래그 이동은 날짜만 바꾸며 제목·속성·메모·상담 기록과 여러 날 일정의 기간은 그대로 유지합니다.' }
  },
  {
    section: 'calendar',
    img: 'guide/v67-calendar-delete.png',
    title: '일정 수정과 삭제 — 카드에서 바로 엽니다',
    desc: '<ul>' +
      '<li>월간 달력의 <strong>일정 카드</strong>를 한 번 누르면 수정 창이 열립니다.</li>' +
      '<li>제목·날짜·속성·메모를 고친 뒤 <strong>수정 저장</strong>을 누릅니다.</li>' +
      '<li>지우려면 아래 <strong>🗑️ 일정 삭제</strong>를 누르고 확인합니다. 여러 날 일정은 전체 일정이 삭제됩니다.</li>' +
      '<li>주간·오늘 보기에서도 일정 옆 <strong>삭제</strong> 버튼을 사용할 수 있습니다.</li>' +
      '</ul>',
    callout: { type: 'danger', text: '삭제는 드래그 이동의 실행 취소와 다릅니다. 연결된 Google 일정 삭제도 동기화될 수 있으므로 중요한 일정은 백업 후 처리하세요.' }
  },
  {
    section: 'calendar',
    img: 'guide/22A-calendar-properties.png',
    title: '📅 일정관리 ② — 속성과 상담 기록',
    desc: '<ul>' +
      '<li><strong>속성 없음</strong>을 선택하면 일정명 앞에 꼬리표 없이 저장합니다.</li>' +
      '<li><strong>기본 속성</strong>은 상담·출장·회의이며 필요할 때만 선택합니다.</li>' +
      '<li><strong>＋ · ✎ · −</strong> 버튼으로 필요한 속성을 추가하고, 이름을 수정하거나 삭제합니다.</li>' +
      '<li>달력에는 <strong>[상담] · [출장] · [회의]</strong>처럼 일정명 앞에 속성이 표시됩니다.</li>' +
      '<li><strong>상담</strong>을 고르면 학생 또는 학부모 대상·관련 학생·상담 방식·진행 상태·후속 조치 입력칸이 열립니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '직접 추가한 속성도 이 브라우저에 저장되고 데이터 백업 코드에 포함됩니다.' }
  },
  {
    section: 'calendar',
    img: 'guide/22B-calendar-export.png',
    title: '📅 일정관리 ③ — 가져오기와 상담 기록 내보내기',
    desc: '<ul>' +
      '<li><strong>구글시트 가져오기</strong> — 날짜·종료일·일정명·속성·메모 열을 복사해 붙여넣습니다.</li>' +
      '<li><strong>Google 동기화</strong> — 로그인 계정의 Google 기본 캘린더와 양방향으로 연결합니다. 자동 동기화를 켜고 앱을 활성 상태로 열어 두면 변경 후·약 1분 간격으로 확인합니다.</li>' +
      '<li><strong>개인정보 보호</strong> — 상담 일정과 메모는 기본적으로 Google에 보내지 않으며, 접근 토큰은 현재 탭의 메모리에만 둡니다.</li>' +
      '<li><strong>ICS</strong> — 전체 일정을 Google 캘린더로 옮길 일회성 파일을 만듭니다.</li>' +
      '<li><strong>상담기록 엑셀</strong> — 상담 속성의 기록만 학생별 열로 정리해 내려받습니다.</li>' +
      '<li><strong>상담기록 시트 복사</strong> — 표 형태로 복사해 구글시트 첫 칸에 바로 붙여넣습니다. 옆의 <strong>?</strong>에서 실제 붙여넣기 예시를 확인합니다.</li>' +
      '</ul>',
    callout: { type: '', text: '상담 기록에는 날짜·상담 대상·관련 학생·상담 방식·진행 상태·상담 내용·후속 조치가 포함됩니다.' }
  },
  {
    section: 'calendar',
    img: 'guide/v67-google-calendar-sync.png',
    title: '📅 일정관리 ④ — Google 캘린더 양방향 동기화',
    desc: '<ul>' +
      '<li><strong>앱 관리자 설정의 ? 버튼</strong>에서 OAuth 클라이언트 ID 발급과 연결 과정을 4단계 화면으로 확인합니다.</li>' +
      '<li><strong>① Google 계정 연결</strong>을 누르면 로그인 계정의 <strong>Google 기본 캘린더</strong>와 연결합니다.</li>' +
      '<li>연결 권한이 유효하고 자동 동기화가 켜져 있으면 T-ong의 변경 후 동기화하고, 활성 화면에서 Google 변경을 <strong>약 1분 간격</strong>으로 확인합니다. 앱을 닫으면 자동 확인도 멈춥니다.</li>' +
      '<li><strong>지금 양방향 동기화</strong>로 수동 실행합니다. 다른 캘린더·공유 캘린더의 일정은 기본 캘린더 동기화 대상이 아닙니다.</li>' +
      '<li><strong>② Google 일정 전체 다시 가져오기</strong>는 증분 기준을 사용하지 않고 기본 캘린더 전체를 다시 확인합니다. 학급·학생 자료는 지우지 않지만 <strong>양방향</strong>이므로 T-ong의 일정 수정·삭제도 Google에 반영됩니다.</li>' +
      '<li><strong>상담 일정 · 메모 전송</strong>은 개인정보 보호를 위해 기본 해제되어 있습니다. 필요한 경우에만 직접 켭니다.</li>' +
      '<li>접근 토큰은 현재 탭의 메모리에만 있으며 만료되면 <strong>Google 계정 연결</strong>을 다시 누릅니다.</li>' +
      '<li>공용 PC에서 서로 다른 교사가 쓸 때는 각자의 브라우저 프로필을 사용합니다.</li>' +
      '</ul>',
    callout: { type: '', text: '사진은 계정 연결 전 실제 화면입니다. 누락 시 같은 계정의 기본 캘린더인지 확인하세요. 전체 다시 가져오기 전에는 백업하고 확인창을 읽으세요. 시간 일정은 날짜 중심으로 표시됩니다.' }
  },
  {
    section: 'favorites',
    img: 'guide/23-custom-favorites.png',
    title: '🔖 즐겨찾기 — 교사용 사이트를 업무별로 정리합니다',
    desc: '<ul>' +
      '<li><strong>수학관련 사이트</strong> — ASK MATH·지오지브라·데스모스·통그라미와 직접 추가한 사이트를 함께 관리합니다.</li>' +
      '<li><strong>수업 및 클래스</strong> — 미리캔버스·클래스룸·AI 교수학습 플랫폼·EBSi 교사지원센터·캔바와 직접 추가한 사이트를 함께 관리합니다.</li>' +
      '<li><strong>교과서 사이트</strong> — 기본 사이트와 직접 추가한 사이트를 함께 관리합니다.</li>' +
      '<li><strong>생성형 AI</strong> — ChatGPT·Claude·Gemini·Perplexity·NotebookLM·Qwen을 모았습니다.</li>' +
      '<li><strong>교무업무</strong> — 교육과정 성취기준과 학교생활기록부를 엽니다.</li>' +
      '<li><strong>진로진학</strong> — 기본 사이트와 직접 추가한 사이트를 함께 관리합니다.</li>' +
      '<li><strong>교사 도구함</strong> — 기출문항 크롭퍼부터 프롬프트·업무 도구까지 모아 둡니다.</li>' +
      '<li><strong>나만의 즐겨찾기</strong>는 항상 메뉴 맨 아래에 있으며 이름과 주소를 직접 등록·수정·삭제합니다.</li>' +
      '<li>사이트 이름은 <strong>긴 제목부터 짧은 제목 순서</strong>로 정리해 줄을 맞췄습니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '직접 추가한 즐겨찾기는 이 브라우저에만 저장되며 데이터 백업 코드에 포함됩니다.' }
  },
  {
    section: 'classroom',
    img: 'guide/23A-comcigan.png',
    title: '🕒 컴시간알리미(교사용) — 학교 코드를 설정해 내 시간표를 봅니다',
    desc: '<ul>' +
      '<li><strong>교실운영 맨 위 → 컴시간알리미(교사용)</strong>을 누르면 외부 사이트가 새 창으로 열립니다.</li>' +
      '<li>처음 이용할 때는 화면 안내에 따라 <strong>학교 코드</strong>를 입력해야 합니다.</li>' +
      '<li>학교 코드는 컴시간알리미를 관리하는 학교 담당자에게 확인할 수 있습니다.</li>' +
      '<li>설정 후 <strong>일자와 교사</strong>를 선택하면 교사용 개인 시간표를 확인할 수 있습니다.</li>' +
      '<li>컴시간알리미(교사용)는 외부 서비스이므로 학교 설정과 시간표 제공 상태는 해당 서비스 기준으로 동작합니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: 'T-ong은 컴시간의 학교 코드나 개인 시간표를 서버에 저장하지 않습니다. 컴시간알리미(교사용) 사이트가 새 창에서 직접 처리합니다.' }
  },  {
    section: 'games',
    img: 'guide/24-games.png',
    title: '🎮 게임 — 짧은 활동과 수업 전환에 활용합니다',
    desc: '<ul>' +
      '<li><strong>핸드러시</strong> — 손의 빠른 반응을 활용하는 자동차 활동 게임입니다.</li>' +
      '<li><strong>핸드 카트라이더</strong> — 손동작으로 카트를 조작하는 교실 활동 게임입니다.</li>' +
      '<li><strong>온몸으로 찰칵! 포즈 챌린지</strong> — 일곱 가지 포즈를 따라 하며 도전하는 몸동작 활동입니다. 연습 모드로 먼저 익힐 수 있습니다.</li>' +
      '<li><strong>야구게임</strong> — 학급이 함께 즐기는 야구 방식의 참여 게임입니다.</li>' +
      '<li><strong>SET(셋)</strong> — 카드의 색·모양·개수·채움을 관찰하여 같은 규칙의 세 장을 찾는 게임입니다.</li>' +
      '<li><strong>랜덤게임</strong> — 룰렛을 이용해 학생·활동·순서를 무작위로 정합니다.</li>' +
      '<li><strong>윷놀이</strong> — 학급 활동에 활용할 수 있는 온라인 윷놀이입니다.</li>' +
      '<li><strong>기지개팡팡</strong> — 수업 중 몸을 풀고 집중을 환기하는 활동입니다.</li>' +
      '<li><strong>모의주식투자</strong> — 가상 투자로 경제 개념과 의사결정을 경험합니다.</li>' +
      '<li><strong>거북목탈출</strong> — 바른 자세를 점검하고 움직임을 유도합니다.</li>' +
      '<li><strong>나는 가수다</strong> — 노래와 참여 활동을 위한 교실 게임입니다.</li>' +
      '<li><strong>안면근육훈련</strong> — 표정과 안면 근육을 움직이며 긴장을 풀고 집중을 환기합니다.</li>' +
      '</ul>',
    callout: { type: '', text: '게임은 새 탭에서 열립니다. 학교 계정 정책이나 외부 사이트 상태에 따라 로그인이 필요할 수 있습니다.' }
  },
  {
    section: 'help',
    img: 'guide/25-help.png',
    title: '❓ 도움말 — 설명서·업데이트·응원·데이터 안내',
    desc: '<ul>' +
      '<li><strong>홈 ▾ → 사용 설명서</strong>에서 이 캐러셀을 언제든 다시 엽니다.</li>' +
      '<li><strong>버전 안내</strong> — 현재 버전과 로컬 저장 원칙을 확인합니다.</li>' +
      '<li><strong>업데이트 사항</strong> — 최근 추가·변경된 기능을 확인합니다.</li>' +
      '<li><strong>개발자 응원하기</strong> — 응원 설문폼에 사용 후기와 따뜻한 한마디를 남깁니다.</li>' +
      '<li><strong>오늘의 학교(실험)</strong> — 학교를 설정하고 날짜별 학급 시간표와 급식을 조회합니다.</li>' +
      '<li><strong>오류·기능개선 문의</strong> — 사용 중 발견한 문제나 개선 의견을 보냅니다.</li>' +
      '<li><strong>데이터 저장 안내</strong> — 방문 통계와 브라우저 로컬 저장 범위를 확인합니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '명렬표·좌석·출결·상담 기록은 T-ong 서버가 아니라 현재 브라우저에 저장됩니다. Google 동기화는 사용자가 켠 일정만 본인 계정으로 직접 전송합니다.' }
  },
  {
    section: 'help',
    img: 'guide/26-school-today.png',
    title: '🏫 오늘의 학교(실험) — 시간표와 급식을 날짜별로',
    desc: '<ul>' +
      '<li><strong>학교 설정</strong>에서 학교 이름을 검색하고 학년·반을 한 번 저장합니다.</li>' +
      '<li>상단 날짜를 바꾸면 해당 날짜의 <strong>학급 시간표와 급식</strong>을 함께 조회합니다.</li>' +
      '<li>교사용 개인 시간표는 <strong>교실운영 맨 위 → 컴시간알리미(교사용)</strong>을 이용합니다.</li>' +
      '<li>학교·학년·반 설정은 현재 브라우저에만 저장됩니다.</li>' +
      '</ul>',
    callout: { type: '', text: '조회에는 NEIS 공개정보를 사용합니다. 휴일이거나 학교에서 아직 등록하지 않은 날은 결과가 비어 있을 수 있습니다.' }
  },  {
    img: 'guide/0C-leftmenu.png',
    title: '좌측 메뉴 — 설정과 기록을 정리했습니다',
    desc: '<ul>' +
      '<li><strong>☰</strong>을 누르면 아이콘 메뉴가 글자 메뉴로 펼쳐집니다.</li>' +
      '<li><strong>⚙️ 설정 · 명렬표</strong>에서 학급 명렬표, 좌석 배치, 데이터 백업을 관리합니다.</li>' +
      '<li><strong>📖 사용법</strong>은 이 설명서를 다시 엽니다.</li>' +
      '<li><strong>📦 데이터 · 기록</strong>에서 저장한 배치, 통합 기록, 백업, 초기화를 찾을 수 있습니다.</li>' +
      '</ul>',
    callout: { type: '', text: '평소에는 접어 두면 좌석표가 더 넓어집니다. 펼침 상태는 이 브라우저에 기억됩니다.' }
  },
  {
    img: 'guide/0D-seatpanel.png',
    title: '우측 좌석 설정 패널 — 추첨까지 한 흐름으로',
    desc: '<ul>' +
      '<li><strong>명렬표 업로드</strong> — 엑셀 업로드·직접 입력·샘플 양식을 선택합니다.</li>' +
      '<li><strong>좌석 설정</strong> — 현재 책상 수·분단 수·자리 형태를 확인하고 바꿉니다.</li>' +
      '<li><strong>추첨 시작</strong> — 현재 조건으로 새 자리를 추첨합니다.</li>' +
      '<li><strong>배치 순서</strong> — 무작위·학번 가로순·학번 세로순 중에서 고릅니다. 모든 자리 형태에 적용됩니다.</li>' +
      '<li><strong>저장 · 배치 불러오기</strong> — 지금 배치를 이력에 남기고 저장한 배치를 복원합니다.</li>' +
      '<li><strong>교탁 시점</strong> — 인쇄 없이 수업 중 화면 방향을 뒤집습니다. 출력은 별도 버튼입니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '주요 버튼은 같은 색·크기·정렬로 통일했습니다. 위에서 아래로 따라가면 됩니다.' }
  },
  {
    img: 'guide/01-wizard.png',    title: '자리배치 시작 — 3단계로 끝납니다',
    desc: '<ul>' +
      '<li>홈에서 상단 <strong>자리배치</strong>를 누르면 이 설정 화면으로 이동합니다.</li>' +
      '<li>위쪽 <strong>1 교실 크기 → 2 자리 형태 → 3 학생 수</strong> 탭을 따라가면 됩니다. 탭을 눌러 아무 단계로나 바로 이동할 수 있습니다.</li>' +
      '<li><strong>1단계</strong> 총 책상 개수와 열(분단) 수를 정합니다.</li>' +
      '<li><strong>2단계</strong> 자리 형태는 시험대형(1인) · 짝꿍(2인) · 트리플(3인) · <strong>모둠</strong> 중에 고릅니다.</li>' +
      '<li>학생 수를 입력하면 <strong>빈자리가 몇 개 생기는지</strong> 바로 알려줍니다.</li>' +
      '<li>오른쪽 미리보기로 실제 배열을 확인한 뒤 <strong>🎉 시작하기</strong>를 누르세요.</li>' +
      '</ul>'
  },
  {
    img: 'guide/14-group.png',
    title: '모둠 배치 — 책상을 붙여서 모둠으로',
    desc: '<p>자리 형태에서 <strong>모둠</strong>을 고르면 책상이 모둠 단위로 묶여 배치됩니다.</p>' +
      '<ul>' +
      '<li><strong>모둠 인원</strong>은 4명 · 5명 · 6명 중에 고릅니다.</li>' +
      '<li><strong>가로 모둠 수</strong>로 한 줄에 몇 모둠을 둘지 정합니다. (예: 3모둠 × 2줄 = 6모둠)</li>' +
      '<li>모둠마다 <strong>1모둠 · 2모둠 …</strong> 번호가 붙고, 좌석 카드에도 표시됩니다.</li>' +
      '<li>추첨 · 자리 고정 · 관계 설정 · 인쇄 모두 일반 배치와 똑같이 동작합니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '모둠 안에서도 <strong>💚 가까이 · 💔 멀리</strong> 관계가 그대로 반영됩니다.' }
  },
  {
    img: 'guide/20-attend.png',
    title: '✅ 출결관리 — 조회부터 기록까지 한 화면에서',
    desc: '<p><strong>상단의 출결관리 탭</strong>을 누르면 학급별 출결 화면으로 전환됩니다.</p>' +
      '<ul>' +
      '<li><strong>날짜 조회</strong> — 달력과 이전·다음 버튼으로 오늘뿐 아니라 지난 날짜의 기록도 바로 찾고 수정합니다.</li>' +
      '<li><strong>상태</strong> — 학생별로 결석·지각·조퇴·결과를 고릅니다.</li>' +
      '<li><strong>사유</strong> — 질병·인정·미인정·기타를 선택하고, 필요한 내용은 메모에 남깁니다.</li>' +
      '<li><strong>학생 이름</strong> — 이름을 누르면 좌석·칭찬·관찰·세특이 연결된 학생 종합 기록이 열립니다.</li>' +
      '<li><strong>상단 요약</strong> — 출석·결석·지각·조퇴·결과 인원이 즉시 집계됩니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '<strong>조회·종례에서 특히 중요합니다.</strong> 담임은 학급 출결을 빠르게 확인하고, 교과교사는 담당 학급을 바꿔 가며 같은 방식으로 기록할 수 있습니다.' }
  },
  {
    section: 'attend',
    img: 'guide/v67-student-record.png',
    title: '👤 출결 학생 팝업 — 칭찬과 관찰·세특까지 연결합니다',
    desc: '<ul>' +
      '<li>출결관리에서 <strong>학생 이름</strong>을 누르면 좌석표와 연결된 학생 종합 기록 팝업이 열립니다.</li>' +
      '<li><strong>⭐ 칭찬 +1</strong>과 <strong>➖ 칭찬 −1</strong>로 점수를 조절하며, 좌석 카드의 성장 배지에도 반영됩니다.</li>' +
      '<li>팝업 위쪽 <strong>오늘 출결 바로 등록</strong>은 오늘 날짜용입니다. 다른 날짜는 출결관리 화면에서 선택합니다.</li>' +
      '<li><strong>✍️ 관찰·세특</strong>에서 날짜·교과·기록 유형과 구체적인 관찰 내용을 누적합니다.</li>' +
      '<li><strong>종합</strong> 탭에서는 자리·짝·출결·1인 1역·칭찬·메모·관찰 기록을 한눈에 확인합니다.</li>' +
      '<li>학생번호를 기준으로 연결되므로 자리배치에서 열어도 같은 기록이 나타납니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '수업 중 출결을 확인하면서 바로 칭찬하거나 관찰 내용을 적어 두면, 통합 기록과 세특 엑셀·시트 복사에 함께 활용할 수 있습니다.' }
  },  {
    img: 'guide/20A-attend-excel.png',
    title: '📥 출결 엑셀 — 날짜별 기록과 월간 합계를 내려받습니다',
    desc: '<ul>' +
      '<li><strong>이번 달 · 지난 달 · 전체 기간</strong> 중 필요한 범위를 먼저 고릅니다.</li>' +
      '<li><strong>출결 엑셀</strong>을 누르면 날짜별 상태와 사유, 학생별 결석·지각·조퇴·결과 합계가 한 줄로 정리됩니다.</li>' +
      '<li><strong>사유·메모</strong> 열에는 날짜와 함께 입력한 내용이 모여 학급 출결 조회와 후속 확인에 유용합니다.</li>' +
      '<li><strong>구글시트로 복사</strong>를 누르면 새 시트에 바로 붙여넣을 수 있습니다.</li>' +
      '</ul>',
    callout: { type: '', text: '<strong>다운로드 후 초기화</strong>는 엑셀 저장을 먼저 시도한 뒤 브라우저 기록을 비웁니다. 학기말에는 백업 파일도 함께 보관하세요.' }
  },
  {
    img: 'guide/0E-five-minute-message.png',
    title: '💬 교실운영 — 담임 5분 멘트',
    desc: '<p>상단 <strong>교실운영 ▾ → 담임 5분 멘트</strong>를 누르면 조회 시간에 바로 활용할 수 있는 Gemini 도구가 열립니다.</p>' +
      '<ul>' +
      '<li>간단한 숫자만 입력해도 <strong>실시간 뉴스 · 마음 명언 · 감동 에피소드</strong>를 한 번에 구성합니다.</li>' +
      '<li>화면의 QR 코드를 학생용 기기나 휴대전화로 스캔해 열 수도 있습니다.</li>' +
      '<li><a class="guide-link" href="https://gemini.google.com/gem/6aafa9cba56f" target="_blank" rel="noopener noreferrer">담임 5분 멘트 바로 열기 ↗</a></li>' +
      '</ul>',
    callout: { type: 'good', text: '링크 응답을 확인했습니다. Gemini 맞춤 Gem 특성상 Google 로그인이나 학교 계정의 Gemini 사용 권한이 필요할 수 있습니다.' }
  },  {
    img: 'guide/17-timer.png',
    title: '⏱️ 타이머 — 시작하면 전체화면으로',
    desc: '<ul>' +
      '<li><strong>시작</strong>을 누르면 화면 전체가 타이머로 바뀝니다. 교실 뒤에서도 잘 보입니다.</li>' +
      '<li><strong>모래시계</strong>의 모래가 실제로 줄어들고, 원형 막대도 함께 줄어듭니다.</li>' +
      '<li>마지막 10초는 숫자가 <strong>빨갛게 커지며</strong> 똑딱 소리가 나고, 끝나면 알람과 색종이가 터집니다.</li>' +
      '<li><strong>작게 보기</strong>(또는 Esc)를 누르면 좌석표로 돌아갑니다. <strong>타이머는 계속 돌아갑니다.</strong></li>' +
      '<li>오른쪽 작은 시계를 다시 누르면 전체화면으로 돌아옵니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '조회·종례뿐 아니라 발표, 모둠활동, 시험 대형 정리 시간에도 쓰세요.' }
  },
  {
    img: 'guide/18-settings-steps.png',
    title: '⚙️ 설정 메뉴 — 번호 순서대로',
    desc: '<p>설정 메뉴는 <strong>실제 작업 순서대로 번호</strong>가 매겨져 있습니다.</p>' +
      '<ul>' +
      '<li><strong>1 명렬표 입력</strong> — 학번·이름·성별을 직접 입력합니다.</li>' +
      '<li><strong>1 엑셀로 한 번에 올리기</strong> — 이미 명렬표 파일이 있다면 이쪽이 빠릅니다. (둘 중 하나만 하면 됩니다)</li>' +
      '<li><strong>📄 엑셀 샘플 양식 받기</strong> — 형식을 모를 때 먼저 받아보세요.</li>' +
      '<li><strong>📊 현재 명렬표 엑셀 받기</strong> — 직접 입력하거나 불러온 현재 학급 명렬표를 다시 엑셀로 저장합니다.</li>' +
      '<li><strong>2 자리 배치 설정</strong> — 책상 수 · 열(모둠) 수 · 자리 형태를 정합니다.</li>' +
      '<li><strong>3 데이터 백업</strong> — 다른 컴퓨터로 옮기거나 학기말에 보관합니다.</li>' +
      '<li><strong>! 전체 초기화</strong> — 되돌릴 수 없으니 주의하세요.</li>' +
      '</ul>'
  },
  {
    img: 'guide/15-notice.png',
    title: '📢 오늘의 알림장',
    desc: '<ul>' +
      '<li>준비물 · 과제 · 공지사항을 적어두면 <strong>좌석표 옆에 계속 표시</strong>됩니다.</li>' +
      '<li><strong>🖥️ 크게 보기</strong>를 누르면 교실 TV에 띄우기 좋은 큰 화면으로 바뀝니다. (Esc로 닫기)</li>' +
      '<li>알림장은 <strong>학급마다 따로</strong> 저장됩니다.</li>' +
      '</ul>'
  },
  {
    section: 'classroom',
    img: 'guide/v67-praise-growth.png',
    title: '🏆 칭찬 포인트',
    desc: '<ul>' +
      '<li><strong>학생 팝업의 칭찬 +1·−1</strong>로 바로 조절할 수 있고, 칭찬 목록에서는 이름을 누르면 +1, 오른쪽 −를 누르면 1점이 차감됩니다.</li>' +
      '<li><strong>🥚 알 0점 → 🐣 부화 중 1~4점 → 🐥 병아리 5~9점 → 🐔 닭 10점 이상</strong>으로 성장합니다. 실제 숫자 점수도 함께 표시합니다.</li>' +
      '<li>학생 팝업·칭찬 목록·좌석 카드의 성장 표시가 연결됩니다. 0점에서는 −1이 비활성화됩니다.</li>' +
      '<li>목록은 점수가 높은 순으로 정렬되어, 누가 많이 받았는지 한눈에 보입니다.</li>' +
      '<li>학급마다 따로 관리되고, <strong>전체 초기화</strong>로 새 학기에 리셋할 수 있습니다.</li>' +
      '</ul>'
  },
  {
    img: 'guide/02-roster.png',
    title: '명렬표를 입력합니다',
    desc: '<ul>' +
      '<li><em>설정 → 명렬표 수정</em>에서 <strong>학번 이름 성별</strong>을 한 줄에 한 명씩 적습니다. (예: <code>1 홍길동 남</code>)</li>' +
      '<li>성별을 넣으면 좌석표에서 남학생(파랑) · 여학생(분홍)으로 구분되어 보입니다.</li>' +
      '<li>명렬표는 <strong>이름 오름차순</strong>으로 정리됩니다. 자리의 학번순 배치와는 별개입니다.</li>' +
      '<li>담당 학급을 추가한 뒤 상단 탭으로 학급을 오갈 수 있습니다.</li>' +
      '</ul>'
  },
  {
    img: 'guide/03-excel.png',
    title: '엑셀 명렬표를 그대로 올려도 됩니다',
    desc: '<ul>' +
      '<li>파일을 끌어다 놓기만 하면 양식을 <strong>자동으로 알아봅니다</strong>.</li>' +
      '<li>지원 양식 3가지 — <strong>시트별 학급</strong> / <strong>학급 열이 있는 한 장짜리</strong> / <strong>헤더가 학급명인 피벗형</strong></li>' +
      '<li>어떻게 만들지 모르겠다면 <strong>샘플 양식 다운로드</strong>를 먼저 받아보세요.</li>' +
      '</ul>'
  },
  {
    img: 'guide/04-seating.png',
    title: '좌석표 읽는 법',
    desc: '<ul>' +
      '<li>카드마다 <strong>번호 · 이름 · 몇 분단 몇 행</strong>인지 표시됩니다.</li>' +
      '<li>카드를 <strong>끌어다 놓으면 두 학생의 자리가 서로 바뀝니다.</strong></li>' +
      '<li>학생 이름을 클릭하면 <strong>오늘 출결·칭찬·관찰·세특</strong>을 바로 기록하고 종합 정보·메모·관계·이력을 확인합니다.</li>' +
      '</ul>'
  },
  {
    section: 'seat',
    img: 'guide/v67-number-order.png',
    title: '학번순 배치 — 가로순과 세로순을 고릅니다',
    desc: '<ul>' +
      '<li>좌석 설정 패널의 <strong>① 🔢 배치 순서</strong>에서 <strong>↔ 가로순</strong> 또는 <strong>↕ 세로순</strong>을 선택합니다.</li>' +
      '<li><strong>가로순</strong> — 앞줄 왼쪽부터 오른쪽으로 채운 뒤 다음 줄로 이동합니다.</li>' +
      '<li><strong>세로순</strong> — 왼쪽 열을 앞에서 뒤로 채운 뒤 다음 열로 이동합니다.</li>' +
      '<li>해당 방향의 <strong>② 학번 가로순 배치 / 학번 세로순 배치</strong> 버튼을 누릅니다. 시험대형뿐 아니라 짝꿍·트리플·모둠에도 적용됩니다.</li>' +
      '<li>고정 좌석은 유지하고 남은 학생을 학번·번호 오름차순으로 배치합니다. 학번순에서는 관계·성별 조건보다 번호 순서가 우선입니다.</li>' +
      '</ul>',
    callout: { type: '', text: '무작위 추첨으로 돌아가려면 🎱 무작위를 고르세요. 가로·세로 방향은 학생 시점 기준이며 교탁 시점으로 바꾸어도 실제 배치는 바뀌지 않습니다.' }
  },
  {
    section: 'seat',
    img: 'guide/v67-same-gender.png',
    title: '같은 성별끼리 우선 배치합니다',
    desc: '<ul>' +
      '<li>명렬표에 <strong>남·여 성별</strong>을 입력합니다. 성별이 없으면 같은 성별 판단을 할 수 없습니다.</li>' +
      '<li><strong>① 배치 순서 → 🎱 무작위</strong>를 고르고 <strong>② 👫 같은 성별끼리 배치</strong>를 체크한 뒤 추첨합니다.</li>' +
      '<li>짝·분단·모둠 안에서 같은 성별을 우선합니다. 인원 비율·빈자리·고정 좌석에 따라 완전한 분리는 보장되지 않습니다.</li>' +
      '<li>성별 우선 배치에서는 관계 조건이 모두 보장되지 않으며 <strong>멀리 배치 조건은 별도로 적용하지 않습니다.</strong></li>' +
      '</ul>',
    callout: { type: '', text: '학번 가로순·세로순에서는 성별 옵션이 비활성화됩니다. 조건이 중요한 학생의 자리는 결과를 확인하고 직접 이동·고정해 주세요.' }
  },
  {
    section: 'attend',
    img: 'guide/v67-seat-attendance.png',
    title: '자리에서 학생 이름 클릭 — 오늘 출결 바로 등록',
    desc: '<ul>' +
      '<li>자리배치 화면에서 <strong>학생 이름</strong>을 누르면 학생 기록 팝업이 열립니다.</li>' +
      '<li><strong>① ✅ 오늘 출결 바로 등록</strong>에서 출석·결석·지각·조퇴·결과를 선택합니다.</li>' +
      '<li>필요하면 <strong>사유</strong>(질병·인정·미인정·기타)와 짧은 <strong>메모</strong>를 입력합니다.</li>' +
      '<li>선택·입력은 <strong>즉시 저장</strong>됩니다. 팝업을 닫고 출결관리에서 같은 학급·오늘 날짜의 상태를 확인합니다.</li>' +
      '<li>잘못 선택했다면 출석 등 올바른 상태로 다시 누릅니다. 다른 날짜 기록은 출결관리에서 날짜를 바꿔 수정합니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '자리에서 입력한 출결과 출결관리 화면은 같은 학생·학급 기록을 사용합니다. 공식 학교 출결 시스템으로 자동 전송되는 기능은 아닙니다.' }
  },
  {
    section: 'seat',
    img: 'guide/v67-teacher-view.png',
    title: '교탁 시점 — 출력하지 않고 화면을 뒤집습니다',
    desc: '<ul>' +
      '<li>좌석 설정 패널의 <strong>👩‍🏫 교탁 시점</strong>을 누르면 교탁에서 학생들을 마주 보는 방향으로 화면을 뒤집습니다.</li>' +
      '<li>앞뒤·좌우 표시 방향을 바꾸는 기능이며 실제 자리·학생 기록을 옮기지 않습니다.</li>' +
      '<li>뒤집힌 화면에서도 학생 이름을 눌러 출결·칭찬·관찰 기록을 등록할 수 있습니다.</li>' +
      '<li><strong>🧑‍🎓 학생 시점으로</strong>를 누르면 원래 화면으로 돌아옵니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '수업 중 확인은 교탁 시점, 종이 배부는 출력의 학생용·교탁용을 이용하세요.' }
  },
  {
    img: 'guide/05a-splash.png',
    title: '추첨 시작 — 3 · 2 · 1',
    desc: '<p><strong>🎱 추첨 시작</strong>을 누르면 화면 전체가 추첨 무대로 바뀝니다.</p>' +
      '<ul>' +
      '<li>학급명과 인원이 크게 뜨고 <strong>3 · 2 · 1 카운트다운</strong>이 진행됩니다.</li>' +
      '<li>숫자마다 효과음이 나고, 마지막에 색종이가 터지며 추첨이 시작됩니다.</li>' +
      '<li>연출이 길게 느껴지면 오른쪽 아래 <strong>건너뛰기</strong>를 누르세요. (Esc·Enter도 됩니다)</li>' +
      '</ul>',
    callout: { type: 'good', text: '소리는 앱이 직접 만들어 내므로 <strong>버튼을 누르는 즉시</strong> 울립니다. 음원을 불러오느라 늦어지지 않습니다.' }
  },
  {
    img: 'guide/05-lottery.png',
    title: '추첨을 시작합니다',
    desc: '<ul>' +
      '<li><strong>추첨 시작</strong>을 누르면 <strong>3 · 2 · 1 카운트다운</strong> 뒤에 추첨이 시작됩니다.</li>' +
      '<li>공이 하나씩 뽑혀 자리로 날아가고, 확정될 때마다 색종이가 터집니다.</li>' +
      '<li><strong>🐢 천천히 / ▶ 보통 / ⏩ 빠르게</strong>로 진행 속도를 바꿀 수 있습니다.</li>' +
      '<li><strong>마지막 3명</strong>은 자동으로 느려지면서 긴장감이 올라갑니다.</li>' +
      '<li>소리가 부담스러운 교실이라면 <strong>🔇 버튼</strong>으로 끄면 됩니다.</li>' +
      '<li>바로 결과만 보려면 <strong>건너뛰기</strong>를 누르세요.</li>' +
      '</ul>'
  },
  {
    img: 'guide/06-memo.png',
    title: '학생별 메모를 남깁니다',
    desc: '<ul>' +
      '<li>좌석 카드를 클릭하고 <strong>📝 메모</strong> 탭에 적으면 <strong>자동 저장</strong>됩니다.</li>' +
      '<li>시력 · 청력 · 키 · 배려사항처럼 자리를 정할 때 참고할 내용을 적어두세요.</li>' +
      '<li>메모는 뒤에서 설명할 <strong>누가기록 엑셀</strong>에도 같이 나옵니다.</li>' +
      '</ul>'
  },
  {
    img: 'guide/07-relations.png',
    title: '붙일 학생, 떼어놓을 학생을 지정합니다',
    desc: '<ul>' +
      '<li><strong>💚 가까이 배치</strong> — 되도록 옆·앞뒤로 붙여 앉힙니다.</li>' +
      '<li><strong>💔 멀리 배치</strong> — 서로 인접하지 않도록 떨어뜨립니다.</li>' +
      '<li><strong>무작위 배치에서 성별 우선 옵션을 끈 경우</strong> 관계를 반영하고 나머지를 무작위로 채웁니다.</li>' +
      '</ul>',
    callout: { type: '', text: '관계 설정이 너무 많으면 조건을 모두 만족할 수 없어 일부는 무작위로 배치됩니다.' }
  },
  {
    img: 'guide/08-lock.png',
    title: '특정 자리를 고정합니다',
    desc: '<ul>' +
      '<li>카드의 <strong>🔒</strong>를 누르면 그 학생은 다시 추첨해도 <strong>자리가 유지</strong>됩니다.</li>' +
      '<li>빈자리도 고정할 수 있어, 늘 비워둘 자리를 정해둘 수 있습니다.</li>' +
      '<li>시력이 나쁜 학생을 앞자리에 두거나, 교탁 옆을 비워둘 때 유용합니다.</li>' +
      '</ul>'
  },
  {
    img: 'guide/v67-history-load.png',
    title: '배치를 저장하고 언제든 되돌립니다',
    desc: '<ul>' +
      '<li><strong>💾 저장</strong>으로 지금 배치를 메모와 함께 기록해 둡니다. (예: “2학기 1차 배치”)</li>' +
      '<li><strong>📂 배치 불러오기</strong>에서 기록을 골라 <strong>① 불러오기</strong>를 누르면 좌석·배치 설정·고정 좌석을 복원합니다.</li>' +
      '<li>복원 전에 현재 배치도 저장해 두세요. 학생을 삭제·변경했다면 복원 결과를 다시 확인합니다.</li>' +
      '</ul>',
    callout: { type: 'danger', text: '저장 버튼을 눌러야 기록에 남습니다. 추첨만 하고 넘어가면 이력이 쌓이지 않습니다.' }
  },
  {
    img: 'guide/10-download.png',
    title: '학생 누가기록을 엑셀로 내려받습니다',
    desc: '<p>이 프로그램의 핵심 기능입니다. <strong>📊 이력 다운로드</strong>를 누르면 지금까지 저장한 모든 배치가 <strong>학생 한 명당 한 줄</strong>로 정리된 엑셀 파일이 만들어집니다.</p>' +
      '<ul>' +
      '<li>학번 · 이름 · 성별 · <strong>메모</strong></li>' +
      '<li>저장한 배치마다 그 학생이 <strong>몇 분단 몇 행</strong>에 앉았는지</li>' +
      '<li>그때 <strong>양옆 짝이 누구였는지</strong></li>' +
      '</ul>' +
      '<p>학기말 상담자료나 생활기록 참고자료로 그대로 쓸 수 있습니다.</p>',
    callout: { type: 'good', text: '자리를 바꿀 때마다 저장해 두면, 한 학기 치 짝 이력이 자동으로 누적됩니다.' }
  },
  {
    img: 'guide/11-print.png',
    title: '좌석표를 인쇄합니다',
    desc: '<ul>' +
      '<li><strong>🖨️ 학생용</strong> — 학생이 보는 방향 그대로. 교실 게시용입니다.</li>' +
      '<li><strong>🖨️ 교탁용</strong> — 교탁에서 학생들을 마주 보는 방향의 좌석표입니다.</li>' +
      '<li>종이 출력 없이 확인할 때는 좌석 설정 패널의 <strong>👩‍🏫 교탁 시점</strong>을 사용합니다.</li>' +
      '<li>인쇄 창에서 <em>배경 그래픽</em>을 켜면 남녀 구분 색이 그대로 인쇄됩니다.</li>' +
      '</ul>'
  },
  {
    img: 'guide/21A-duty-excel.png',
    title: '1인 1역 — 엑셀 배정도 바로 연결됩니다',
    desc: '<ul>' +
      '<li><strong>엑셀 양식 받기</strong>를 누르면 두 시트가 만들어집니다.</li>' +
      '<li><strong>① 직접배정</strong> — 학번·이름·역할을 적어 학생별 역할을 그대로 불러옵니다.</li>' +
      '<li><strong>② 역할목록</strong> — 역할만 적어 올린 뒤 무작위 배정을 누릅니다.</li>' +
      '<li>불러온 역할은 <strong>좌석표·학생 종합 카드·통합 기록</strong>에 즉시 연결됩니다.</li>' +
      '</ul>',
    callout: { type: 'good', text: '기존에 쓰던 1인 1역 엑셀이 있다면 양식의 열 이름만 맞춰 그대로 활용할 수 있습니다.' }
  },  {
    section: 'settings',
    img: 'guide/v67-data-backup.png',
    title: '홈에서 데이터 백업 — 기기를 바꾸기 전에 꼭 보관하세요',
    desc: '<p><strong>오늘 방문자 수와 총 접속 횟수만 익명 통계로 집계</strong>합니다. 명렬표 · 출결 · 좌석 · 메모 · 상담 기록은 T-ong 서버에 저장하지 않고 <strong>지금 쓰고 있는 브라우저 안에만</strong> 저장됩니다. Google 동기화를 직접 켠 경우 선택한 일정만 본인 Google 계정에 저장됩니다.</p>' +
      '<ul>' +
      '<li><strong>홈 ▾ → 📦 데이터 백업 → ① 코드 생성 → ② 복사</strong>를 눌러 백업 코드를 안전한 곳에 보관합니다. 기존 설정 메뉴에서도 열 수 있습니다.</li>' +
      '<li>학교 컴퓨터에서 만든 학생 자료는 <strong>집 컴퓨터·태블릿에서 자동으로 보이지 않습니다.</strong></li>' +
      '<li>같은 PC라도 <strong>Chrome과 Edge는 서로 다른 저장소</strong>를 씁니다.</li>' +
      '<li><strong>운영 주소와 로컬 확인 주소</strong>도 각각 다른 저장소입니다. 같은 4321 주소로 접속해도 기기별 학생 자료는 별개입니다.</li>' +
      '<li>새 기기에서는 같은 운영 주소를 열고 코드를 붙여넣은 뒤 <strong>③ 가져오기</strong>를 누릅니다. 현재 자료를 대체하므로 <strong>대상 기기의 자료도 먼저 백업</strong>합니다.</li>' +
      '<li>브라우저에서 <em>쿠키 및 사이트 데이터 삭제</em>를 하면 <strong>모두 지워집니다.</strong></li>' +
      '<li>시크릿(비공개) 창에서 쓴 내용은 창을 닫으면 사라집니다.</li>' +
      '</ul>' +
      '<p>처음 접속해 필요한 파일이 캐시된 뒤에는 핵심 로컬 기능을 오프라인으로 사용할 수 있습니다. Google 동기화·외부 사이트·학교 정보 조회는 인터넷이 필요합니다.</p>',
    callout: { type: 'danger', text: '백업 코드는 암호화된 보관함이 아니며 학생 기록을 포함합니다. 공개 블로그·메신저 단체방에 올리지 말고 학교의 자료 보관·개인정보 처리 기준에 따라 안전하게 관리하세요.' }
  }
];

var GUIDE_SECTIONS = [
  { key: 'start', label: '시작' },
  { key: 'seat', label: '자리배치' },
  { key: 'attend', label: '출결관리' },
  { key: 'classroom', label: '교실운영' },
  { key: 'calendar', label: '일정관리' },
  { key: 'favorites', label: '즐겨찾기' },
  { key: 'games', label: '게임' },
  { key: 'help', label: '도움말' },
  { key: 'settings', label: '설정·데이터' }
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
    el.className = 'guide-slide';

    var callout = '';
    if (slide.callout) {
      callout = '<div class="guide-callout ' + (slide.callout.type || '') + '">💡 ' + slide.callout.text + '</div>';
    }

    el.innerHTML =
      '<div class="guide-step">' + guideSectionLabel(slide) + ' · STEP ' + (i + 1) + ' / ' + GUIDE_SLIDES.length + '</div>' +
      '<h4>' + slide.title + '</h4>' +
      '<img class="guide-shot" src="' + slide.img + '?v=' + encodeURIComponent(APP_VERSION) + '" alt="' + escapeHtml(slide.title) + ' 화면" loading="lazy">' +
      '<div class="guide-side">' +
        '<div class="guide-desc">' + slide.desc + '</div>' +
        callout +
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
// 도움말 실험 기능 — NEIS 오늘의 학교
// ===================================================================
var SchoolToday = (function () {
  var PROFILE_KEY='tongSchoolProfile_v1';
  var CACHE_KEY='tongSchoolTodayCache_v1';
  var profile=null;
  var selectedSchool=null;
  var requestNumber=0;

  function el(id){return document.getElementById(id);}
  function readJson(key,fallback){try{var value=JSON.parse(localStorage.getItem(key)||'null');return value==null?fallback:value;}catch(e){return fallback;}}
  function writeJson(key,value){try{localStorage.setItem(key,JSON.stringify(value));}catch(e){}}
  function compactDate(key){return String(key||'').replace(/-/g,'');}
  function weekdayText(key){var d=new Date(String(key)+'T00:00:00');return ['일','월','화','수','목','금','토'][d.getDay()]+'요일';}
  function timetableDataset(kind){if(kind==='초등학교')return'elsTimetable';if(kind==='중학교')return'misTimetable';if(kind==='특수학교')return'spsTimetable';return'hisTimetable';}
  function semesterFor(key){var d=new Date(String(key)+'T00:00:00'),month=d.getMonth()+1,year=d.getFullYear();if(month<=2)return{year:year-1,semester:2};return{year:year,semester:month<=8?1:2};}
  function apiRows(data,setName){
    if(data&&data.RESULT&&data.RESULT.CODE==='INFO-200')return[];
    var set=data&&data[setName];if(!Array.isArray(set))return[];
    var block=set.find(function(part){return part&&Array.isArray(part.row);});return block?block.row:[];
  }
  async function request(setName,params){
    var url=new URL('https://open.neis.go.kr/hub/'+setName);url.searchParams.set('Type','json');url.searchParams.set('pIndex','1');url.searchParams.set('pSize',params.pSize||'100');
    Object.keys(params).forEach(function(key){if(key!=='pSize'&&params[key]!==''&&params[key]!=null)url.searchParams.set(key,params[key]);});
    var response=await fetch(url.toString(),{cache:'no-store'});if(!response.ok)throw new Error('NEIS '+response.status);return apiRows(await response.json(),setName);
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
  function saveCache(date,item){var all=readJson(CACHE_KEY,{});all[cacheKey(date)]={savedAt:Date.now(),timetable:item.timetable,meal:item.meal};var keys=Object.keys(all).sort(function(a,b){return Number(all[b].savedAt)-Number(all[a].savedAt);});keys.slice(20).forEach(function(key){delete all[key];});writeJson(CACHE_KEY,all);}
  function renderData(date,data,cached){el('schoolTimetableMeta').textContent=profile.grade+'학년 '+profile.className+'반';renderTimetable(data.timetable||[]);renderMeal(data.meal||[]);setStatus((Number(date.slice(5,7)))+'월 '+Number(date.slice(8,10))+'일 '+weekdayText(date)+' · NEIS 공개정보'+(cached?' · 저장된 조회 결과':''),cached?'cached':'ok');}
  async function refresh(force){
    if(!profile){setSettings(true);setStatus('학교를 먼저 설정해 주세요.','error');return;}
    var date=el('schoolTodayDate').value||formatDate(new Date()),cached=!force&&readCache(date);if(cached){renderData(date,cached,true);return;}
    var ownRequest=++requestNumber;loading();setStatus('NEIS 공개정보를 조회하고 있습니다.');
    try{var semester=semesterFor(date),common={ATPT_OFCDC_SC_CODE:profile.atptCode,SD_SCHUL_CODE:profile.schoolCode};var tableParams=Object.assign({},common,{AY:String(semester.year),SEM:String(semester.semester),ALL_TI_YMD:compactDate(date),GRADE:String(profile.grade),CLASS_NM:String(profile.className),pSize:'100'});var mealParams=Object.assign({},common,{MLSV_YMD:compactDate(date),pSize:'10'});var values=await Promise.all([request(timetableDataset(profile.schoolKind),tableParams),request('mealServiceDietInfo',mealParams)]);if(ownRequest!==requestNumber)return;var data={timetable:values[0],meal:values[1]};saveCache(date,data);renderData(date,data,false);}catch(error){if(ownRequest!==requestNumber)return;el('schoolTimetableList').innerHTML=emptyHtml('시간표를 불러오지 못했습니다.','잠시 후 새로고침해 주세요.');el('schoolMealList').innerHTML=emptyHtml('급식을 불러오지 못했습니다.','네트워크 연결을 확인해 주세요.');setStatus('NEIS 조회에 실패했습니다. 잠시 후 다시 시도해 주세요.','error');}
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
  var SCOPES='openid email https://www.googleapis.com/auth/calendar.events.owned';
  var CALENDAR_ID='primary',SYNC_MODE='primary-v1';
  var API='https://www.googleapis.com/calendar/v3',USERINFO='https://openidconnect.googleapis.com/v1/userinfo';
  var token='',expiresAt=0,tokenClient=null,loadingPromise=null,syncing=false,timer=null,debounceTimer=null;
  var chooseAccount=false,expiryNotified=false; // 연결 해제 후에는 계정 선택 창부터, 만료 안내는 한 번만
  var config={clientId:'',activeAccountId:'',autoSync:true,includeConsult:false,includeMemo:false};
  var accountStates={};
  function el(id){return document.getElementById(id);}
  function parse(value,fallback){try{var data=JSON.parse(value);return data&&typeof data==='object'?data:fallback;}catch(e){return fallback;}}
  function metaClientId(){var node=document.querySelector('meta[name="google-oauth-client-id"]');return node?String(node.content||'').trim():'';}
  function load(){
    var stored={};try{stored=parse(localStorage.getItem(CONFIG_KEY)||'{}',{});accountStates=parse(localStorage.getItem(STATE_KEY)||'{}',{});}catch(e){}
    config.clientId=String(metaClientId()||stored.clientId||'').trim();config.activeAccountId=String(stored.activeAccountId||'');
    config.autoSync=stored.autoSync!==false;config.includeConsult=stored.includeConsult===true;config.includeMemo=stored.includeMemo===true;
  }
  function saveConfig(){try{localStorage.setItem(CONFIG_KEY,JSON.stringify(config));}catch(e){}}
  function saveStates(){try{localStorage.setItem(STATE_KEY,JSON.stringify(accountStates));}catch(e){}}
  function activeState(){return config.activeAccountId&&accountStates[config.activeAccountId]||null;}
  function ensureState(user){var id=String(user.sub||user.email||'');var state=accountStates[id]||{};state.accountId=id;state.email=user.email||state.email||'Google 계정';state.links=state.links||{};state.pendingDeletes=Array.isArray(state.pendingDeletes)?state.pendingDeletes:[];accountStates[id]=state;config.activeAccountId=id;saveConfig();saveStates();return state;}
  function validToken(){return !!token&&Date.now()<expiresAt-30000;}
  function setBusy(value){syncing=value;var button=el('googleCalSyncBtn');if(button){button.disabled=!validToken()||value;button.textContent=value?'동기화 중…':'지금 양방향 동기화';}var full=el('googleCalFullSyncBtn');if(full)full.disabled=!validToken()||value;var connect=el('googleCalConnectBtn');if(connect)connect.disabled=value;}
  function setStatus(text,kind,detail){var status=el('googleCalStatus'),badge=el('googleCalStateBadge'),summary=el('googleCalSummaryStatus');if(status)status.textContent=text;if(badge){badge.textContent=kind==='ok'?'연결됨':kind==='busy'?'동기화 중':kind==='error'?'확인 필요':'대기';badge.dataset.state=kind||'idle';}if(summary)summary.textContent=kind==='ok'?'연결됨':kind==='busy'?'동기화 중':kind==='error'?'확인 필요':'연결 안 됨';if(detail&&el('googleCalAccount'))el('googleCalAccount').textContent=detail;}
  function render(){
    var state=activeState(),connected=validToken();if(el('googleClientIdInput'))el('googleClientIdInput').value=config.clientId||'';
    if(el('googleCalAutoInput'))el('googleCalAutoInput').checked=config.autoSync;if(el('googleCalConsultInput'))el('googleCalConsultInput').checked=config.includeConsult;if(el('googleCalMemoInput'))el('googleCalMemoInput').checked=config.includeMemo;
    if(el('googleCalConnectBtn'))el('googleCalConnectBtn').textContent=connected?'다른 Google 계정 연결':(quickHint()?'다시 연결 · '+quickHint():'Google 계정 연결');
    var dev=el('googleCalDeveloperSettings');if(dev)dev.hidden=!!metaClientId()&&!/[?&]admin\b/.test(location.search);if(el('googleCalSyncBtn'))el('googleCalSyncBtn').disabled=!connected||syncing;if(el('googleCalFullSyncBtn'))el('googleCalFullSyncBtn').disabled=!connected||syncing;if(el('googleCalDisconnectBtn'))el('googleCalDisconnectBtn').disabled=!connected;
    if(connected&&state)setStatus('Google 캘린더와 연결되었습니다.','ok',(state.email||'Google 계정')+' · '+(state.calendarName||'Google 기본 캘린더'));
    else if(!config.clientId)setStatus('앱 관리자 설정이 필요합니다.','error','OAuth 클라이언트 ID를 입력한 뒤 계정을 연결해 주세요.');
    else setStatus('Google 계정을 연결해 주세요.','idle',state&&state.email?'최근 계정 · '+state.email:'로그인 계정의 Google 기본 캘린더와 양방향 동기화합니다.');
    var last=el('googleCalLastSync');if(last)last.textContent='마지막 동기화: '+(state&&state.lastSyncAt?new Date(state.lastSyncAt).toLocaleString('ko-KR'):'없음');
  }
  function loadIdentity(){
    if(window.google&&google.accounts&&google.accounts.oauth2)return Promise.resolve();if(loadingPromise)return loadingPromise;
    loadingPromise=new Promise(function(resolve,reject){var old=document.querySelector('script[data-tong-google-identity]');if(old){old.addEventListener('load',resolve,{once:true});old.addEventListener('error',function(){reject(new Error('Google 로그인 모듈을 불러오지 못했습니다.'));},{once:true});return;}var script=document.createElement('script');script.src='https://accounts.google.com/gsi/client';script.async=true;script.defer=true;script.dataset.tongGoogleIdentity='1';script.onload=resolve;script.onerror=function(){reject(new Error('Google 로그인 모듈을 불러오지 못했습니다.'));};document.head.appendChild(script);});return loadingPromise;
  }
  function initTokenClient(){tokenClient=google.accounts.oauth2.initTokenClient({client_id:config.clientId,scope:SCOPES,callback:function(response){if(response&&response.error){setStatus('Google 연결을 완료하지 못했습니다.','error',response.error_description||response.error);setBusy(false);return;}token=response.access_token||'';expiresAt=Date.now()+(Number(response.expires_in)||3600)*1000;chooseAccount=false;expiryNotified=false;finishConnection().catch(handleError);}});}
  // 예전에 연결한 계정이 있으면 계정 선택·동의 화면 없이 한 번에 다시 연결합니다.
  function quickHint(){var state=activeState();return !chooseAccount&&state&&state.email&&state.email.indexOf('@')>0?state.email:'';}
  function connect(){if(!config.clientId){var detail=el('googleCalDeveloperSettings');if(detail)detail.open=true;render();showToast('Google OAuth 클라이언트 ID를 먼저 설정해 주세요','error');return;}setBusy(true);setStatus('Google 로그인 창을 준비하고 있습니다.','busy','연결할 계정을 선택해 주세요.');var hint=validToken()?'':quickHint();loadIdentity().then(function(){initTokenClient();tokenClient.requestAccessToken(hint?{prompt:'',login_hint:hint}:{prompt:'select_account'});}).catch(handleError);}
  async function api(pathname,options){
    options=options||{};var headers=Object.assign({'Authorization':'Bearer '+token},options.headers||{});if(options.body)headers['Content-Type']='application/json';var response=await fetch(pathname.indexOf('http')===0?pathname:API+pathname,Object.assign({},options,{headers:headers}));
    if(response.status===204)return null;var body=null;try{body=await response.json();}catch(e){}if(!response.ok){var err=new Error(body&&body.error&&body.error.message||('Google API 오류 '+response.status));err.status=response.status;throw err;}return body;
  }
  async function finishConnection(){var user=await api(USERINFO),state=ensureState(user);await ensureCalendar(state);setBusy(false);render();await syncNow(true);startTimer();}
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
  function init(){load();render();var connectButton=el('googleCalConnectBtn'),syncButton=el('googleCalSyncBtn'),fullSyncButton=el('googleCalFullSyncBtn'),disconnectButton=el('googleCalDisconnectBtn'),saveButton=el('googleClientIdSaveBtn');if(connectButton)connectButton.addEventListener('click',connect);if(syncButton)syncButton.addEventListener('click',function(){syncNow(false);});if(fullSyncButton)fullSyncButton.addEventListener('click',fullSync);if(disconnectButton)disconnectButton.addEventListener('click',disconnect);if(saveButton)saveButton.addEventListener('click',saveClientId);var helpButton=el('googleOAuthHelpBtn');if(helpButton)helpButton.addEventListener('click',function(e){e.preventDefault();e.stopPropagation();openModal('googleOAuthHelpModal');});['googleCalAutoInput','googleCalConsultInput','googleCalMemoInput'].forEach(function(id){var node=el(id);if(node)node.addEventListener('change',optionChanged);});document.addEventListener('visibilitychange',function(){if(document.visibilityState==='visible'&&validToken()&&config.autoSync)syncNow(true);});}
  return{init:init,openPanel:openPanel,onLocalChanged:onLocalChanged,noteLocalDelete:noteLocalDelete,syncNow:syncNow};
})();
// ===================================================================
// AI 도우미 — 선생님 본인의 구독 계정(Claude·ChatGPT·Gemini)으로 요청을 넘깁니다.
// API 키를 쓰지 않으며, 요청문·답변은 T-ong 서버로 보내지 않습니다.
// ===================================================================
var AiAssistant = (function () {
  var PREF_KEY = 'tongAiAssistant_v1';
  var MAX_URL_PROMPT = 1800;
  var PROVIDERS = {
    claude: { name: 'Claude', url: function (q) { return 'https://claude.ai/new' + (q ? '?q=' + encodeURIComponent(q) : ''); } },
    chatgpt: { name: 'ChatGPT', url: function (q) { return 'https://chatgpt.com/' + (q ? '?q=' + encodeURIComponent(q) : ''); } },
    gemini: { name: 'Gemini', url: function () { return 'https://gemini.google.com/app'; } }
  };
  var TASKS = {
    notice: { placeholder: '예) 내일 현장체험학습 알림장. 준비물은 도시락, 물, 편한 운동화', prefix: '초등·중등 담임교사가 학생과 학부모에게 보내는 짧은 알림장을 써 주세요. 번호 목록으로 3~6줄, 다정하고 간결한 말투로, 설명 없이 알림장 본문만 출력하세요.\n\n요청: ' },
    morning: { placeholder: '예) 시험 앞두고 긴장한 아이들에게 힘이 되는 이야기', prefix: '담임교사가 아침 조회 시간에 학생들에게 1~2분 동안 들려줄 멘트를 써 주세요. 따뜻하고 구체적인 예시 하나를 넣어 주세요.\n\n주제: ' },
    letter: { placeholder: '예) 다음 주 학부모 상담주간 안내 (기간, 신청 방법)', prefix: '학교에서 학부모에게 보내는 가정통신문을 써 주세요. 제목, 인사, 본문, 맺음말 순서로 공손한 공문 말투를 사용하세요.\n\n내용: ' },
    consult: { placeholder: '상담 메모를 붙여넣으세요 (학생 이름은 자동으로 가려집니다)', prefix: '아래 학생 상담 메모를 교사용 상담 기록 양식으로 정리해 주세요. 항목: 상담 일시·방법 / 주요 내용 / 학생 상황 판단 / 후속 조치. 사실만 간결하게 쓰고 추측은 표시해 주세요.\n\n메모:\n' },
    free: { placeholder: '무엇이든 물어보세요', prefix: '' }
  };
  var pref = { provider: 'claude', task: 'notice', mask: true };
  var nameMap = {};   // '[학생A]' → 실제 이름

  function el(id) { return document.getElementById(id); }
  function loadPref() { try { var v = JSON.parse(localStorage.getItem(PREF_KEY) || '{}'); if (PROVIDERS[v.provider]) pref.provider = v.provider; if (TASKS[v.task]) pref.task = v.task; if (v.mask === false) pref.mask = false; } catch (e) {} }
  function savePref() { try { localStorage.setItem(PREF_KEY, JSON.stringify(pref)); } catch (e) {} }

  function render() {
    Array.prototype.forEach.call(document.querySelectorAll('#aiProviderSeg [data-provider]'), function (b) { b.classList.toggle('active', b.dataset.provider === pref.provider); });
    Array.prototype.forEach.call(document.querySelectorAll('#aiTaskChips [data-task]'), function (b) { b.classList.toggle('active', b.dataset.task === pref.task); });
    var input = el('aiRequestInput'); if (input) input.placeholder = TASKS[pref.task].placeholder;
    var mask = el('aiMaskInput'); if (mask) mask.checked = pref.mask;
    var send = el('aiSendBtn'); if (send) send.textContent = PROVIDERS[pref.provider].name + ' 구독 계정으로 열기 ↗';
  }

  function open() { render(); el('aiDrawer').classList.add('show'); el('aiDrawer').setAttribute('aria-hidden', 'false'); el('aiDrawerBackdrop').classList.add('show'); setTimeout(function () { var i = el('aiRequestInput'); if (i) i.focus(); }, 60); }
  function close() { el('aiDrawer').classList.remove('show'); el('aiDrawer').setAttribute('aria-hidden', 'true'); el('aiDrawerBackdrop').classList.remove('show'); }

  // 명렬표 이름을 [학생A], [학생B] … 로 바꿉니다. 긴 이름부터 바꿔 '김민'이 '김민준'을 먼저 먹지 않게 합니다.
  function maskNames(text) {
    nameMap = {};
    var names = (typeof students !== 'undefined' && Array.isArray(students) ? students : [])
      .map(function (s) { return String(s && s.name || '').trim(); })
      .filter(function (n, i, arr) { return n.length >= 2 && !/^학생\d+$/.test(n) && arr.indexOf(n) === i; })
      .sort(function (a, b) { return b.length - a.length; });
    var count = 0;
    names.forEach(function (name) {
      if (text.indexOf(name) === -1) return;
      var label = '[학생' + (count < 26 ? String.fromCharCode(65 + count) : (count + 1)) + ']';
      count++; nameMap[label] = name; text = text.split(name).join(label);
    });
    return text;
  }
  function unmaskNames(text) { Object.keys(nameMap).forEach(function (label) { text = text.split(label).join(nameMap[label]); }); return text; }

  function copyText(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) return navigator.clipboard.writeText(text).then(function () { return true; }, function () { return false; });
    return Promise.resolve(false);
  }

  function send() {
    var request = String(el('aiRequestInput').value || '').trim();
    if (!request) { showToast('요청 내용을 적어 주세요', 'error'); el('aiRequestInput').focus(); return; }
    var prompt = TASKS[pref.task].prefix + request;
    if (pref.mask) prompt = maskNames(prompt); else nameMap = {};
    var provider = PROVIDERS[pref.provider];
    var inUrl = pref.provider !== 'gemini' && prompt.length <= MAX_URL_PROMPT;
    // 팝업 차단을 피하려고 클릭 직후 바로 창을 엽니다. 클립보드 복사는 그다음에 진행합니다.
    var win = window.open(provider.url(inUrl ? prompt : ''), '_blank');
    if (win) { try { win.opener = null; } catch (e) {} }
    copyText(prompt).then(function (copied) {
      var masked = Object.keys(nameMap).length;
      var msg = provider.name + '을(를) 열었습니다' + (inUrl ? '' : copied ? ' · 입력칸에 붙여넣기 하세요' : '') + (masked ? ' · 이름 ' + masked + '개 가림' : '');
      showToast(msg);
      if (!inUrl && !copied) showToast('클립보드 복사에 실패했습니다. 요청 내용을 직접 복사해 주세요', 'error');
    });
    if (!win) showToast('새 창이 차단되었습니다. 팝업 허용 후 다시 눌러 주세요', 'error');
  }

  function paste() {
    if (!navigator.clipboard || !navigator.clipboard.readText) { showToast('입력칸을 길게 눌러 붙여넣기 해 주세요'); el('aiAnswerInput').focus(); return; }
    navigator.clipboard.readText().then(function (text) { el('aiAnswerInput').value = text || ''; if (!text) showToast('클립보드가 비어 있습니다', 'error'); },
      function () { showToast('입력칸을 길게 눌러 붙여넣기 해 주세요'); el('aiAnswerInput').focus(); });
  }

  function toNotice() {
    var text = unmaskNames(String(el('aiAnswerInput').value || '').trim());
    if (!text) { showToast('먼저 AI 답변을 붙여넣어 주세요', 'error'); return; }
    var cls = allClassData[currentClassIndex]; if (!cls) return;
    if ((cls.notice || '').trim() && !confirm('이미 작성한 알림장이 있습니다. AI 답변으로 바꿀까요?')) return;
    cls.notice = text; cls.noticeAt = formatDateTime(new Date());
    saveToLocalStorage(); Notice.renderPreview(); close();
    showToast('알림장에 넣었습니다', '', '열기', function () { var b = el('noticeBtn'); if (b) b.click(); });
  }

  function init() {
    if (!el('aiDrawer')) return;
    loadPref(); render();
    Array.prototype.forEach.call(document.querySelectorAll('[data-ai-open]'), function (b) { b.addEventListener('click', open); });
    el('aiDrawerClose').addEventListener('click', close);
    el('aiDrawerBackdrop').addEventListener('click', close);
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && el('aiDrawer').classList.contains('show')) close(); });
    el('aiProviderSeg').addEventListener('click', function (e) { var b = e.target.closest('[data-provider]'); if (!b) return; pref.provider = b.dataset.provider; savePref(); render(); });
    el('aiTaskChips').addEventListener('click', function (e) { var b = e.target.closest('[data-task]'); if (!b) return; pref.task = b.dataset.task; savePref(); render(); });
    el('aiMaskInput').addEventListener('change', function () { pref.mask = this.checked; savePref(); });
    el('aiSendBtn').addEventListener('click', send);
    el('aiPasteBtn').addEventListener('click', paste);
    el('aiToNoticeBtn').addEventListener('click', toNotice);
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
  document.querySelectorAll('.hero-tag').forEach(function (t) {
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

function initModeTabs() {
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


})();
