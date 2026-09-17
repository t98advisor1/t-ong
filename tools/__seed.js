/* 스크린샷 촬영 전용 시드 스크립트 — 배포에 포함되지 않습니다(.vercelignore).
   등장 인물은 모두 가상의 이름입니다. */
(function () {
  var params = new URLSearchParams(location.search);
  var shot = params.get('s') || '04';

  try { localStorage.clear(); } catch (e) { return; }
  localStorage.setItem('seatGuide_v2_seen', '1');   // 촬영 중 안내창 자동 열림 방지
  if (shot === '23') localStorage.setItem('seatCustomFavorites_v1', JSON.stringify([{id:'fav-1',name:'우리 학교 홈페이지',url:'https://school.example.com/',kind:'general'},{id:'fav-2',name:'수업 자료 공유방',url:'https://drive.google.com/',kind:'general'},{id:'book-1',name:'우리 교과 자료실',url:'https://textbook.example.com/',kind:'textbook'}]));
  if (shot === '26') localStorage.setItem('tongSchoolProfile_v1', JSON.stringify({atptCode:'B10',schoolCode:'7010083',schoolName:'서울고등학교',schoolKind:'고등학교',location:'서울특별시',address:'서울특별시 서초구 효령로 197',grade:'1',className:'1'}));  if (shot === '22') localStorage.setItem('seatCalendar_v1', JSON.stringify([{id:'cal-1',title:'개학식',start:'2026-08-17',end:'2026-08-17',property:'회의',target:'전체',memo:'2학기 시작'},{id:'cal-2',title:'학생 상담',start:'2026-08-20',end:'2026-08-20',property:'상담',target:'학생',studentId:'3',studentName:'3번 김도윤',consultMethod:'학부모 상담',status:'후속 필요',followUp:'다음 주 확인'}]));

  if (shot === '01') return;                        // 설정 위자드 화면은 빈 상태로 촬영

  var NAMES = [
    ['강민준', '남'], ['고은서', '여'], ['김도윤', '남'], ['김서아', '여'],
    ['김지호', '남'], ['김하은', '여'], ['나윤우', '남'], ['노예린', '여'],
    ['문시우', '남'], ['박다인', '여'], ['박준서', '남'], ['배소율', '여'],
    ['서지훈', '남'], ['손채원', '여'], ['신건우', '남'], ['안유진', '여'],
    ['오현우', '남'], ['유서연', '여'], ['윤태오', '남'], ['이나윤', '여'],
    ['이준혁', '남'], ['임수아', '여'], ['장민재', '남'], ['전예은', '여'],
    ['정우진', '남'], ['조하린', '여'], ['최시윤', '남'], ['한지우', '여']
  ];

  var students = NAMES.map(function (n, i) {
    return { id: String(i + 1), name: n[0], gender: n[1] };
  });

  var ROWS = 5, COLS = 6;
  var seatIds = [];
  for (var r = 0; r < ROWS; r++) {
    for (var c = 0; c < COLS; c++) seatIds.push(r + '-' + c);
  }

  // 고정 배열(촬영 결과를 매번 동일하게)
  var ORDER = [12, 3, 25, 7, 18, 0, 21, 9, 14, 2, 27, 5, 16, 23, 1, 10,
               19, 6, 26, 11, 4, 22, 8, 17, 24, 13, 20, 15];

  function buildSeats(offset) {
    var seats = {};
    seatIds.forEach(function (id) { seats[id] = null; });
    ORDER.forEach(function (studentIdx, pos) {
      var s = students[(studentIdx + offset) % students.length];
      seats[seatIds[pos]] = { studentId: s.id, name: s.name };
    });
    return seats;
  }

  var currentSeats = buildSeats(0);

  // 모둠 배치 시안(shot 14): 3모둠 × 4명 × 2행
  var isGroup = (shot === '14');
  var groupConfig = { totalDesks: 24, columns: 3, seatType: 'group', layoutIndex: 0, rows: 2 };
  var groupSeats = (function () {
    if (!isGroup) return null;
    var s = {};
    var idx = 0;
    for (var r = 0; r < 2; r++) {
      for (var c = 0; c < 12; c++) {
        var st = students[ORDER[idx % ORDER.length]];
        s[r + '-' + c] = { studentId: st.id, name: st.name };
        idx++;
      }
    }
    return s;
  })();

  var data = {
    settings: { periodStart: '2026-08-24', periodEnd: '2026-10-23' },
    currentClassIndex: 0,
    allClassData: {
      0: {
        className: '경수',
        students: students,
        currentSeats: isGroup ? groupSeats : currentSeats,
        lockedSeats: isGroup ? {} : { '0-0': '13', '4-5': 'empty' },
        config: isGroup ? groupConfig : { totalDesks: 28, columns: 5, seatType: 'pair', layoutIndex: 0, rows: ROWS },
        notice: '\u00b7 \ub0b4\uc77c \uccb4\uc721\ubcf5 \uc9c0\ucc38\n\u00b7 \uc218\ud559 \uc775\ud798\ucc45 42~45\ucabd\n\u00b7 \uae08\uc694\uc77c \ud604\uc7a5\uccb4\ud5d8\ud559\uc2b5 \ub3d9\uc758\uc11c \uc81c\ucd9c',
        noticeAt: '2026-08-24 08:12',
        praise: { '1': 4, '4': 5, '13': 3, '19': 3, '2': 2, '8': 1 },
        relations: {
          close: { '3': ['11'], '11': ['3'] },
          far: { '21': ['23'], '23': ['21'] }
        },
        studentMemos: {
          '5': '시력이 낮아 앞자리 배치가 필요합니다. 안경 착용.',
          '14': '왼손잡이 — 분단 왼쪽 끝자리를 선호합니다.',
          '21': '수업 중 집중이 흐트러질 때가 있어 앞쪽 배치 권장.'
        },
        attendance: {
          '2026-08-03': { '3': '지각', '16': '지각' },
          '2026-08-04': { '8': '조퇴', '27': '결석' },
          '2026-08-05': { '5': '결석', '22': '조퇴' },
          '2026-08-06': { '16': '결석' },
          '2026-08-07': { '11': '결과' }
        },
        attendReasons: {
          '2026-08-03': { '3': '인정', '16': '미인정' },
          '2026-08-04': { '8': '질병', '27': '인정' },
          '2026-08-05': { '5': '질병', '22': '인정' },
          '2026-08-06': { '16': '질병' },
          '2026-08-07': { '11': '인정' }
        },
        attendMemos: {
          '2026-08-03': { '3': '병원 진료 후 등교', '16': '1교시 후 등교' },
          '2026-08-04': { '8': '보건실 확인', '27': '체험학습' },
          '2026-08-05': { '5': '감기', '22': '가족 행사' },
          '2026-08-06': { '16': '몸살' },
          '2026-08-07': { '11': '대회 참가' }
        },        savedRecords: [
          {
            id: 'rec3', date: '2026-08-24T01:10:00.000Z', memo: '2학기 1차 배치',
            seats: currentSeats,
            config: { totalDesks: 28, columns: 5, seatType: 'pair', layoutIndex: 0, rows: ROWS },
            lockedSeats: { '0-0': '5' }, studentCount: 28
          },
          {
            id: 'rec2', date: '2026-06-15T02:20:00.000Z', memo: '1학기 3차 배치',
            seats: buildSeats(7),
            config: { totalDesks: 28, columns: 5, seatType: 'pair', layoutIndex: 0, rows: ROWS },
            lockedSeats: {}, studentCount: 28
          },
          {
            id: 'rec1', date: '2026-04-20T01:05:00.000Z', memo: '1학기 2차 배치',
            seats: buildSeats(15),
            config: { totalDesks: 28, columns: 5, seatType: 'pair', layoutIndex: 0, rows: ROWS },
            lockedSeats: {}, studentCount: 28
          }
        ]
      },
      1: { className: '3학년 5반', students: [], currentSeats: {}, lockedSeats: {},
           config: { totalDesks: 30, columns: 5, seatType: 'pair', layoutIndex: 0, rows: 6 },
           relations: { close: {}, far: {} }, studentMemos: {}, savedRecords: [] }
    }
  };

  localStorage.setItem('seatArrangement_v3', JSON.stringify(data));
})();
