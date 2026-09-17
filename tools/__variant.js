/* 메인 화면 구성 시안 비교용 — 배포 제외(.vercelignore).
   ?v=base | A | B | C  로 각 안을 실제 UI 위에 적용해 렌더링합니다. */
(function () {
  var v = new URLSearchParams(location.search).get('v') || 'base';

  function css(text) {
    var st = document.createElement('style');
    st.textContent = text;
    document.head.appendChild(st);
  }

  function label(text, cls) {
    var d = document.createElement('div');
    d.className = 'cr-side ' + cls;
    d.textContent = text;
    return d;
  }

  // ── A안: 교실 느낌 강화 (복도/창가 표시 + 교탁 강조 + 좌석 확대)
  function applyA() {
    css(
      '.cr-wrap{display:flex;align-items:stretch;justify-content:center;gap:14px}' +
      '.cr-side{writing-mode:vertical-rl;text-orientation:upright;letter-spacing:6px;' +
        'font-weight:700;font-size:0.82rem;color:#8a90a6;background:#eef0f7;border-radius:10px;' +
        'padding:10px 4px;display:flex;align-items:center;justify-content:center;min-width:30px;' +
        'border:1px dashed #d3d7e6}' +
      '.cr-side.window{background:#eaf4ff;color:#5b87c7;border-color:#c6ddf5}' +
      '.desk{max-width:none !important;width:min(560px,70%);padding:12px 40px !important;' +
        'border-radius:0 0 18px 18px;letter-spacing:10px;font-size:1.05rem;box-shadow:0 6px 18px rgba(102,126,234,0.32)}' +
      '.seat-card{width:120px;min-height:94px}' +
      '.seat-row{gap:30px}' +
      '.classroom-floor{border:2px dashed #dfe3ee;border-radius:18px;padding:16px 18px 18px;background:#fbfcff}'
    );
    var area = document.getElementById('seatingArea');
    var floor = document.createElement('div');
    floor.className = 'classroom-floor';
    var wrap = document.createElement('div');
    wrap.className = 'cr-wrap';
    area.parentNode.insertBefore(floor, area);
    floor.appendChild(wrap);
    wrap.appendChild(label('복도', 'hall'));
    wrap.appendChild(area);
    wrap.appendChild(label('창가', 'window'));
  }

  // ── B안: 상단 버튼 정리 (주요 3개 + 더보기)
  function applyB() {
    css(
      '.action-bar{padding:8px 20px}' +
      '.btn{padding:8px 12px;font-size:0.86rem}' +
      '.more-wrap{position:relative}' +
      '.more-menu{position:absolute;right:0;top:calc(100% + 8px);background:#fff;border-radius:12px;' +
        'box-shadow:0 8px 28px rgba(20,26,60,0.16);border:1px solid #e8eaf2;padding:6px;min-width:190px;z-index:60}' +
      '.more-menu button{display:block;width:100%;text-align:left;background:none;border:none;' +
        'padding:10px 12px;border-radius:8px;cursor:pointer;font-size:0.88rem;font-family:inherit}' +
      '.more-menu button:hover{background:#f2f4fb}' +
      '.seat-card{width:118px;min-height:92px}'
    );
    var right = document.querySelector('.action-right');
    var keep = ['saveBtn', 'studentViewBtn', 'guideBtn'];
    var moved = [];
    [].slice.call(right.children).forEach(function (b) {
      if (keep.indexOf(b.id) === -1) moved.push(b);
    });
    var wrap = document.createElement('div');
    wrap.className = 'more-wrap';
    var more = document.createElement('button');
    more.className = 'btn';
    more.textContent = '⋯ 더보기';
    wrap.appendChild(more);
    var menu = document.createElement('div');
    menu.className = 'more-menu';
    moved.forEach(function (b) { menu.appendChild(b); });
    wrap.appendChild(menu);
    right.appendChild(wrap);
  }

  // ── C안: 분단 통로 + 행 번호 + 분단 라벨
  function applyC() {
    css(
      '.seat-row{gap:46px}' +
      '.row-num{position:absolute;left:-6px;top:50%;transform:translateY(-50%);font-size:0.75rem;' +
        'font-weight:700;color:#a3a8bd;background:#eef0f7;border-radius:6px;padding:3px 7px}' +
      '.seat-row{position:relative;padding-left:34px}' +
      '.group-labels{display:flex;justify-content:center;gap:46px;margin:0 0 8px 34px}' +
      '.group-label{font-size:0.76rem;font-weight:700;color:#7b8199;background:#f0f2f9;' +
        'border-radius:999px;padding:4px 14px;text-align:center}'
    );
    var area = document.getElementById('seatingArea');
    var rows = [].slice.call(area.children);
    // 분단 라벨
    if (rows.length) {
      var bar = document.createElement('div');
      bar.className = 'group-labels';
      [].slice.call(rows[0].children).forEach(function (g, i) {
        var l = document.createElement('div');
        l.className = 'group-label';
        l.textContent = (i + 1) + '분단';
        l.style.width = g.getBoundingClientRect().width + 'px';
        bar.appendChild(l);
      });
      area.parentNode.insertBefore(bar, area);
    }
    // 행 번호
    rows.forEach(function (r, i) {
      var n = document.createElement('div');
      n.className = 'row-num';
      n.textContent = (i + 1) + '행';
      r.appendChild(n);
    });
  }

  window.addEventListener('load', function () {
    setTimeout(function () {
      if (v === 'A') applyA();
      else if (v === 'B') applyB();
      else if (v === 'C') applyC();
      window.dispatchEvent(new Event('resize'));   // 좌석 자동 맞춤 재계산
    }, 350);
  });
})();
