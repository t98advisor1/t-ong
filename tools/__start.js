/* 시작화면 시안 비교용 — 배포 제외. ?v=A|B|C */
(function () {
  var v = new URLSearchParams(location.search).get('v') || 'A';
  function css(t) { var s = document.createElement('style'); s.textContent = t; document.head.appendChild(s); }

  // A — 3단계를 없애고 모든 설정을 한 화면에 작게 펼침
  function A() {
    css(
      '.step-tabs,.step-progress{display:none !important}' +
      '.wizard-hero{padding:14px 24px 12px !important}' +
      '.wizard-hero h2{font-size:1.25rem !important;margin-bottom:0 !important}' +
      '.wizard{grid-template-columns:minmax(0,1fr) 240px !important;max-width:1060px !important}' +
      '.wizard-content{display:grid !important;grid-template-columns:1fr 1fr;gap:10px 20px;' +
        'align-content:start;padding:16px 22px !important}' +
      '.wizard-content>.step{margin-bottom:0 !important;display:block !important}' +
      '.wizard-content>.step:nth-of-type(2){grid-column:1/-1}' +
      '.wizard-content>#layoutStep{display:block !important}' +
      '.wizard-content .step h3{font-size:0.88rem !important;margin-bottom:6px !important}' +
      '.wizard-nav{grid-column:1/-1;justify-content:flex-end;margin-top:6px}' +
      '#stepPrevBtn,#stepNextBtn{display:none !important}' +
      '#applyConfigBtn{display:inline-flex !important;min-width:190px}' +
      '.type-card{padding:7px 9px !important}' +
      '.seat-type-options{gap:8px !important}'
    );
    document.querySelectorAll('.wizard-content .step').forEach(function (s) { s.style.display = 'block'; });
  }

  // B — 지금 구조 유지 + 미리보기 썸네일을 크게
  function B() {
    css(
      '.wizard{grid-template-columns:minmax(0,1fr) 420px !important;max-width:1120px !important}' +
      '.preview-panel{background:linear-gradient(160deg,#f5f7ff,#eaeeff) !important;padding:20px !important;' +
        'justify-content:center !important;gap:14px !important}' +
      '.preview-panel h3{font-size:0.95rem !important}' +
      '.preview-panel .preview-desk{padding:10px 34px !important;font-size:0.9rem !important;' +
        'letter-spacing:8px;border-radius:0 0 14px 14px}' +
      '.preview-grid{gap:8px !important}' +
      '.preview-grid .preview-row{gap:14px !important}' +
      '.preview-grid .seat,.preview-grid>div>div>div{width:22px !important;height:20px !important;border-radius:5px !important}' +
      '.preview-info{font-size:0.95rem !important;font-weight:700;color:#4a4fd0 !important}' +
      '.preview-panel::after{content:"교실에 이렇게 놓입니다";display:block;text-align:center;' +
        'font-size:0.76rem;color:#9aa0b4;margin-top:4px}'
    );
  }

  // C — 시작화면을 기능 선택 허브로
  function C() {
    css(
      '.step-tabs,.step-progress,.wizard-content,.preview-panel{display:none !important}' +
      '.wizard{display:block !important;max-width:900px !important;overflow:hidden}' +
      '.wizard-hero{padding:20px 26px 18px !important}' +
      '.hub{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;padding:20px 26px 24px}' +
      '.hub-card{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:7px;' +
        'padding:20px 12px;border-radius:16px;cursor:pointer;border:none;font-family:inherit;' +
        'background:linear-gradient(145deg,#f5f7ff,#e9edff);box-shadow:0 2px 0 #d7ddf5,0 5px 14px rgba(102,126,234,.14)}' +
      '.hub-card.main{grid-column:1/-1;background:linear-gradient(135deg,#667eea,#764ba2);color:#fff;' +
        'box-shadow:0 6px 22px rgba(102,126,234,.42);padding:26px 12px}' +
      '.hub-ic{font-size:1.9rem;line-height:1}' +
      '.hub-card.main .hub-ic{font-size:2.4rem}' +
      '.hub-tt{font-size:0.9rem;font-weight:800;color:#3a3f5c}' +
      '.hub-card.main .hub-tt{color:#fff;font-size:1.15rem}' +
      '.hub-sub{font-size:0.72rem;color:#8a90a6}' +
      '.hub-card.main .hub-sub{color:rgba(255,255,255,.82);font-size:0.84rem}'
    );
    var items = [
      ['main', '🪑', '자리배치 시작하기', '교실 구조를 정하고 자리를 추첨합니다'],
      ['', '✅', '출결관리', '오늘 출결 체크'],
      ['', '📊', '통합 기록', '생기부용 자료'],
      ['', '⏱️', '간편 타이머', '조회 · 종례'],
      ['', '🧹', '1인 1역', '역할 배정'],
      ['', '🎯', '룰렛 뽑기', '발표자 정하기'],
      ['', '📖', '사용법', '처음이신가요?']
    ];
    var hub = document.createElement('div');
    hub.className = 'hub';
    items.forEach(function (it) {
      var b = document.createElement('button');
      b.type = 'button';
      b.className = 'hub-card ' + it[0];
      b.innerHTML = '<span class="hub-ic">' + it[1] + '</span>' +
        '<span class="hub-tt">' + it[2] + '</span>' +
        '<span class="hub-sub">' + it[3] + '</span>';
      hub.appendChild(b);
    });
    document.getElementById('setupWizard').appendChild(hub);
  }

  window.addEventListener('load', function () {
    setTimeout(function () {
      if (v === 'A') A(); else if (v === 'B') B(); else C();
    }, 300);
  });
})();
