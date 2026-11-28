/* ===================================================================
 * T-ong 스마트 판서 보드 v2.0
 *  - 한쪽 접이식 통합 팔레트 & 상하 여백 0% 전체화면 보드 (스마트패드/갤럭시탭 최적화)
 *  - 삼성노트 스타일 로컬 폴더 & 판서 노트 보관함 (IndexedDB)
 *  - 고호환 PPTX 파서 (그룹 도형 grpSp, 레이아웃 상속 ph, 문단별·도형별 클릭 액션)
 *  - 내장 웹 브라우저 + 웹 위 ①스탬프/판서 레이어
 *  - SaniZ 4대 모드 (게시·확대·판서·편집) & 무료 OCR 손글씨 변환
 * =================================================================== */
(function () {
  'use strict';

  var OCR_CFG_KEY = 'tong_smart_board_ocr_cfg_v1';
  var MEMO_NOTES_KEY = 'tong_memo_board_notes_v1';
  var FOLDERS_KEY = 'tong_smart_board_folders_v1';
  var FALLBACK_NOTES_KEY = 'tong_smart_board_notes_fallback_v1';
  var IDB_NAME = 'tong_smart_board_db_v1';
  var IDB_STORE = 'notes';

  var bgCanvas = document.getElementById('sbBgCanvas');
  var inkCanvas = document.getElementById('sbInkCanvas');
  var liveCanvas = document.getElementById('sbLiveCanvas');
  var bgCtx = bgCanvas.getContext('2d');
  var inkCtx = inkCanvas.getContext('2d');
  var liveCtx = liveCanvas.getContext('2d');

  var stageWrap = document.getElementById('sbStageWrap');
  var boardFrame = document.getElementById('sbBoardFrame');
  var inlineTextEl = document.getElementById('sbInlineText');
  var pinLayer = document.getElementById('sbPinLayer');
  var editBar = document.getElementById('sbEditBar');
  var webFrame = document.getElementById('sbWebFrame');
  var webBar = document.getElementById('sbWebBar');

  var state = {
    mode: 'draw',          // 'pin' | 'zoom' | 'draw' | 'edit'
    tool: 'pen',           // 'pen' | 'highlighter' | 'laser' | 'line' | 'rect' | 'fillRect' | 'cover' | 'text' | 'stamp' | 'eraser'
    penStyle: 'fountain',  // 'fountain' | 'ballpoint' | 'pencil' | 'brush'
    color: '#ef4444',
    size: 2,
    penSize: 2,
    highlighterSize: 6,
    bgMode: 'dark',        // 'dark' | 'light' | 'grid'
    palmReject: false,
    spenAuto: (function () { try { return localStorage.getItem('sbSpenAuto_v1') !== 'off'; } catch (e) { return true; } })(), // S펜이 닿으면 S펜 전용 모드를 자동으로 켭니다
    spenManualOff: false,
    spenEraserOnBtn: true, // 갤럭시 S펜 옆 버튼 누르기 / 에어액션 시 [지우개] <-> [펜] 양방향 스위치
    lastSpenSwitchTime: 0,
    lastDrawTool: 'pen',
    autoHide: false,        // 상단 팔레트 자동 숨김 (판서 시작 시 자동으로 위로 숨김)
    fillScreen: true,      // 화면 송출 시 여백 없는 100% 꽉 찬 화면 모드
    stampSeries: 0,   // '스탬프 번호 초기화'를 누를 때마다 새 묶음으로 1번부터 다시 셉니다
    autoStampShape: (function () { try { return localStorage.getItem('sbAutoStampShape_v1') !== 'off'; } catch (e) { return true; } })(),
    lastStampShape: 'circle',
    zoomScale: 1,
    panX: 0,
    panY: 0,
    baseScale: 1,
    pages: [],
    pageIndex: 0,
    pdfDoc: null,
    pdfCache: {},
    selection: null,
    laserStrokes: [],
    laserRaf: null,
    ocrPendingBox: null,
    pptxShowAll: false,
    webActive: false,
    webInteract: false,
    currentNoteId: null,
    currentNoteTitle: '새 판서 노트',
    currentFolder: '1학년 수학',
    activeVaultFolder: 'ALL'
  };

  var pointer = {
    down: false,
    activeId: null,
    startX: 0,
    startY: 0,
    lastX: 0,
    lastY: 0,
    moved: false,
    effectiveTool: 'pen',
    currentStroke: null,
    draggingSelection: false,
    lastHoverSideBtn: false,
    pinchMap: {}
  };

  function el(id) { return document.getElementById(id); }

  var toastTimer = null;
  function showToast(msg) {
    var t = el('sbToast');
    if (!t) return;
    t.textContent = msg;
    t.style.display = 'block';
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.style.display = 'none'; }, 2500);
  }

  function uid() { return 'o_' + Date.now() + '_' + Math.floor(Math.random() * 10000); }

  function curPage() {
    if (!state.pages.length) {
      state.pages.push(createBlankPage('dark'));
      state.pageIndex = 0;
    }
    return state.pages[state.pageIndex];
  }

  function getTrueViewportSize() {
    var vv = window.visualViewport;
    var candsW = [window.innerWidth, document.documentElement.clientWidth];
    var candsH = [window.innerHeight, document.documentElement.clientHeight];
    if (vv && vv.width > 100) candsW.push(vv.width);
    if (vv && vv.height > 100) candsH.push(vv.height);
    if (stageWrap) {
      var sr = stageWrap.getBoundingClientRect();
      if (sr.width > 100) candsW.push(sr.width);
      if (sr.height > 100) candsH.push(sr.height);
    }
    var validW = candsW.filter(function (v) { return typeof v === 'number' && v > 100; });
    var validH = candsH.filter(function (v) { return typeof v === 'number' && v > 100; });
    return {
      w: validW.length ? Math.min.apply(null, validW) : 1280,
      h: validH.length ? Math.min.apply(null, validH) : 720
    };
  }

  function getIdealBoardSize() {
    var vp = getTrueViewportSize();
    var rw = Math.max(320, vp.w);
    var rh = Math.max(240, vp.h);
    var bw = 1280;
    var bh = state.fillScreen ? Math.round(bw * (rh / rw)) : 720;
    return { w: bw, h: Math.max(400, Math.min(2000, bh)) };
  }

  function createBlankPage(bgMode) {
    var sz = getIdealBoardSize();
    return {
      kind: 'blank',
      label: bgMode === 'light' ? '백지' : (bgMode === 'grid' ? '모눈' : '칠판'),
      width: sz.w,
      height: sz.h,
      bgMode: bgMode || state.bgMode,
      objects: [],
      redoStack: []
    };
  }

  // ===================================================================
  // 상단 가로형 팔레트 접기/펴기(자동 숨김) & 100% 꽉 찬 전체화면 뷰포트 맞춤
  // ===================================================================
  var autoHideTimer = null;
  function scheduleAutoHidePalette(delayMs) {
    clearTimeout(autoHideTimer);
    if (!state.autoHide) return;
    autoHideTimer = setTimeout(function () {
      var anyModalOpen = document.querySelector('.sb-modal.show');
      if (!anyModalOpen) togglePalette(true);
    }, delayMs || 3200);
  }

  function togglePalette(forceCollapse) {
    var pal = el('sbPalette');
    var opener = el('sbFloatingOpener');
    if (!pal) return;
    var collapsed = typeof forceCollapse === 'boolean' ? forceCollapse : !pal.classList.contains('collapsed');
    pal.classList.toggle('collapsed', collapsed);
    if (opener) opener.style.display = collapsed ? 'inline-flex' : 'none';
    setTimeout(fitViewport, 30);
    setTimeout(fitViewport, 220);
  }

  function resizeBoardToPage() {
    var pg = curPage();
    if (pg.kind === 'blank' && state.fillScreen) {
      var sz = getIdealBoardSize();
      pg.width = sz.w;
      pg.height = sz.h;
    }
    var w = pg.width || 1280;
    var h = pg.height || 720;
    [bgCanvas, inkCanvas, liveCanvas].forEach(function (cv) {
      if (cv.width !== w || cv.height !== h) {
        cv.width = w;
        cv.height = h;
      }
    });
    fitViewport();
  }

  function fitViewport() {
    var pg = curPage();
    var vp = getTrueViewportSize();
    if (state.webActive) {
      var palette = el('sbPalette');
      var top = palette && !palette.classList.contains('collapsed') ? palette.getBoundingClientRect().height + 8 : 42;
      var bottom = (webBar ? webBar.getBoundingClientRect().height : 72) + 24;
      document.documentElement.style.setProperty('--sb-web-top', top + 'px');
      document.documentElement.style.setProperty('--sb-web-bottom', bottom + 'px');
      vp = {w: vp.w - 16, h: Math.max(100, vp.h - top - bottom)};
    }
    var rw = Math.max(100, vp.w);
    var rh = Math.max(100, vp.h);
    var w = pg.width || 1280;
    var h = pg.height || 720;
    var sx = rw / w;
    var sy = rh / h;
    // 1:1 정방형 픽셀 비율 유지 (원형 스탬프가 절대 찌그러지지 않으면서 화면을 꽉 채움)
    state.baseScale = state.fillScreen ? Math.max(sx, sy) : Math.min(sx, sy);
    if (state.webActive || pg.kind === 'pdf' || pg.kind === 'pptx' || pg.kind === 'image') {
      // 문서/슬라이드는 우측·하단 잘림 없이 100% 화면 안에 들어오도록 비율 맞춤
      state.baseScale = Math.min(sx, sy);
    }
    // 1280px 고정 DOM 박스로 인한 태블릿(960px 등) 레이아웃 뷰포트 확장·우측 1/4 잘림을 원천 차단하기 위해
    // boardFrame 의 CSS 크기를 실제 화면 맞춤 픽셀(fitW x fitH)로 직접 지정하고 내부 캔버스(1280x720)만 GPU 스케일링
    var fitW = Math.max(100, Math.round(w * state.baseScale));
    var fitH = Math.max(100, Math.round(h * state.baseScale));
    boardFrame.style.width = fitW + 'px';
    boardFrame.style.height = fitH + 'px';
    clampPan();
    applyTransform();
  }

  // ===================================================================
  // 확대·이동 — 판서판 왼쪽 위 고정 레이아웃 위치(offsetLeft/offsetTop)를 기준점으로 확대하고,
  // 이동 범위는 실제 보이는 화면(#sbViewport) 기준으로 제한합니다.
  // ===================================================================
  function roundSubpx(v) {
    return Math.round(v * 100) / 100;
  }
  function zoomGeometry() {
    var vpEl = el('sbViewport') || stageWrap;
    var sr = vpEl.getBoundingClientRect();
    // style.transform 과 무관한 순수 레이아웃 원점을 사용해 확대·터치 이동 중 기준점 진동 및 좌측 쏠림을 원천 차단합니다.
    return {
      sr: sr,
      fw: boardFrame.offsetWidth,
      fh: boardFrame.offsetHeight,
      left0: sr.left + boardFrame.offsetLeft,
      top0: sr.top + boardFrame.offsetTop
    };
  }
  function panRange(g, z) {
    function axis(viewStart, viewSize, start0, size) {
      var content = size * z;
      if (content <= viewSize) { var c = viewStart + (viewSize - content) / 2 - start0; return { lo: c, hi: c }; }
      return { lo: viewStart + viewSize - start0 - content, hi: viewStart - start0 };
    }
    return { x: axis(g.sr.left, g.sr.width, g.left0, g.fw), y: axis(g.sr.top, g.sr.height, g.top0, g.fh) };
  }
  function clampPan() {
    if (state.zoomScale <= 1.02) { state.zoomScale = 1; state.panX = 0; state.panY = 0; return; }
    var r = panRange(zoomGeometry(), state.zoomScale);
    state.panX = roundSubpx(Math.min(r.x.hi, Math.max(r.x.lo, state.panX)));
    state.panY = roundSubpx(Math.min(r.y.hi, Math.max(r.y.lo, state.panY)));
  }
  // 화면의 한 점(clientX, clientY)을 붙잡은 채 배율을 바꿉니다(휠·두 손가락·확대 시작).
  function zoomAt(clientX, clientY, nextZoom, fromZoom, fromPanX, fromPanY) {
    var g = zoomGeometry();
    var z0 = fromZoom != null ? fromZoom : state.zoomScale;
    var px0 = fromPanX != null ? fromPanX : state.panX, py0 = fromPanY != null ? fromPanY : state.panY;
    var z1 = Math.max(1, Math.min(4, nextZoom));
    // 기준 좌표가 판서판 바깥 여백에 찍혀 절반이 한쪽으로 밀리는 현상을 막기 위해 판 영역 안으로 보정합니다.
    var curLeft = g.left0 + px0, curTop = g.top0 + py0;
    var curW = g.fw * z0, curH = g.fh * z0;
    var safeX = Math.min(curLeft + curW, Math.max(curLeft, clientX));
    var safeY = Math.min(curTop + curH, Math.max(curTop, clientY));
    var ux = (safeX - g.left0 - px0) / z0, uy = (safeY - g.top0 - py0) / z0;
    state.zoomScale = z1;
    state.panX = safeX - g.left0 - z1 * ux;
    state.panY = safeY - g.top0 - z1 * uy;
    clampPan();
    applyTransform();
  }

  function applyTransform() {
    boardFrame.style.transformOrigin = '0 0';
    boardFrame.style.transform = 'translate3d(' + roundSubpx(state.panX) + 'px, ' + roundSubpx(state.panY) + 'px, 0) scale(' + state.zoomScale + ')';
    var zb = el('sbZoomBadge');
    var zt = el('sbZoomText');
    if (zb && zt) {
      if (state.zoomScale > 1.02 || state.mode === 'zoom') {
        zb.style.display = 'inline-flex';
        zt.textContent = (state.mode === 'zoom' ? '확대 모드 ' : '확대 고정 ') + state.zoomScale.toFixed(1) + 'x';
      } else {
        zb.style.display = 'none';
      }
    }
  }

  // 확대 켜기/끄기 (Ctrl+1 · 확대 버튼) — 확대 중이면 100%로 돌아가 판서 모드로, 아니면 확대 모드로 들어갑니다.
  function toggleZoom() {
    if (state.mode === 'zoom' || state.zoomScale > 1.02) {
      resetZoom();
      restorePreZoom();
      showToast('확대를 끄고 확대 전 화면·도구로 돌아왔습니다.');
    } else {
      setMode('zoom');
    }
  }

  function resetZoom() {
    state.zoomScale = 1;
    state.panX = 0;
    state.panY = 0;
    applyTransform();
  }

  function clientToBoard(clientX, clientY) {
    var r = boardFrame.getBoundingClientRect();
    var pg = curPage();
    return {
      x: ((clientX - r.left) / r.width) * pg.width,
      y: ((clientY - r.top) / r.height) * pg.height
    };
  }

  // ===================================================================
  // 배경 렌더링 (칠판 / 백지 / 모눈 / 이미지 / PDF / PPTX)
  // ===================================================================
  async function renderBackground() {
    var pg = curPage();
    var w = pg.width, h = pg.height;
    bgCtx.clearRect(0, 0, w, h);

    if (pg.kind === 'blank') {
      drawBlankBg(bgCtx, w, h, pg.bgMode || state.bgMode);
    } else if (pg.kind === 'image') {
      bgCtx.fillStyle = '#0f172a';
      bgCtx.fillRect(0, 0, w, h);
      if (!pg.imageObj && pg.imageDataUrl) {
        pg.imageObj = await loadImgElement(pg.imageDataUrl);
      }
      if (pg.imageObj) bgCtx.drawImage(pg.imageObj, 0, 0, w, h);
    } else if (pg.kind === 'pdf' && state.pdfDoc) {
      bgCtx.fillStyle = '#ffffff';
      bgCtx.fillRect(0, 0, w, h);
      var cached = state.pdfCache[pg.pdfPageNum];
      if (!cached) {
        cached = await renderPdfPageToOffscreen(pg.pdfPageNum);
        state.pdfCache[pg.pdfPageNum] = cached;
      }
      if (cached) bgCtx.drawImage(cached, 0, 0, w, h);
    } else if (pg.kind === 'pptx' && pg.pptxSlide) {
      await ensurePptxSlideImages(pg.pptxSlide);
      renderPptxSlideToCtx(pg.pptxSlide, state.pptxShowAll ? 9999 : (pg.pptxStep || 0), bgCtx, w, h);
    }
    updateFooterUi();
  }

  function loadImgElement(src) {
    return new Promise(function (resolve) {
      var im = new Image();
      im.crossOrigin = 'anonymous';
      im.onload = function () { resolve(im); };
      im.onerror = function () { resolve(null); };
      im.src = src;
    });
  }

  function drawBlankBg(ctx, w, h, bgMode) {
    if (bgMode === 'light') {
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, w, h);
    } else if (bgMode === 'grid') {
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, w, h);
      ctx.strokeStyle = '#e2e8f0';
      ctx.lineWidth = 1;
      for (var x = 0; x <= w; x += 32) { ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, h); ctx.stroke(); }
      for (var y = 0; y <= h; y += 32) { ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(w, y); ctx.stroke(); }
      ctx.strokeStyle = '#cbd5e1';
      ctx.lineWidth = 1.6;
      ctx.beginPath(); ctx.moveTo(w / 2, 0); ctx.lineTo(w / 2, h); ctx.stroke();
      ctx.beginPath(); ctx.moveTo(0, h / 2); ctx.lineTo(w, h / 2); ctx.stroke();
    } else {
      ctx.fillStyle = '#142029';
      ctx.fillRect(0, 0, w, h);
    }
  }

  // ===================================================================
  // 판서(벡터 객체) 렌더링 및 바운딩 박스
  // ===================================================================
  function getObjBounds(o) {
    if (o.type === 'stroke' || o.type === 'highlighter') {
      var minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
      (o.points || []).forEach(function (pt) {
        if (pt.x < minX) minX = pt.x;
        if (pt.y < minY) minY = pt.y;
        if (pt.x > maxX) maxX = pt.x;
        if (pt.y > maxY) maxY = pt.y;
      });
      var pad = (o.size || 5) * 0.7;
      return { x: minX - pad, y: minY - pad, w: (maxX - minX) + pad * 2, h: (maxY - minY) + pad * 2 };
    }
    if (o.type === 'line') {
      var lx = Math.min(o.x1, o.x2), ly = Math.min(o.y1, o.y2);
      var lp = o.arrow ? arrowHeadSize(o) + 4 : 6;
      return { x: lx - lp, y: ly - lp, w: Math.abs(o.x2 - o.x1) + lp * 2, h: Math.abs(o.y2 - o.y1) + lp * 2 };
    }
    if (o.type === 'rect' || o.type === 'fillRect' || o.type === 'cover' || o.type === 'circle') {
      return { x: Math.min(o.x, o.x + o.w), y: Math.min(o.y, o.y + o.h), w: Math.abs(o.w), h: Math.abs(o.h) };
    }
    if (o.type === 'stamp') {
      var r = o.radius || 22;
      return { x: o.x - r, y: o.y - r, w: r * 2, h: r * 2 };
    }
    if (o.type === 'text') {
      var lines = String(o.text || '').split('\n');
      var fs = o.fontSize || 26;
      var maxLen = Math.max.apply(null, lines.map(function (l) { return l.length; }).concat([1]));
      return { x: o.x, y: o.y - fs, w: Math.max(40, maxLen * fs * 0.85), h: lines.length * fs * 1.35 + 8 };
    }
    return { x: 0, y: 0, w: 0, h: 0 };
  }

  function rectsIntersect(a, b) {
    return !(a.x + a.w < b.x || b.x + b.w < a.x || a.y + a.h < b.y || b.y + b.h < a.y);
  }

  function translateObj(o, dx, dy) {
    if (o.type === 'stroke' || o.type === 'highlighter') {
      o.points.forEach(function (pt) { pt.x += dx; pt.y += dy; });
    } else if (o.type === 'line') {
      o.x1 += dx; o.y1 += dy; o.x2 += dx; o.y2 += dy;
    } else if (o.type === 'rect' || o.type === 'fillRect' || o.type === 'cover' || o.type === 'circle' || o.type === 'stamp' || o.type === 'text') {
      o.x += dx; o.y += dy;
    }
  }

  // 화살표: 직선(line) 객체에 arrow = 1(한쪽) / 2(양쪽)을 붙여 그립니다. 선택·이동·지우개는 직선과 똑같이 동작합니다.
  function arrowHeadSize(o) { return Math.max(20, (o.size || 5) * 4.2); }
  function drawArrowLine(ctx, o) {
    var dx = o.x2 - o.x1, dy = o.y2 - o.y1, len = Math.sqrt(dx * dx + dy * dy) || 1;
    var ux = dx / len, uy = dy / len, hs = Math.min(arrowHeadSize(o), len * 0.6);
    var sx = o.x1, sy = o.y1, ex = o.x2, ey = o.y2;
    // 선 끝이 화살촉 밖으로 삐져나오지 않도록 촉 길이의 절반만큼 줄여서 긋습니다.
    ex -= ux * hs * 0.5; ey -= uy * hs * 0.5;
    if (o.arrow === 2) { sx += ux * hs * 0.5; sy += uy * hs * 0.5; }
    ctx.lineCap = 'round';
    ctx.beginPath(); ctx.moveTo(sx, sy); ctx.lineTo(ex, ey); ctx.stroke();
    ctx.fillStyle = o.color || '#ef4444';
    function head(tx, ty, vx, vy) {
      var bx = tx - vx * hs, by = ty - vy * hs, w = hs * 0.55;
      ctx.beginPath();
      ctx.moveTo(tx, ty);
      ctx.lineTo(bx - vy * w, by + vx * w);
      ctx.lineTo(bx + vy * w, by - vx * w);
      ctx.closePath();
      ctx.fill();
    }
    head(o.x2, o.y2, ux, uy);
    if (o.arrow === 2) head(o.x1, o.y1, -ux, -uy);
  }

  function drawInkObject(ctx, o) {
    ctx.save();
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    if (o.type === 'stroke' || o.type === 'highlighter') {
      var pts = o.points || [];
      if (!pts.length) { ctx.restore(); return; }
      var baseSize = o.size || 2;
      var pStyle = o.type === 'highlighter' ? 'highlighter' : (o.penStyle || 'ballpoint');
      ctx.strokeStyle = o.color || '#ef4444';

      if (pStyle === 'highlighter') {
        ctx.lineWidth = Math.max(12, baseSize * 4.4);
        ctx.globalAlpha = 0.34;
        ctx.beginPath();
        ctx.moveTo(pts[0].x, pts[0].y);
        if (pts.length === 1) {
          ctx.lineTo(pts[0].x + 0.1, pts[0].y + 0.1);
        } else {
          for (var hi = 1; hi < pts.length; hi++) {
            var hmx = (pts[hi - 1].x + pts[hi].x) / 2;
            var hmy = (pts[hi - 1].y + pts[hi].y) / 2;
            ctx.quadraticCurveTo(pts[hi - 1].x, pts[hi - 1].y, hmx, hmy);
          }
          ctx.lineTo(pts[pts.length - 1].x, pts[pts.length - 1].y);
        }
        ctx.stroke();
      } else if (pStyle === 'fountain' || pStyle === 'brush') {
        // 삼성노트 만년필/붓펜: 필압·속도 기반 가변 굵기 + 부드러운 2차 베지어 곡선
        var minMul = pStyle === 'brush' ? 0.42 : 0.68;
        var maxMul = pStyle === 'brush' ? 1.65 : 1.18;
        if (pts.length === 1) {
          ctx.fillStyle = o.color || '#ef4444';
          ctx.beginPath();
          ctx.arc(pts[0].x, pts[0].y, Math.max(0.8, baseSize * 0.5), 0, Math.PI * 2);
          ctx.fill();
        } else {
          var prevMx = pts[0].x, prevMy = pts[0].y;
          for (var fi = 1; fi < pts.length; fi++) {
            var pPrev = pts[fi - 1], pCur = pts[fi];
            var mx = (pPrev.x + pCur.x) / 2;
            var my = (pPrev.y + pCur.y) / 2;
            var pr = typeof pCur.p === 'number' ? pCur.p : 0.5;
            // 끝단 자연스러운 테이퍼링
            var endFade = 1;
            if (fi === 1) endFade = 0.82;
            else if (fi === pts.length - 1) endFade = 0.72;
            var segW = Math.max(0.8, baseSize * (minMul + (maxMul - minMul) * pr) * endFade);
            ctx.lineWidth = segW;
            ctx.beginPath();
            ctx.moveTo(prevMx, prevMy);
            ctx.quadraticCurveTo(pPrev.x, pPrev.y, mx, my);
            ctx.stroke();
            prevMx = mx;
            prevMy = my;
          }
        }
      } else if (pStyle === 'pencil') {
        // 삼성노트 연필 효과: 필압 반응형 흑연 코어 + 건조한 다중 섬유결 + 종이 요철 미세 흑연 입자(깜빡임 없는 결정론적 해시)
        var pHash = function (a, b, c) {
          var v = Math.sin(a * 12.9898 + b * 78.233 + c * 37.719) * 43758.5453;
          return v - Math.floor(v);
        };
        var pCol = o.color || '#ef4444';
        ctx.fillStyle = pCol;
        if (pts.length === 1) {
          ctx.globalAlpha = 0.52;
          ctx.beginPath();
          ctx.arc(pts[0].x, pts[0].y, Math.max(0.9, baseSize * 0.55), 0, Math.PI * 2);
          ctx.fill();
        } else {
          // 1층: 필압에 따라 농도와 굵기가 변하는 부드러운 흑연 중심선
          var pPrevMx = pts[0].x, pPrevMy = pts[0].y;
          for (var pi = 1; pi < pts.length; pi++) {
            var q0 = pts[pi - 1], q1 = pts[pi];
            var qmx = (q0.x + q1.x) / 2, qmy = (q0.y + q1.y) / 2;
            var qpr = typeof q1.p === 'number' ? q1.p : 0.55;
            ctx.globalAlpha = 0.34 + 0.36 * qpr;
            ctx.lineWidth = Math.max(0.75, baseSize * (0.8 + 0.34 * qpr));
            ctx.beginPath();
            ctx.moveTo(pPrevMx, pPrevMy);
            ctx.quadraticCurveTo(q0.x, q0.y, qmx, qmy);
            ctx.stroke();
            pPrevMx = qmx;
            pPrevMy = qmy;
          }
          // 2층: 좌우 미세하게 갈라지는 건조한 연필심 섬유결 2가닥
          var strandSpread = Math.max(0.55, baseSize * 0.44);
          for (var stIdx = -1; stIdx <= 1; stIdx += 2) {
            ctx.globalAlpha = 0.17;
            ctx.lineWidth = Math.max(0.5, baseSize * 0.46);
            ctx.beginPath();
            ctx.moveTo(pts[0].x + stIdx * strandSpread * 0.5, pts[0].y - stIdx * strandSpread * 0.5);
            for (var si = 1; si < pts.length; si++) {
              var sp0 = pts[si - 1], sp1 = pts[si];
              var sdx = sp1.x - sp0.x, sdy = sp1.y - sp0.y;
              var slen = Math.hypot(sdx, sdy) || 1;
              var nx = -sdy / slen, ny = sdx / slen;
              var jit = (pHash(si, Math.round(sp1.x), stIdx + 3) - 0.5) * baseSize * 0.38;
              var off = stIdx * strandSpread + jit;
              ctx.lineTo(sp1.x + nx * off, sp1.y + ny * off);
            }
            ctx.stroke();
          }
          // 3층: 종이 결(Paper tooth) 위에 묻어나는 사각거리는 흑연 가루 입자
          var speckBudget = 900;
          var speckCount = 0;
          for (var gi = 1; gi < pts.length && speckCount < speckBudget; gi++) {
            var g0 = pts[gi - 1], g1 = pts[gi];
            var gdx = g1.x - g0.x, gdy = g1.y - g0.y;
            var segL = Math.hypot(gdx, gdy);
            if (segL < 0.5) continue;
            var gnx = -gdy / segL, gny = gdx / segL;
            var gpr = typeof g1.p === 'number' ? g1.p : 0.55;
            var steps = Math.min(18, Math.max(1, Math.floor(segL / 2.3)));
            for (var gs = 0; gs < steps && speckCount < speckBudget; gs++) {
              var gt = (gs + 0.5) / steps;
              var gx = g0.x + gdx * gt;
              var gy = g0.y + gdy * gt;
              var r1 = pHash(gi * 31 + gs, Math.round(gx), Math.round(gy));
              if (r1 > 0.74) continue; // 자연스러운 종이 빈틈 유지
              var r2 = pHash(gs * 17 + gi, Math.round(gy), Math.round(gx));
              var spread = (r2 - 0.5) * 2 * Math.max(1.1, baseSize * 0.82);
              var sz = Math.max(0.65, baseSize * (0.24 + 0.24 * r1));
              ctx.globalAlpha = (0.16 + 0.32 * gpr) * (0.55 + 0.45 * r1);
              ctx.fillRect(gx + gnx * spread - sz * 0.5, gy + gny * spread - sz * 0.5, sz, sz);
              speckCount++;
            }
          }
        }
      } else {
        // 볼펜 (ballpoint): 일정하고 또렷한 선 (작은 글씨·수식 최적)
        ctx.lineWidth = Math.max(0.9, baseSize * 0.94);
        ctx.beginPath();
        ctx.moveTo(pts[0].x, pts[0].y);
        if (pts.length === 1) {
          ctx.lineTo(pts[0].x + 0.1, pts[0].y + 0.1);
        } else {
          for (var bi = 1; bi < pts.length; bi++) {
            var bmx = (pts[bi - 1].x + pts[bi].x) / 2;
            var bmy = (pts[bi - 1].y + pts[bi].y) / 2;
            ctx.quadraticCurveTo(pts[bi - 1].x, pts[bi - 1].y, bmx, bmy);
          }
          ctx.lineTo(pts[pts.length - 1].x, pts[pts.length - 1].y);
        }
        ctx.stroke();
      }
    } else if (o.type === 'line') {
      ctx.strokeStyle = o.color || '#ef4444';
      ctx.lineWidth = o.size || 5;
      if (o.isHighlighter) ctx.globalAlpha = 0.34;
      if (o.arrow) { drawArrowLine(ctx, o); }
      else {
        ctx.beginPath();
        ctx.moveTo(o.x1, o.y1);
        ctx.lineTo(o.x2, o.y2);
        ctx.stroke();
      }
    } else if (o.type === 'rect') {
      ctx.strokeStyle = o.color || '#ef4444';
      ctx.lineWidth = o.size || 5;
      if (o.isHighlighter) ctx.globalAlpha = 0.34;
      ctx.strokeRect(o.x, o.y, o.w, o.h);
    } else if (o.type === 'circle') {
      ctx.strokeStyle = o.color || '#ef4444';
      ctx.lineWidth = o.size || 5;
      if (o.isHighlighter) ctx.globalAlpha = 0.34;
      var cx = o.x + o.w / 2;
      var cy = o.y + o.h / 2;
      var rx = Math.max(1, Math.abs(o.w) / 2);
      var ry = Math.max(1, Math.abs(o.h) / 2);
      if (o.perfect) {
        var pr = Math.max(rx, ry);
        rx = pr; ry = pr;
      }
      ctx.beginPath();
      ctx.ellipse(cx, cy, rx, ry, 0, 0, Math.PI * 2);
      ctx.stroke();
    } else if (o.type === 'fillRect') {
      ctx.fillStyle = o.color || '#ef4444';
      ctx.globalAlpha = 0.88;
      ctx.fillRect(o.x, o.y, o.w, o.h);
    } else if (o.type === 'cover') {
      var rx2 = Math.min(o.x, o.x + o.w), ry2 = Math.min(o.y, o.y + o.h);
      var rw = Math.abs(o.w), rh = Math.abs(o.h);
      if (o.revealed) {
        ctx.strokeStyle = '#38bdf8';
        ctx.setLineDash([6, 4]);
        ctx.lineWidth = 2;
        ctx.strokeRect(rx2, ry2, rw, rh);
      } else {
        ctx.fillStyle = '#1e293b';
        ctx.fillRect(rx2, ry2, rw, rh);
        ctx.strokeStyle = '#38bdf8';
        ctx.lineWidth = 2.5;
        ctx.strokeRect(rx2, ry2, rw, rh);
        ctx.fillStyle = '#bae6fd';
        ctx.font = 'bold 16px Pretendard, "Noto Sans KR", sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        if (rw > 70 && rh > 26) ctx.fillText('탭하여 공개', rx2 + rw / 2, ry2 + rh / 2);
      }
    } else if (o.type === 'stamp') {
      var rad = Math.round(o.radius || 22);
      var sx = Math.round(o.x);
      var sy = Math.round(o.y);
      ctx.fillStyle = o.color || '#ef4444';
      ctx.strokeStyle = '#ffffff';
      ctx.lineWidth = 2.6;
      if (o.shape === 'square') {
        // 사각 번호 스탬프 (정사각형 1:1 비율)
        var boxX = sx - rad;
        var boxY = sy - rad;
        var side = rad * 2;
        ctx.beginPath();
        if (ctx.roundRect) ctx.roundRect(boxX, boxY, side, side, 5);
        else ctx.rect(boxX, boxY, side, side);
        ctx.fill();
        ctx.stroke();
      } else {
        // 완벽한 정원(1:1 원형) 번호 스탬프
        ctx.beginPath();
        ctx.arc(sx, sy, rad, 0, Math.PI * 2, false);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
      }
      var numStr = String(o.num || 1);
      ctx.fillStyle = (o.color === '#facc15' || o.color === '#f8fafc') ? '#0f172a' : '#ffffff';
      var fontPx = Math.round(rad * (numStr.length >= 2 ? 0.96 : 1.12));
      ctx.font = '900 ' + fontPx + 'px Pretendard, "Noto Sans KR", sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'alphabetic';
      var m = ctx.measureText(numStr);
      var asc = (typeof m.actualBoundingBoxAscent === 'number' && m.actualBoundingBoxAscent > 0)
        ? m.actualBoundingBoxAscent
        : fontPx * 0.72;
      var desc = (typeof m.actualBoundingBoxDescent === 'number')
        ? m.actualBoundingBoxDescent
        : fontPx * 0.08;
      var textY = sy + (asc - desc) / 2;
      ctx.fillText(numStr, sx, textY);
    } else if (o.type === 'text') {
      var fs = o.fontSize || 26;
      ctx.fillStyle = o.color || '#f8fafc';
      ctx.font = 'bold ' + fs + 'px Pretendard, "Noto Sans KR", "Malgun Gothic", sans-serif';
      ctx.textBaseline = 'top';
      String(o.text || '').split('\n').forEach(function (ln, li) {
        ctx.fillText(ln, o.x, o.y - fs + li * fs * 1.32);
      });
    }
    ctx.restore();
  }

  function renderInk() {
    var pg = curPage();
    inkCtx.clearRect(0, 0, pg.width, pg.height);
    pg.objects.forEach(function (o) { drawInkObject(inkCtx, o); });
    renderSelectionOverlay();
    if (el('sbStampNum')) el('sbStampNum').textContent = nextStampNumber();
  }

  // 다음 번호 스탬프 — 현재 페이지(슬라이드)의 현재 묶음에서 아직 쓰지 않은 가장 작은 번호.
  // 되돌리기·지우개·선택 삭제로 빠진 번호는 다시 채우며, 페이지가 바뀌면 자동으로 1번부터 새로 시작합니다.
  function getPageStampSeries(pg) {
    if (!pg) return 0;
    if (typeof pg.stampSeries !== 'number') pg.stampSeries = 0;
    (pg.objects || []).forEach(function (o) {
      if (o.type === 'stamp' && (o.series || 0) > pg.stampSeries) {
        pg.stampSeries = o.series;
      }
    });
    return pg.stampSeries;
  }

  function nextStampNumber() {
    var pg = curPage();
    var series = getPageStampSeries(pg);
    var used = {};
    (pg.objects || []).forEach(function (o) {
      if (o.type === 'stamp' && (o.series || 0) === series) {
        used[o.num] = true;
      }
    });
    var n = 1;
    while (used[n]) n++;
    return n;
  }

  function resetStampNumber(silent, customMsg) {
    var pg = curPage();
    var curSeries = getPageStampSeries(pg);
    var hasInSeries = (pg.objects || []).some(function (o) {
      return o.type === 'stamp' && (o.series || 0) === curSeries;
    });
    if (hasInSeries) {
      pg.stampSeries = curSeries + 1;
    }
    state.stampSeries = pg.stampSeries;
    if (el('sbStampNum')) el('sbStampNum').textContent = '1';
    if (!silent) {
      showToast(customMsg || '번호 스탬프를 1번부터 다시 시작합니다. (SS · 0)');
    }
  }

  function renderSelectionOverlay() {
    var pg = curPage();
    liveCtx.clearRect(0, 0, pg.width, pg.height);
    if ((state.mode === 'edit' || state.tool === 'select') && state.selection) {
      var s = state.selection;
      liveCtx.save();
      liveCtx.fillStyle = 'rgba(124, 58, 237, 0.14)';
      liveCtx.fillRect(s.x, s.y, s.w, s.h);
      liveCtx.strokeStyle = '#a78bfa';
      liveCtx.lineWidth = 2.5;
      liveCtx.setLineDash([8, 5]);
      liveCtx.strokeRect(s.x, s.y, s.w, s.h);
      liveCtx.restore();
    }
  }

  // ===================================================================
  // 레이저 포인터 (획 단위 독립 분리 + 각 획 3초간 선명하게 유지 후 부드럽게 페이드아웃)
  // ===================================================================
  // 레이저 유지 시간: 1·2·3·5·10초 또는 0(계속 유지 — 도구를 바꾸거나 [레이저 지우기]·Esc를 누를 때까지)
  var LASER_HOLD_OPTIONS = [1000, 2000, 3000, 5000, 10000, 0];
  var LASER_HOLD_MS = 3000;
  try { var _lh = parseInt(localStorage.getItem('sbLaserHold_v1'), 10); if (LASER_HOLD_OPTIONS.indexOf(_lh) >= 0) LASER_HOLD_MS = _lh; } catch (e) {}
  var LASER_FADE_MS = 600;
  function laserHoldLabel(ms) { return ms ? (ms / 1000) + '초' : '계속 유지'; }
  function setLaserHold(ms, quiet) {
    LASER_HOLD_MS = ms;
    try { localStorage.setItem('sbLaserHold_v1', String(ms)); } catch (e) {}
    // 이미 화면에 남아 있는 레이저 획도 새 시간 기준으로 사라지도록 끝난 시각을 지금으로 맞춥니다.
    var now = Date.now();
    (state.laserStrokes || []).forEach(function (s) { if (!s.active) s.endT = now; });
    syncLaserHoldUi();
    if (!quiet) showToast('레이저 유지 시간: ' + laserHoldLabel(ms) + (ms ? '' : ' (도구를 바꾸거나 Esc·[레이저 지우기]로 지움)'));
  }
  function syncLaserHoldUi() {
    document.querySelectorAll('[data-laser-hold]').forEach(function (b) { b.classList.toggle('active', +b.getAttribute('data-laser-hold') === LASER_HOLD_MS); });
    var lb = el('sbLaserHoldBadge'); if (lb) lb.textContent = LASER_HOLD_MS ? (LASER_HOLD_MS / 1000) + '초' : '∞';
    var btn = document.querySelector('.sb-tool-laser-btn');
    if (btn) btn.title = '레이저 포인터 (D) — 그린 선이 ' + (LASER_HOLD_MS ? (LASER_HOLD_MS / 1000) + '초 뒤 사라집니다' : '지울 때까지 남습니다') + '. 다시 누르거나 ▾를 누르면 유지 시간 선택 (Shift+D: 시간 바꾸기)';
  }
  function cycleLaserHold() {
    var i = LASER_HOLD_OPTIONS.indexOf(LASER_HOLD_MS);
    setLaserHold(LASER_HOLD_OPTIONS[(i + 1) % LASER_HOLD_OPTIONS.length]);
  }
  function clearLaserStrokes() {
    state.laserStrokes = (state.laserStrokes || []).filter(function (s) { return s.active; });
    if (!state.laserRaf) state.laserRaf = requestAnimationFrame(renderLaserLoop);
  }

  function renderLaserLoop() {
    var now = Date.now();
    state.laserStrokes = (state.laserStrokes || []).filter(function (s) {
      if (!s || !s.points || !s.points.length) return false;
      if (s.active) return true;
      if (!LASER_HOLD_MS) return true;
      return (now - (s.endT || s.lastT || now)) < LASER_HOLD_MS + LASER_FADE_MS;
    });
    var pg = curPage();
    liveCtx.clearRect(0, 0, pg.width, pg.height);
    if ((state.mode === 'edit' || state.tool === 'select') && state.selection) {
      renderSelectionOverlay();
    }
    if (state.laserStrokes.length >= 1) {
      state.laserStrokes.forEach(function (s) {
        var pts = s.points;
        var age = s.active ? 0 : Math.max(0, now - (s.endT || s.lastT || now));
        var alpha = !LASER_HOLD_MS ? 0.96 : age < LASER_HOLD_MS
          ? (0.96 - (age / LASER_HOLD_MS) * 0.08)
          : Math.max(0, 0.88 * (1 - (age - LASER_HOLD_MS) / LASER_FADE_MS));
        if (alpha <= 0.01) return;

        liveCtx.save();
        liveCtx.lineCap = 'round';
        liveCtx.lineJoin = 'round';
        liveCtx.shadowColor = '#ef4444';
        liveCtx.shadowBlur = 13;
        liveCtx.strokeStyle = 'rgba(239, 68, 68, ' + alpha + ')';
        liveCtx.lineWidth = 4.6 * alpha + 2.2;

        if (pts.length === 1) {
          liveCtx.fillStyle = 'rgba(239, 68, 68, ' + alpha + ')';
          liveCtx.beginPath();
          liveCtx.arc(pts[0].x, pts[0].y, 4.5, 0, Math.PI * 2);
          liveCtx.fill();
        } else {
          liveCtx.beginPath();
          liveCtx.moveTo(pts[0].x, pts[0].y);
          for (var i = 1; i < pts.length; i++) {
            liveCtx.lineTo(pts[i].x, pts[i].y);
          }
          liveCtx.stroke();

          // 중심부 밝은 코어 선을 한 겹 더해 여러 획이 겹쳐도 선명하게 보이도록 표현
          liveCtx.shadowBlur = 0;
          liveCtx.strokeStyle = 'rgba(255, 235, 235, ' + (alpha * 0.72) + ')';
          liveCtx.lineWidth = Math.max(1.4, 2.0 * alpha);
          liveCtx.stroke();
        }

        // 현재 그리고 있는 획 또는 방금 끝난 획 끝에 포인터 광점 표시
        if (s.active || age < 350) {
          var head = pts[pts.length - 1];
          liveCtx.globalAlpha = s.active ? 1 : Math.max(0, 1 - age / 350);
          liveCtx.fillStyle = '#ffffff';
          liveCtx.shadowColor = '#ef4444';
          liveCtx.shadowBlur = 16;
          liveCtx.beginPath();
          liveCtx.arc(head.x, head.y, 5.5, 0, Math.PI * 2);
          liveCtx.fill();
        }
        liveCtx.restore();
      });
      state.laserRaf = requestAnimationFrame(renderLaserLoop);
    } else {
      state.laserRaf = null;
    }
  }

  // ===================================================================
  // 게시 모드 (스닙페이스트 스타일)
  // ===================================================================
  function createCompositeCanvas(pageObj) {
    var pg = pageObj || curPage();
    var cv = document.createElement('canvas');
    cv.width = pg.width;
    cv.height = pg.height;
    var c = cv.getContext('2d');
    if (pg === curPage()) {
      c.drawImage(bgCanvas, 0, 0);
      c.drawImage(inkCanvas, 0, 0);
    } else {
      if (pg.kind === 'blank') drawBlankBg(c, pg.width, pg.height, pg.bgMode || state.bgMode);
      else if (pg.kind === 'image' && pg.imageObj) c.drawImage(pg.imageObj, 0, 0, pg.width, pg.height);
      else if (pg.kind === 'pdf' && state.pdfCache[pg.pdfPageNum]) c.drawImage(state.pdfCache[pg.pdfPageNum], 0, 0, pg.width, pg.height);
      else if (pg.kind === 'pptx' && pg.pptxSlide) renderPptxSlideToCtx(pg.pptxSlide, state.pptxShowAll ? 9999 : (pg.pptxStep || 0), c, pg.width, pg.height);
      pg.objects.forEach(function (o) { drawInkObject(c, o); });
    }
    return cv;
  }

  function createPinFromBoardRect(rx, ry, rw, rh) {
    if (rw < 18 || rh < 18) return;
    // A cross-origin web frame is not part of either canvas. Never create a
    // misleading dark placeholder and claim that the web material was captured.
    if (state.webActive) {
      showToast('웹 자료 화면은 현재 조각에 담을 수 없습니다. PDF·PPTX 또는 캡처 이미지로 연 뒤 고정해 주세요.');
      return;
    }
    var inkOnly = curPage().kind === 'blank';
    // A plain board pin is an overlay of handwriting, not an opaque board tile.
    // File-backed pages retain their source so the selected document stays readable.
    var comp = inkOnly ? inkCanvas : createCompositeCanvas();
    var sx = Math.max(0, Math.round(rx));
    var sy = Math.max(0, Math.round(ry));
    var sw = Math.min(comp.width - sx, Math.round(rw));
    var sh = Math.min(comp.height - sy, Math.round(rh));
    if (sw < 16 || sh < 16) return;

    var cut = document.createElement('canvas');
    cut.width = sw;
    cut.height = sh;
    var cutCtx = cut.getContext('2d');
    cutCtx.drawImage(comp, sx, sy, sw, sh, 0, 0, sw, sh);
    if (inkOnly) {
      var pixels = cutCtx.getImageData(0, 0, sw, sh).data;
      var hasInk = false;
      for (var ai = 3; ai < pixels.length; ai += 4) {
        if (pixels[ai] > 0) { hasInk = true; break; }
      }
      if (!hasInk) {
        showToast('선택한 영역에 고정할 필기가 없습니다. 필기 영역을 포함해 다시 선택해 주세요.');
        return;
      }
    }
    var dataUrl = cut.toDataURL('image/png');

    var pin = document.createElement('div');
    pin.className = 'sb-pin-item';
    var dispW = Math.min(440, Math.max(120, Math.round(sw * state.baseScale * state.zoomScale)));
    var dispH = Math.round(dispW * (sh / sw));
    var fr = boardFrame.getBoundingClientRect();
    var left = Math.max(12, Math.min(window.innerWidth - dispW - 20, fr.left + (sx / comp.width) * fr.width));
    var top = Math.max(12, Math.min(window.innerHeight - dispH - 40, fr.top + (sy / comp.height) * fr.height));
    pin.style.left = left + 'px';
    pin.style.top = top + 'px';
    pin.style.width = dispW + 'px';
    pin.style.height = dispH + 'px';

    var img = document.createElement('img');
    img.src = dataUrl;
    img.alt = '게시 조각';
    var closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'sb-pin-close';
    closeBtn.innerHTML = '&times;';
    closeBtn.addEventListener('click', function (e) { e.stopPropagation(); pin.remove(); });

    var resHandle = document.createElement('div');
    resHandle.className = 'sb-pin-resize';

    pin.appendChild(img);
    pin.appendChild(closeBtn);
    pin.appendChild(resHandle);
    pinLayer.appendChild(pin);

    pin.addEventListener('contextmenu', function (e) {
      e.preventDefault();
      pin.remove();
      showToast('게시 조각을 삭제했습니다.');
    });

    pin.addEventListener('wheel', function (e) {
      e.preventDefault();
      var ratio = e.deltaY < 0 ? 1.09 : 0.91;
      var curW = parseFloat(pin.style.width);
      var newW = Math.max(90, Math.min(window.innerWidth * 0.85, curW * ratio));
      pin.style.width = Math.round(newW) + 'px';
      pin.style.height = Math.round(newW * (sh / sw)) + 'px';
    }, { passive: false });

    pin.addEventListener('pointerdown', function (e) {
      if (e.target === closeBtn || e.target === resHandle) return;
      e.stopPropagation();
      pin.setPointerCapture(e.pointerId);
      var ox = e.clientX - parseFloat(pin.style.left);
      var oy = e.clientY - parseFloat(pin.style.top);
      function onMove(mv) {
        pin.style.left = Math.max(0, Math.min(window.innerWidth - 40, mv.clientX - ox)) + 'px';
        pin.style.top = Math.max(0, Math.min(window.innerHeight - 40, mv.clientY - oy)) + 'px';
      }
      function onUp() {
        pin.removeEventListener('pointermove', onMove);
        pin.removeEventListener('pointerup', onUp);
      }
      pin.addEventListener('pointermove', onMove);
      pin.addEventListener('pointerup', onUp);
    });

    resHandle.addEventListener('pointerdown', function (e) {
      e.stopPropagation();
      resHandle.setPointerCapture(e.pointerId);
      var startX = e.clientX;
      var startW = parseFloat(pin.style.width);
      function onMove(mv) {
        var nw = Math.max(90, Math.min(window.innerWidth * 0.85, startW + (mv.clientX - startX)));
        pin.style.width = Math.round(nw) + 'px';
        pin.style.height = Math.round(nw * (sh / sw)) + 'px';
      }
      function onUp() {
        resHandle.removeEventListener('pointermove', onMove);
        resHandle.removeEventListener('pointerup', onUp);
      }
      resHandle.addEventListener('pointermove', onMove);
      resHandle.addEventListener('pointerup', onUp);
    });

    showToast(inkOnly ? '칠판 배경을 제외한 필기 조각을 고정했습니다. 뒤쪽 글씨가 비쳐 보입니다.' : '자료 영역을 게시 조각으로 고정했습니다. (페이지를 넘겨도 유지됩니다)');
  }

  // ===================================================================
  // 내장 웹 브라우저 + 유튜브 임베드 + 원형/사각 스탬프/판서 레이어
  // ===================================================================
  function convertYoutubeToEmbedUrl(rawUrl) {
    var s = String(rawUrl || '').trim();
    var m = s.match(/(?:youtube\.com\/(?:watch\?.*?v=|embed\/|shorts\/|live\/)|youtu\.be\/)([A-Za-z0-9_-]{11})/i);
    if (m && m[1]) {
      return 'https://www.youtube-nocookie.com/embed/' + m[1] + '?rel=0&modestbranding=1';
    }
    return null;
  }

  function convertGoogleDriveToEmbedUrl(rawUrl) {
    var s = String(rawUrl || '').trim();
    if (!s) return null;
    var mFile = s.match(/drive\.google\.com\/file\/d\/([A-Za-z0-9_-]+)/i);
    if (mFile && mFile[1]) return 'https://drive.google.com/file/d/' + mFile[1] + '/preview';
    var mOpen = s.match(/drive\.google\.com\/(?:open|uc)\?[^#]*\bid=([A-Za-z0-9_-]+)/i);
    if (mOpen && mOpen[1]) return 'https://drive.google.com/file/d/' + mOpen[1] + '/preview';
    var publishedSlides = s.match(/^https:\/\/docs\.google\.com\/presentation\/d\/e\/([A-Za-z0-9_-]+)(?:\/|$)/i);
    if (publishedSlides) return 'https://docs.google.com/presentation/d/e/' + publishedSlides[1] + '/embed?start=false&loop=false';
    var mSlides = s.match(/docs\.google\.com\/presentation\/d\/([A-Za-z0-9_-]+)/i);
    if (mSlides && mSlides[1]) return 'https://docs.google.com/presentation/d/' + mSlides[1] + '/preview';
    var mDocs = s.match(/docs\.google\.com\/document\/d\/([A-Za-z0-9_-]+)/i);
    if (mDocs && mDocs[1]) return 'https://docs.google.com/document/d/' + mDocs[1] + '/preview';
    var mSheets = s.match(/docs\.google\.com\/spreadsheets\/d\/([A-Za-z0-9_-]+)/i);
    if (mSheets && mSheets[1]) return 'https://docs.google.com/spreadsheets/d/' + mSheets[1] + '/preview';
    return null;
  }

  function syncWebQuickToolsUi() {
    document.querySelectorAll('.sb-web-qtool[data-web-qtool]').forEach(function (b) {
      var qt = b.getAttribute('data-web-qtool');
      var isAct = state.webInteract ? (qt === 'interact') : (qt === state.tool);
      b.classList.toggle('active', isAct);
    });
  }

  function setWebInteractMode(interactOn) {
    state.webInteract = !!interactOn;
    boardFrame.classList.toggle('web-interact-on', state.webInteract);
    syncWebQuickToolsUi();
  }

  if (typeof ResizeObserver !== 'undefined') {
    var webLayoutObserver = new ResizeObserver(function () { if (state.webActive) fitViewport(); });
    if (webBar) webLayoutObserver.observe(webBar);
    if (el('sbPalette')) webLayoutObserver.observe(el('sbPalette'));
  }

  function openWebBrowser(urlStr) {
    var raw = String(urlStr || (el('sbWebUrlInput') && el('sbWebUrlInput').value) || '').trim();
    if (!raw) raw = 'https://ko.m.wikipedia.org/wiki/%EC%A1%B0%EA%B1%B4%EB%B6%80_%ED%99%95%EB%A5%A0';
    // Search portals forbid third-party framing. Keep the board and use a normal tab.
    if (/^https?:\/\/(?:www\.)?(?:google\.com|naver\.com|daum\.net)(?:\/|$)/i.test(raw)) {
      window.open(raw, '_blank', 'noopener,noreferrer');
      showToast('검색 사이트는 새 탭에서 열립니다. 이 탭으로 돌아오면 판서를 계속할 수 있습니다.');
      return;
    }

    var ytEmbed = convertYoutubeToEmbedUrl(raw);
    var driveEmbed = ytEmbed ? null : convertGoogleDriveToEmbedUrl(raw);
    var targetUrl = ytEmbed || driveEmbed || raw;
    if (!ytEmbed && !driveEmbed && !/^https?:\/\//i.test(raw)) {
      if (/^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}(\/.*)?$/.test(raw)) {
        targetUrl = 'https://' + raw;
      } else {
        targetUrl = 'https://ko.m.wikipedia.org/w/index.php?search=' + encodeURIComponent(raw);
      }
    }
    if (el('sbWebUrlInput')) el('sbWebUrlInput').value = targetUrl;
    state.webActive = true;
    state.webSlides = /^https:\/\/docs\.google\.com\/presentation\//i.test(targetUrl);
    document.body.classList.add('sb-web-active');
    if (el('sbSlidesHelp')) el('sbSlidesHelp').style.display = state.webSlides ? 'block' : 'none';
    if (el('sbWebInteractToggleBtn')) el('sbWebInteractToggleBtn').textContent = state.webSlides ? '슬라이드 조작' : '웹 조작';
    resetZoom();
    if (webBar) webBar.style.display = 'flex';
    if (webFrame) {
      webFrame.style.display = 'block';
      webFrame.src = targetUrl;
    }
    fitViewport();
    // 유튜브·구글 드라이브는 바로 재생/페이지 넘김할 수 있도록 웹 조작 모드로, 일반 웹은 현재 도구 유지
    if (ytEmbed) {
      setWebInteractMode(true);
      showToast('유튜브 영상을 보드 안에서 열었습니다. 재생 후 하단 [① 원형스탬프] 또는 [펜 필기]로 영상 위에 필기하세요.');
    } else if (driveEmbed) {
      setWebInteractMode(true);
      showToast('구글 드라이브 자료를 보드 안에서 열었습니다. 페이지를 맞춘 뒤 하단 [펜 필기]·[스탬프]로 바로 판서하세요.');
    } else {
      setWebInteractMode(false);
      showToast('내장 웹 브라우저를 열었습니다. 하단 [① 원형스탬프]·[▣ 사각스탬프]·[펜 필기]·[웹 조작]을 바로 전환하세요.');
    }
  }

  function closeWebBrowser() {
    state.webActive = false;
    state.webSlides = false;
    document.body.classList.remove('sb-web-active');
    setWebInteractMode(false);
    if (webBar) webBar.style.display = 'none';
    if (webFrame) {
      webFrame.style.display = 'none';
      webFrame.src = 'about:blank';
    }
    fitViewport();
  }

  async function snapshotWebToBoard() {
    var url = (el('sbWebUrlInput') && el('sbWebUrlInput').value.trim()) || 'https://ko.m.wikipedia.org';
    showToast('웹 화면 스냅샷을 칠판 배경으로 불러오는 중...');
    var shotUrl = 'https://s.wordpress.com/mshots/v1/' + encodeURIComponent(url) + '?w=1280&h=720';
    var im = await loadImgElement(shotUrl);
    var cv = document.createElement('canvas');
    cv.width = 1280; cv.height = 720;
    var g = cv.getContext('2d');
    g.fillStyle = '#ffffff';
    g.fillRect(0, 0, 1280, 720);
    if (im && im.naturalWidth > 100) {
      g.drawImage(im, 0, 0, 1280, 720);
    } else {
      g.fillStyle = '#0f172a';
      g.fillRect(0, 0, 1280, 56);
      g.fillStyle = '#38bdf8';
      g.font = 'bold 20px Pretendard, sans-serif';
      g.fillText('웹 스냅샷 보드: ' + url.slice(0, 75), 24, 35);
      g.fillStyle = '#334155';
      g.font = 'bold 24px Pretendard, sans-serif';
      g.fillText('외부 사이트 스냅샷 페이지 (' + url.slice(0, 60) + ')', 60, 160);
    }
    var curObjs = curPage().objects.slice();
    closeWebBrowser();
    var dataUrl = cv.toDataURL('image/png');
    var finalImg = await loadImgElement(dataUrl);
    state.pages[state.pageIndex] = {
      kind: 'image',
      label: '웹 스냅샷',
      width: 1280,
      height: 720,
      imageDataUrl: dataUrl,
      imageObj: finalImg,
      objects: curObjs,
      redoStack: []
    };
    resizeBoardToPage();
    await renderBackground();
    renderInk();
    showToast('웹 화면을 칠판 이미지로 고정했습니다. 이제 게시(Ctrl+`)와 PDF 저장에도 모두 포함됩니다.');
  }

  // ===================================================================
  // 모드 & 도구 전환
  // ===================================================================
  // 확대 전 모드·도구를 기억했다가, 확대를 끄면(Ctrl+1 한 번 더) 그대로 되돌립니다.
  function restorePreZoom() {
    var pz = state.preZoom;
    state.preZoom = null;
    if (pz && pz.tool && pz.tool !== state.tool) setTool(pz.tool);
    setMode(pz && (pz.mode === 'draw' || pz.mode === 'edit') ? pz.mode : 'draw');
  }
  function setMode(nextMode) {
    if (nextMode === 'zoom' && state.mode !== 'zoom' && state.zoomScale <= 1.02) state.preZoom = { mode: state.mode, tool: state.tool };
    commitInlineText();
    if (nextMode !== 'edit') clearSelection();
    state.mode = nextMode;
    if (state.webActive && state.webInteract) setWebInteractMode(false);
    document.querySelectorAll('.sb-mode-btn').forEach(function (b) {
      b.classList.toggle('active', b.getAttribute('data-sb-mode') === nextMode);
    });
    if (nextMode === 'zoom') {
      if (state.zoomScale < 1.3) {
        var zg0 = zoomGeometry();
        var pg0 = curPage();
        // 세로형 PDF·문서(가로 여백이 넓은 문서)는 가로 폭에 딱 맞는 풀사이즈 배율로 정중앙 정렬하여 좌우 밀림을 방지합니다.
        var isPortraitDoc = (pg0.kind === 'pdf' || pg0.kind === 'image') && zg0.fw > 0 && zg0.fw < zg0.sr.width * 0.88;
        var initZoom = isPortraitDoc
          ? Math.max(1.8, Math.min(3.2, Math.round(((zg0.sr.width * 0.98) / zg0.fw) * 100) / 100))
          : 2.0;
        var fx = isPortraitDoc
          ? (zg0.left0 + zg0.fw / 2)
          : (pointer.lastClientX != null ? pointer.lastClientX : (zg0.left0 + zg0.fw / 2));
        var fy = pointer.lastClientY != null ? pointer.lastClientY : (isPortraitDoc ? zg0.top0 : (zg0.top0 + zg0.fh / 2));
        zoomAt(fx, fy, initZoom);
      } else applyTransform();
      liveCanvas.style.cursor = 'zoom-in';
      showToast('확대 모드: 커서를 움직여 볼 곳을 고르고(휠·두 손가락으로 조절) 누르면 그 자리에서 판서합니다. 끄기는 Ctrl+1');
    } else if (nextMode === 'pin') {
      liveCanvas.style.cursor = 'crosshair';
      showToast('게시 모드(스닙페이스트): 화면에 고정할 영역을 드래그하세요.');
    } else if (nextMode === 'edit') {
      liveCanvas.style.cursor = 'crosshair';
      showToast('편집 모드: 이동하거나 글씨로 변환(OCR)할 판서 영역을 드래그하세요.');
    } else {
      liveCanvas.style.cursor = state.tool === 'text' ? 'text' : 'crosshair';
    }
    if (nextMode !== 'zoom') applyTransform();   // 확대 배지 글자(확대 모드 → 확대 고정) 갱신
  }

  // 단축키 설정에서 바꾼 키를 버튼 글자에 보여 주기 위한 함수(단축키 설정 코드가 실제 내용으로 바꿔 끼웁니다)
  var scDisplayKey = function (defKey) { return String(defKey || '').toUpperCase(); };
  function setKbd(kb, defKey, prefix) {
    if (!kb) return;
    kb.setAttribute('data-sc-def', defKey);
    kb.setAttribute('data-sc-pre', prefix || ''); kb.setAttribute('data-sc-post', '');
    var shown = scDisplayKey(defKey);
    kb.textContent = shown ? (prefix || '') + shown : '';
    kb.hidden = !shown;
  }

  function setTool(nextTool) {
    commitInlineText();
    var prevTool = state.tool;
    // 지우개·레이저·선택·이동(손)은 잠깐 쓰는 도구라, 다시 누르면 돌아갈 '마지막 그리기 도구'로 기억하지 않습니다.
    if (['eraser', 'laser', 'select', 'hand'].indexOf(nextTool) < 0) {
      state.lastDrawTool = nextTool;
    }
    if (nextTool !== 'select') clearSelection();
    if (prevTool === 'laser' && nextTool !== 'laser' && !LASER_HOLD_MS && state.laserStrokes && state.laserStrokes.length) clearLaserStrokes();
    state.tool = nextTool;
    // 형광펜 기본 6pt ↔ 일반 펜 기본 2pt 개별 굵기 자동 전환
    if (nextTool === 'stamp') state.lastStampShape = 'circle';
    else if (nextTool === 'sqStamp') state.lastStampShape = 'square';
    if (nextTool === 'highlighter' && prevTool !== 'highlighter') {
      setSize(state.highlighterSize || 6);
    } else if (nextTool === 'pen' && prevTool !== 'pen') {
      setSize(state.penSize || 2);
    }
    if (state.webActive && state.webInteract) setWebInteractMode(false);
    if (state.mode !== 'draw') setMode('draw');
    document.querySelectorAll('.sb-tool-btn[data-sb-tool]').forEach(function (b) {
      b.classList.toggle('active', b.getAttribute('data-sb-tool') === nextTool);
    });
    // 펜 통합 버튼(#sbPenMenuBtn) 활성 상태 및 현재 펜 이름 동기화
    var PEN_STYLE_LABELS = {
      fountain: '만년필', ballpoint: '볼펜', pencil: '연필', brush: '붓펜', highlighter: '형광펜'
    };
    var penBtn = el('sbPenMenuBtn');
    var penLbl = el('sbPenMenuLabel');
    if (penBtn && penLbl) {
      var isPenActive = (nextTool === 'pen' || nextTool === 'highlighter');
      penBtn.classList.toggle('active', isPenActive);
      var curStyleKey = nextTool === 'highlighter' ? 'highlighter' : (state.penStyle || 'fountain');
      penLbl.textContent = PEN_STYLE_LABELS[curStyleKey] || '만년필';
      setKbd(el('sbPenMenuKey'), curStyleKey === 'highlighter' ? 'h' : 'p');
      document.querySelectorAll('.sb-pen-opt[data-pen-style]').forEach(function (opt) {
        opt.classList.toggle('active', opt.getAttribute('data-pen-style') === curStyleKey);
      });
    }
    // 도형·스탬프 드롭다운 메뉴 안의 도구가 선택되면 트리거 버튼에 현재 도구 표시
    var SUB_TOOL_LABELS = {
      line: '직선', arrow: '화살표', arrow2: '양쪽화살표', rect: '사각형', circle: '원', fillRect: '채움사각',
      cover: '가림막', text: '텍스트', stamp: '①스탬프', sqStamp: '▣스탬프'
    };
    var SUB_TOOL_KEYS = { line: 'l', arrow: 'l', arrow2: 'l', rect: 'r', circle: 'c', text: 't', stamp: 's', sqStamp: 'q' };
    var moreBtn = el('sbMoreToolsBtn');
    var moreLbl = el('sbMoreToolsLabel');
    var moreKey = el('sbMoreToolsKey');
    if (moreBtn && moreLbl) {
      if (SUB_TOOL_LABELS[nextTool]) {
        moreBtn.classList.add('has-active');
        moreLbl.textContent = SUB_TOOL_LABELS[nextTool];
      } else {
        moreBtn.classList.remove('has-active');
        moreLbl.textContent = '도형';
      }
      // 지금 고른 도형의 단축키를 버튼 옆에 보여 줍니다(단축키가 없는 도구는 숨김).
      if (moreKey) {
        if (SUB_TOOL_KEYS[nextTool]) setKbd(moreKey, SUB_TOOL_KEYS[nextTool], (nextTool === 'arrow' || nextTool === 'arrow2') ? '⇧' : '');
        else { moreKey.textContent = ''; moreKey.hidden = true; moreKey.removeAttribute('data-sc-def'); }
      }
    }
    syncWebQuickToolsUi();
    liveCanvas.style.cursor = nextTool === 'select' ? 'default' : nextTool === 'hand' ? 'grab' : (nextTool === 'text' ? 'text' : 'crosshair');
    if (el('sbZoomHandBtn')) el('sbZoomHandBtn').classList.toggle('active', nextTool === 'hand');
  }

  // 갤럭시 S펜 옆 버튼(배럴 버튼: button=2 / buttons bit 2, 지우개 버튼: button=5 / buttons bit 32) 감지
  function isSpenSideButtonPressed(ev) {
    if (!ev) return false;
    var btns = typeof ev.buttons === 'number' ? ev.buttons : 0;
    var btn = typeof ev.button === 'number' ? ev.button : -1;
    if ((btns & 32) !== 0 || btn === 5) return true;
    if ((btns & 2) !== 0 || btn === 2) {
      // S펜(pen)이거나 S펜 버튼 모드가 켜져 있을 때 보조 버튼 입력을 S펜 옆 버튼으로 처리
      if (ev.pointerType === 'pen' || state.spenEraserOnBtn) return true;
    }
    return false;
  }

  function activateSpenEraserButton(toggleIfAlreadyEraser, reasonText) {
    if (!state.spenEraserOnBtn) return false;
    var now = Date.now();
    // 한 번의 물리 버튼 클릭으로 hover + pointerdown 이 연속 발생할 때 이중 토글 방지 (280ms 디바운스)
    if (now - (state.lastSpenSwitchTime || 0) < 280) return true;
    state.lastSpenSwitchTime = now;

    // 항상 양방향 스위치(지우개 <-> 이전 펜 도구)로 동작!
    if (state.tool === 'eraser') {
      var prev = state.lastDrawTool || 'pen';
      setTool(prev);
      showToast((reasonText || 'S펜 버튼 스위치') + ': 다시 [' + prev + '] 모드로 전환되었습니다.');
      return true;
    } else {
      setTool('eraser');
      showToast((reasonText || 'S펜 버튼 스위치') + ': [지우개] 모드가 활성화되었습니다. (한 번 더 누르면 펜 복귀)');
      return true;
    }
  }

  function setColor(hex) {
    state.color = hex;
    document.querySelectorAll('#sbColorGroup .sb-color-dot').forEach(function (d) {
      d.classList.toggle('active', d.getAttribute('data-color') === hex);
    });
    if (inlineTextEl.style.display !== 'none') inlineTextEl.style.color = hex;
  }

  var SIZE_STEPS = [1, 1.5, 2, 2.5, 3, 3.5, 4, 5, 6, 8, 10, 14, 18, 24, 30];

  function formatSizePt(v) {
    var r = Math.round(v * 2) / 2;
    return (r % 1 === 0 ? String(r) : r.toFixed(1)) + 'pt';
  }

  function setSize(px) {
    var clamped = Math.max(1, Math.min(30, Math.round(px * 2) / 2));
    state.size = clamped;
    if (state.tool === 'highlighter') state.highlighterSize = clamped;
    else if (state.tool === 'pen') state.penSize = clamped;
    if (el('sbSizeLabel')) el('sbSizeLabel').textContent = formatSizePt(state.size);
    document.querySelectorAll('.sb-size-preset-btn[data-size-preset]').forEach(function (b) {
      var pv = parseFloat(b.getAttribute('data-size-preset'));
      b.classList.toggle('active', Math.abs(pv - state.size) < 0.05);
    });
  }

  function stepSize(dir) {
    var cur = state.size;
    if (dir > 0) {
      for (var i = 0; i < SIZE_STEPS.length; i++) {
        if (SIZE_STEPS[i] > cur + 0.01) { setSize(SIZE_STEPS[i]); return; }
      }
      setSize(SIZE_STEPS[SIZE_STEPS.length - 1]);
    } else {
      for (var j = SIZE_STEPS.length - 1; j >= 0; j--) {
        if (SIZE_STEPS[j] < cur - 0.01) { setSize(SIZE_STEPS[j]); return; }
      }
      setSize(SIZE_STEPS[0]);
    }
  }

  function setPenStyle(styleKey) {
    if (styleKey === 'highlighter') {
      setTool('highlighter');
      return;
    }
    state.penStyle = styleKey || 'fountain';
    setTool('pen');
  }

  function setBlankBgMode(bgMode) {
    state.bgMode = bgMode;
    document.querySelectorAll('[data-sb-bg]').forEach(function (b) {
      b.classList.toggle('active', b.getAttribute('data-sb-bg') === bgMode);
    });
    var pg = curPage();
    if (pg.kind === 'blank') {
      pg.bgMode = bgMode;
      pg.label = bgMode === 'light' ? '백지' : (bgMode === 'grid' ? '모눈' : '칠판');
      renderBackground();
    } else {
      state.pages.push(createBlankPage(bgMode));
      state.pageIndex = state.pages.length - 1;
      resetStampNumber(true);
      resizeBoardToPage();
      renderBackground();
      renderInk();
      showToast('새 ' + curPage().label + ' 페이지를 추가했습니다.');
    }
  }

  function clearSelection() {
    state.selection = null;
    if (editBar) editBar.style.display = 'none';
    renderSelectionOverlay();
  }

  function showEditBar() {
    if (!editBar || !state.selection) return;
    if (el('sbEditCount')) el('sbEditCount').textContent = '판서 ' + state.selection.ids.length + '개 선택됨';
    editBar.style.display = 'flex';
  }

  // ===================================================================
  // 인라인 한글 텍스트 입력 (T)
  // ===================================================================
  var pendingTextPos = null;
  function openInlineText(bx, by) {
    commitInlineText();
    var pg = curPage();
    pendingTextPos = { x: bx, y: by };
    var fs = Math.max(18, state.size * 3.8);
    inlineTextEl.style.left = Math.max(0.5, (bx / (pg.width || 1280)) * 100) + '%';
    inlineTextEl.style.top = Math.max(0.5, ((by - fs * 0.5) / (pg.height || 720)) * 100) + '%';
    inlineTextEl.style.fontSize = Math.max(14, Math.round(fs * state.baseScale)) + 'px';
    inlineTextEl.style.color = state.color;
    inlineTextEl.value = '';
    inlineTextEl.style.display = 'block';
    setTimeout(function () { inlineTextEl.focus(); }, 20);
  }

  function commitInlineText() {
    if (inlineTextEl.style.display === 'none' || !pendingTextPos) return;
    var txt = inlineTextEl.value.trim();
    inlineTextEl.style.display = 'none';
    if (txt) {
      var pg = curPage();
      var fs = Math.max(18, state.size * 3.8);
      pg.objects.push({
        id: uid(),
        type: 'text',
        x: pendingTextPos.x,
        y: pendingTextPos.y + fs * 0.5,
        text: txt,
        color: state.color,
        fontSize: fs
      });
      pg.redoStack = [];
      renderInk();
    }
    pendingTextPos = null;
  }

  // ===================================================================
  // 포인터 입력 처리 (S펜 / 손가락 핀치 / 마우스)
  // ===================================================================
  // 되돌리기용 기록 — 지우개 한 번 문지르기 / 전체 지우기를 한꺼번에 되살립니다. (쪽마다 쌓임)
  var undoSnaps = new WeakMap();
  var redoSnaps = new WeakMap();
  function pushUndoSnap(pg, before, after) {
    var st = undoSnaps.get(pg);
    if (!st) { st = []; undoSnaps.set(pg, st); }
    var entry = { before: before, after: after, afterLen: after.length };
    st.push(entry);
    if (st.length > 30) st.shift();
    redoSnaps.delete(pg);
    return entry;
  }
  function eraseAt(bx, by) {
    var pg = curPage();
    var er = Math.max(12, state.size * 2.6);
    var before = pg.objects.length;
    var orig = pg.objects;
    pg.objects = pg.objects.filter(function (o) {
      if (o.type === 'stroke' || o.type === 'highlighter') {
        return !o.points.some(function (pt) {
          var dx = pt.x - bx, dy = pt.y - by;
          return dx * dx + dy * dy <= er * er;
        });
      }
      var b = getObjBounds(o);
      return !(bx >= b.x - er && bx <= b.x + b.w + er && by >= b.y - er && by <= b.y + b.h + er);
    });
    if (pg.objects.length !== before) {
      var en = pointer.eraseEntry;
      if (en && en.pg === pg && en.entry.after === orig) { en.entry.after = pg.objects; en.entry.afterLen = pg.objects.length; }
      else pointer.eraseEntry = { pg: pg, entry: pushUndoSnap(pg, orig, pg.objects) };
      renderInk();
    } else {
      pg.objects = orig; // 지운 것이 없으면 배열을 그대로 둡니다(되돌리기 기록 유지)
    }
  }

  function hitCoverAt(bx, by) {
    var pg = curPage();
    for (var i = pg.objects.length - 1; i >= 0; i--) {
      var o = pg.objects[i];
      if (o.type === 'cover') {
        var b = getObjBounds(o);
        if (bx >= b.x && bx <= b.x + b.w && by >= b.y && by <= b.y + b.h) return o;
      }
    }
    return null;
  }

  // ===================================================================
  // 그리다가 꾸욱 — 펜으로 그린 뒤 끝에서 잠깐(0.55초) 멈추면 반듯한 도형(직선·삼각형·사각형·원)으로 바꿉니다.
  // ===================================================================
  var SHAPE_HOLD_MS = 550;
  function armShapeHold(pt) {
    var a = pointer.holdAnchor;
    if (a && Math.hypot(pt.x - a.x, pt.y - a.y) <= 6) return;   // 손이 거의 멈춰 있으면 타이머를 그대로 둡니다
    pointer.holdAnchor = { x: pt.x, y: pt.y };
    clearTimeout(pointer.holdTimer);
    pointer.holdTimer = setTimeout(snapHeldStroke, SHAPE_HOLD_MS);
  }
  function snapHeldStroke() {
    if (!pointer.down || !pointer.currentStroke || pointer.snapShape || (pointer.effectiveTool !== 'pen' && pointer.effectiveTool !== 'highlighter')) return;
    var shape = recognizeShape(pointer.currentStroke.points);
    if (!shape) return;
    var isHl = pointer.effectiveTool === 'highlighter' || pointer.currentStroke.type === 'highlighter';
    shape.id = uid();
    shape.color = pointer.currentStroke.color;
    if (isHl) {
      if (shape.type === 'stroke') {
        shape.type = 'highlighter';
        shape.penStyle = 'highlighter';
        shape.size = pointer.currentStroke.size || 6;
      } else {
        shape.isHighlighter = true;
        shape.size = Math.max(12, (pointer.currentStroke.size || 6) * 4.4);
      }
    } else {
      shape.size = Math.max(2, pointer.currentStroke.size);
    }
    pointer.snapShape = shape;
    var pg = curPage();
    liveCtx.clearRect(0, 0, pg.width, pg.height);
    drawInkObject(liveCtx, shape);
  }
  function rdp(points, eps) {
    if (points.length < 3) return points.slice();
    var a = points[0], b = points[points.length - 1], maxD = -1, idx = 0;
    var dx = b.x - a.x, dy = b.y - a.y, len = Math.hypot(dx, dy) || 1;
    for (var i = 1; i < points.length - 1; i++) {
      var d = Math.abs(dy * points[i].x - dx * points[i].y + b.x * a.y - b.y * a.x) / len;
      if (d > maxD) { maxD = d; idx = i; }
    }
    if (maxD <= eps) return [a, b];
    return rdp(points.slice(0, idx + 1), eps).slice(0, -1).concat(rdp(points.slice(idx), eps));
  }
  // 경로를 같은 간격 n개 점으로 다시 뽑습니다(빨리 그은 곳·천천히 그은 곳의 점 개수 차이를 없앰).
  function resamplePath(pts, n, pathLen) {
    var step = pathLen / n, out = [{ x: pts[0].x, y: pts[0].y }], acc = 0;
    for (var i = 1; i < pts.length && out.length < n; i++) {
      var a = pts[i - 1], b = pts[i], d = Math.hypot(b.x - a.x, b.y - a.y);
      while (acc + d >= step && out.length < n) {
        var t = (step - acc) / d;
        a = { x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t };
        out.push(a);
        d = Math.hypot(b.x - a.x, b.y - a.y);
        acc = 0;
      }
      acc += d;
    }
    return out;
  }
  function detectStar(pts, pathLen) {
    var N = 160, rs = resamplePath(pts, N, pathLen);
    if (rs.length < N * 0.8) return null;
    var cx = 0, cy = 0;
    rs.forEach(function (p) { cx += p.x; cy += p.y; });
    cx /= rs.length; cy /= rs.length;
    var raw = rs.map(function (p) { return Math.hypot(p.x - cx, p.y - cy); });
    var L = raw.length, r = raw.map(function (_, i) { return (raw[(i - 2 + L) % L] + raw[(i - 1 + L) % L] + raw[i] + raw[(i + 1) % L] + raw[(i + 2) % L]) / 5; });
    var maxR = Math.max.apply(null, r);
    if (maxR < 20) return null;
    var minIdx = 0;
    r.forEach(function (v, i) { if (v < r[minIdx]) minIdx = i; });
    if (r[minIdx] > 0.6 * maxR) return null;
    var peaks = [], high = false, best = -1;
    for (var k = 0; k <= L; k++) {
      var i = (minIdx + k) % L;
      if (!high && r[i] > 0.75 * maxR) { high = true; best = i; }
      else if (high && r[i] > r[best]) best = i;
      if (high && (r[i] < 0.6 * maxR || k === L)) { high = false; peaks.push(best); }
    }
    if (peaks.length !== 5) return null;
    // 꼭짓점 방향: 가장 위쪽 꼭짓점을 기준으로 반듯한 오각별을 만듭니다.
    var R = 0;
    peaks.forEach(function (pi) { R += Math.hypot(rs[pi].x - cx, rs[pi].y - cy); });
    R /= 5;
    var top = peaks.reduce(function (a, b) { return rs[b].y < rs[a].y ? b : a; });
    var rot = Math.atan2(rs[top].y - cy, rs[top].x - cx);
    var verts = [];
    for (var v = 0; v < 10; v++) {
      var rr = v % 2 ? R * 0.4 : R, ang = rot + v * Math.PI / 5;
      verts.push({ x: cx + rr * Math.cos(ang), y: cy + rr * Math.sin(ang) });
    }
    verts.push(verts[0]);
    var out = [];
    for (var e = 0; e < 10; e++) for (var t = 0; t < 1; t += 0.08) out.push({ x: verts[e].x + (verts[e + 1].x - verts[e].x) * t, y: verts[e].y + (verts[e + 1].y - verts[e].y) * t, p: 0.6 });
    out.push({ x: verts[0].x, y: verts[0].y, p: 0.6 });
    return { type: 'stroke', penStyle: 'ballpoint', points: out };
  }

  function recognizeShape(pts) {
    if (!pts || pts.length < 6) return null;
    var minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity, pathLen = 0;
    pts.forEach(function (p, i) {
      minX = Math.min(minX, p.x); minY = Math.min(minY, p.y); maxX = Math.max(maxX, p.x); maxY = Math.max(maxY, p.y);
      if (i) pathLen += Math.hypot(p.x - pts[i - 1].x, p.y - pts[i - 1].y);
    });
    var w = maxX - minX, h = maxY - minY, diag = Math.hypot(w, h);
    if (diag < 24) return null;
    var s0 = pts[0], s1 = pts[pts.length - 1], chord = Math.hypot(s1.x - s0.x, s1.y - s0.y);
    // 직선: 거의 곧게 그은 선 — 수평·수직에 가까우면 반듯하게 맞춥니다.
    if (chord > 0.9 * pathLen) {
      var x2 = s1.x, y2 = s1.y, ang = Math.abs(Math.atan2(s1.y - s0.y, s1.x - s0.x) * 180 / Math.PI);
      if (ang < 7 || ang > 173) y2 = s0.y; else if (Math.abs(ang - 90) < 7) x2 = s0.x;
      return { type: 'line', x1: s0.x, y1: s0.y, x2: x2, y2: y2 };
    }
    // 닫힌 도형: 시작점과 끝점이 가까울 때만
    if (chord > Math.max(40, 0.3 * diag)) return null;
    // 별: 가운데에서 잰 거리가 '멀다-가깝다'를 다섯 번 되풀이하면 별(한 붓 별 ☆ / 테두리 별 모두)
    var star = detectStar(pts, pathLen);
    if (star) return star;
    // 닫힌 선은 시작점과 가장 먼 점에서 둘로 나눠 꼭짓점을 찾습니다(시작=끝이면 한 번에 계산할 수 없음).
    var far = 0, farD = -1;
    pts.forEach(function (p, i) { var dd = Math.hypot(p.x - s0.x, p.y - s0.y); if (dd > farD) { farD = dd; far = i; } });
    var eps = 0.07 * diag;
    var poly = rdp(pts.slice(0, far + 1), eps).slice(0, -1).concat(rdp(pts.slice(far).concat([s0]), eps));
    var corners = [];
    poly.slice(0, -1).forEach(function (p) {
      if (!corners.some(function (c) { return Math.hypot(c.x - p.x, c.y - p.y) < 0.15 * diag; })) corners.push(p);
    });
    if (corners.length === 3) {
      // 삼각형 — 변을 따라 점을 촘촘히 두어 반듯한 선으로 그립니다.
      var tri = [], v = corners.concat([corners[0]]);
      for (var k = 0; k < 3; k++) for (var t = 0; t < 1; t += 0.05) tri.push({ x: v[k].x + (v[k + 1].x - v[k].x) * t, y: v[k].y + (v[k + 1].y - v[k].y) * t, p: 0.6 });
      tri.push({ x: v[0].x, y: v[0].y, p: 0.6 });
      return { type: 'stroke', penStyle: 'ballpoint', points: tri };
    }
    if (corners.length === 4) return { type: 'rect', x: minX, y: minY, w: w, h: h };
    // 그 밖에는 원·타원 — 가로세로가 비슷하면 정원으로
    var ratio = w / Math.max(1, h);
    if (ratio > 0.82 && ratio < 1.22) {
      var r = (w + h) / 2, cx = (minX + maxX) / 2, cy = (minY + maxY) / 2;
      return { type: 'circle', x: cx - r / 2, y: cy - r / 2, w: r, h: r };
    }
    return { type: 'circle', x: minX, y: minY, w: w, h: h };
  }

  // 터치 제스처 — 두 손가락은 확대·축소와 이동을 함께, 세 손가락 이상은 이동만 합니다.
  // 두 손가락 pointermove 이벤트가 한 프레임에 번갈아 들어오며 파르르 떨리는 현상을 막기 위해
  // requestAnimationFrame 단위 코얼레싱과 순수 이동 시 핀치 배율 고정 데드존을 적용합니다.
  var touchGestureRaf = null;
  function touchPoints() { return Object.keys(pointer.pinchMap).map(function (k) { return pointer.pinchMap[k]; }); }
  function startTouchGesture() {
    if (touchGestureRaf) { cancelAnimationFrame(touchGestureRaf); touchGestureRaf = null; }
    var pts = touchPoints();
    if (!pts.length) { pointer.gesture = null; return; }
    var cx = 0, cy = 0;
    pts.forEach(function (p) { cx += p.x; cy += p.y; });
    cx /= pts.length; cy /= pts.length;
    pointer.gesture = {
      count: pts.length, cx: cx, cy: cy,
      dist: pts.length === 2 ? Math.hypot(pts[1].x - pts[0].x, pts[1].y - pts[0].y) : 0,
      pinchActive: false,
      zoom: state.zoomScale, panX: state.panX, panY: state.panY
    };
  }
  function moveTouchGesture() {
    touchGestureRaf = null;
    var g = pointer.gesture, pts = touchPoints();
    if (!g || !pts.length) return;
    if (pts.length !== g.count) { startTouchGesture(); return; }
    var cx = 0, cy = 0;
    pts.forEach(function (p) { cx += p.x; cy += p.y; });
    cx /= pts.length; cy /= pts.length;
    if (pts.length === 2 && g.dist > 14) {
      var dist = Math.hypot(pts[1].x - pts[0].x, pts[1].y - pts[0].y);
      var distDelta = Math.abs(dist - g.dist);
      // 두 손가락을 나란히 쓸어 넘길 때 미세한 간격 변화로 화면이 떨리지 않도록 데드존(12px 및 7%) 안에서는 배율을 고정합니다.
      if (!g.pinchActive && distDelta >= Math.max(12, g.dist * 0.07)) {
        g.pinchActive = true;
        g.dist = dist;
        g.cx = cx;
        g.cy = cy;
        g.zoom = state.zoomScale;
        g.panX = state.panX;
        g.panY = state.panY;
      }
      if (g.pinchActive) {
        var rawZ = g.zoom * (dist / g.dist);
        var z1 = Math.max(1, Math.min(4, rawZ));
        var geo = zoomGeometry();
        var ux = (g.cx - geo.left0 - g.panX) / g.zoom, uy = (g.cy - geo.top0 - g.panY) / g.zoom;
        state.zoomScale = z1;
        state.panX = cx - geo.left0 - z1 * ux;
        state.panY = cy - geo.top0 - z1 * uy;
      } else {
        state.panX = g.panX + (cx - g.cx);
        state.panY = g.panY + (cy - g.cy);
      }
    } else {
      state.panX = g.panX + (cx - g.cx);
      state.panY = g.panY + (cy - g.cy);
    }
    clampPan();
    applyTransform();
  }
  function scheduleTouchGestureMove() {
    if (!touchGestureRaf) {
      touchGestureRaf = requestAnimationFrame(moveTouchGesture);
    }
  }

  // S펜 자동 — 펜이 화면에 닿거나 가까이 오면 S펜 전용 모드(손바닥 무시·손가락은 확대/이동)를 자동으로 켭니다.
  function setPalmReject(on, msg) {
    state.palmReject = !!on;
    var b = el('sbPalmRejectBtn');
    if (b) { b.classList.toggle('active', state.palmReject); b.textContent = state.palmReject ? 'S펜전용: 켜짐' : 'S펜전용: 꺼짐'; }
    if (msg) showToast(msg);
  }
  function autoDetectSpen(ev) {
    if (ev.pointerType !== 'pen' || !state.spenAuto || state.palmReject || state.spenManualOff) return;
    setPalmReject(true, 'S펜을 감지해 S펜 전용 모드를 자동으로 켰습니다. (손바닥은 무시 · 손가락은 확대·이동)');
  }
  liveCanvas.addEventListener('pointerover', autoDetectSpen);
  liveCanvas.addEventListener('pointerdown', function (ev) {
    autoDetectSpen(ev);
    pointer.eraseEntry = null;
    if (ev.pointerType === 'touch') {
      pointer.pinchMap[ev.pointerId] = { x: ev.clientX, y: ev.clientY, type: ev.pointerType };
      if (Object.keys(pointer.pinchMap).length >= 2) {
        // 두 손가락: 벌리고 오므려 확대·축소하며 이동 / 세 손가락: 화면 이동. 쓰던 획은 버립니다.
        if (pointer.down && pointer.effectiveTool !== 'pan') { liveCtx.clearRect(0, 0, curPage().width, curPage().height); pointer.currentStroke = null; }
        if (pointer.currentLaserStroke) { pointer.currentLaserStroke.active = false; pointer.currentLaserStroke.endT = Date.now(); pointer.currentLaserStroke = null; }
        clearTimeout(pointer.holdTimer); pointer.snapShape = null;
        pointer.down = false;
        startTouchGesture();
        return;
      }
    }

    if (state.palmReject && ev.pointerType === 'touch' && state.mode === 'draw') {
      pointer.down = true;
      pointer.activeId = ev.pointerId;
      pointer.effectiveTool = 'pan';
      pointer.lastX = ev.clientX;
      pointer.lastY = ev.clientY;
      try { liveCanvas.setPointerCapture(ev.pointerId); } catch (e) {}
      return;
    }

    if (state.mode === 'zoom') {
      setMode('draw');
      // 터치로 확대 화면을 쓸어 넘기려는 경우(S펜 전용 모드 또는 손 도구) 첫 터치부터 끊김 없이 바로 이동을 시작합니다.
      if (ev.pointerType === 'touch' && (state.palmReject || state.tool === 'hand')) {
        pointer.down = true;
        pointer.activeId = ev.pointerId;
        pointer.effectiveTool = 'pan';
        pointer.lastX = ev.clientX;
        pointer.lastY = ev.clientY;
        try { liveCanvas.setPointerCapture(ev.pointerId); } catch (e) {}
        return;
      }
      showToast('확대 화면을 고정하고 판서 모드로 전환했습니다.');
      return;
    }

    // 손 도구(M · Move): 끌어서 확대한 화면을 옮깁니다.
    if (state.tool === 'hand') {
      if (state.zoomScale <= 1.02) showToast('확대(Ctrl+1)한 화면에서 끌어서 옮길 수 있습니다.');
      pointer.down = true;
      pointer.activeId = ev.pointerId;
      pointer.effectiveTool = 'pan';
      pointer.lastX = ev.clientX;
      pointer.lastY = ev.clientY;
      liveCanvas.style.cursor = 'grabbing';
      try { liveCanvas.setPointerCapture(ev.pointerId); } catch (e) {}
      return;
    }

    var pt = clientToBoard(ev.clientX, ev.clientY);
    pointer.down = true;
    pointer.activeId = ev.pointerId;
    pointer.startX = pt.x;
    pointer.startY = pt.y;
    pointer.lastX = pt.x;
    pointer.lastY = pt.y;
    pointer.moved = false;
    try { liveCanvas.setPointerCapture(ev.pointerId); } catch (e) {}

    // 자동 숨김이 켜져 있으면 판서 시작 시 상단 팔레트를 자동으로 위로 숨겨 100% 꽉 찬 화면 확보
    if (state.autoHide && el('sbPalette') && !el('sbPalette').classList.contains('collapsed')) {
      togglePalette(true);
    }

    if (state.mode === 'pin') {
      pointer.effectiveTool = 'pin_select';
      return;
    }

    if (state.mode === 'edit' || (state.mode === 'draw' && state.tool === 'select')) {
      pointer.effectiveTool = 'select';
      if (state.selection &&
          pt.x >= state.selection.x && pt.x <= state.selection.x + state.selection.w &&
          pt.y >= state.selection.y && pt.y <= state.selection.y + state.selection.h) {
        pointer.draggingSelection = true;
      } else {
        // 선택·이동 모드에서는 개별 판서/스탬프/도형을 클릭 즉시 선택하고 바로 드래그 이동 가능
        var curObjs = curPage().objects || [];
        var hitObj = null;
        for (var hi = curObjs.length - 1; hi >= 0; hi--) {
          var hb = getObjBounds(curObjs[hi]);
          if (pt.x >= hb.x - 10 && pt.x <= hb.x + hb.w + 10 && pt.y >= hb.y - 10 && pt.y <= hb.y + hb.h + 10) {
            hitObj = curObjs[hi];
            break;
          }
        }
        if (hitObj) {
          var b0 = getObjBounds(hitObj);
          state.selection = { x: b0.x - 6, y: b0.y - 6, w: b0.w + 12, h: b0.h + 12, ids: [hitObj.id] };
          pointer.draggingSelection = true;
          renderSelectionOverlay();
          showEditBar();
        } else {
          pointer.draggingSelection = false;
          clearSelection();
        }
      }
      return;
    }

    var isSpenBtn = isSpenSideButtonPressed(ev);
    if (isSpenBtn && state.spenEraserOnBtn) {
      // 갤럭시 S펜 옆 버튼을 누르면 [지우개] <-> [펜] 양방향 스위치로 전환!
      activateSpenEraserButton(true, 'S펜 옆 버튼 스위치');
    }
    var eff = state.tool;
    if (ev.ctrlKey && ev.shiftKey) eff = 'fillRect';
    else if (ev.ctrlKey) eff = 'rect';
    else if (ev.shiftKey && eff !== 'circle' && eff !== 'arrow' && eff !== 'arrow2') eff = 'line';
    pointer.effectiveTool = eff;

    if (eff === 'text') {
      pointer.down = false;
      openInlineText(pt.x, pt.y);
      return;
    }
    if (eff === 'stamp' || eff === 'sqStamp') {
      pointer.down = false;
      var pg = curPage();
      pg.objects.push({
        id: uid(),
        type: 'stamp',
        shape: eff === 'sqStamp' ? 'square' : 'circle',
        x: Math.round(pt.x),
        y: Math.round(pt.y),
        num: nextStampNumber(),
        series: getPageStampSeries(pg),
        color: state.color,
        radius: Math.max(18, Math.round(state.size * 2.6))
      });
      pg.redoStack = [];
      redoSnaps.delete(pg);
      renderInk();
      return;
    }
    if (eff === 'eraser') {
      eraseAt(pt.x, pt.y);
      return;
    }
    if (eff === 'laser') {
      var lNow = Date.now();
      var newLaserStroke = { points: [{ x: pt.x, y: pt.y, t: lNow }], lastT: lNow, endT: 0, active: true };
      state.laserStrokes.push(newLaserStroke);
      pointer.currentLaserStroke = newLaserStroke;
      if (!state.laserRaf) state.laserRaf = requestAnimationFrame(renderLaserLoop);
      return;
    }
    if (eff === 'pen' || eff === 'highlighter') {
      var initP = (ev.pointerType === 'pen' && typeof ev.pressure === 'number' && ev.pressure > 0)
        ? Math.max(0.15, Math.min(1, ev.pressure))
        : 0.55;
      pointer.lastMoveTime = Date.now();
      pointer.lastPressure = initP;
      pointer.snapShape = null;
      pointer.holdAnchor = null;
      clearTimeout(pointer.holdTimer);
      pointer.currentStroke = {
        id: uid(),
        type: eff === 'highlighter' ? 'highlighter' : 'stroke',
        penStyle: eff === 'highlighter' ? 'highlighter' : (state.penStyle || 'fountain'),
        color: state.color,
        size: state.size,
        points: [{ x: pt.x, y: pt.y, p: initP }]
      };
    }
  });

  liveCanvas.addEventListener('pointermove', function (ev) {
    if (pointer.pinchMap[ev.pointerId]) {
      pointer.pinchMap[ev.pointerId].x = ev.clientX;
      pointer.pinchMap[ev.pointerId].y = ev.clientY;
    }
    if (ev.pointerType !== 'touch') {
      pointer.lastClientX = ev.clientX;
      pointer.lastClientY = ev.clientY;
    }
    if (pointer.gesture && Object.keys(pointer.pinchMap).length >= 2) {
      scheduleTouchGestureMove();
      return;
    }

    if (state.mode === 'zoom' && !pointer.down && ev.pointerType !== 'touch') {
      // 줌잇처럼 마우스·호버 펜 위치가 화면의 어디쯤인지에 맞춰 확대한 판의 같은 부분을 보여 줍니다(가장자리까지 닿음).
      var zg = zoomGeometry(), zrange = panRange(zg, state.zoomScale);
      var fxr = Math.min(1, Math.max(0, (ev.clientX - zg.sr.left) / zg.sr.width));
      var fyr = Math.min(1, Math.max(0, (ev.clientY - zg.sr.top) / zg.sr.height));
      state.panX = roundSubpx(zrange.x.hi - fxr * (zrange.x.hi - zrange.x.lo));
      state.panY = roundSubpx(zrange.y.hi - fyr * (zrange.y.hi - zrange.y.lo));
      applyTransform();
      return;
    }

    // 갤럭시탭 S펜을 화면 가까이 호버링(hover)한 상태에서 펜 옆 버튼을 딸깍 눌렀을 때 [지우개] 버튼 즉시 토글/활성화!
    if (!pointer.down && (ev.pointerType === 'pen' || state.spenEraserOnBtn)) {
      var hoverSidePressed = ((ev.buttons & 2) !== 0 || (ev.buttons & 32) !== 0);
      if (hoverSidePressed && !pointer.lastHoverSideBtn) {
        pointer.lastHoverSideBtn = true;
        activateSpenEraserButton(true, 'S펜 호버 버튼 클릭');
      } else if (!hoverSidePressed && pointer.lastHoverSideBtn) {
        pointer.lastHoverSideBtn = false;
      }
    }

    if (!pointer.down || ev.pointerId !== pointer.activeId) return;

    if (pointer.effectiveTool === 'pan') {
      state.panX += (ev.clientX - pointer.lastX);
      state.panY += (ev.clientY - pointer.lastY);
      pointer.lastX = ev.clientX;
      pointer.lastY = ev.clientY;
      clampPan();
      applyTransform();
      return;
    }

    var pt = clientToBoard(ev.clientX, ev.clientY);
    if (Math.hypot(pt.x - pointer.startX, pt.y - pointer.startY) > 4) pointer.moved = true;
    var curPg = curPage();

    if (state.mode === 'pin') {
      liveCtx.clearRect(0, 0, curPg.width, curPg.height);
      liveCtx.save();
      liveCtx.fillStyle = 'rgba(245, 158, 11, 0.16)';
      liveCtx.strokeStyle = '#f59e0b';
      liveCtx.lineWidth = 3;
      liveCtx.setLineDash([8, 4]);
      var px = Math.min(pointer.startX, pt.x), py = Math.min(pointer.startY, pt.y);
      var pw = Math.abs(pt.x - pointer.startX), ph = Math.abs(pt.y - pointer.startY);
      liveCtx.fillRect(px, py, pw, ph);
      liveCtx.strokeRect(px, py, pw, ph);
      liveCtx.restore();
      return;
    }

    if (state.mode === 'edit' || pointer.effectiveTool === 'select') {
      if (pointer.draggingSelection && state.selection) {
        var dx = pt.x - pointer.lastX;
        var dy = pt.y - pointer.lastY;
        state.selection.x += dx;
        state.selection.y += dy;
        curPg.objects.forEach(function (o) {
          if (state.selection.ids.indexOf(o.id) >= 0) translateObj(o, dx, dy);
        });
        pointer.lastX = pt.x;
        pointer.lastY = pt.y;
        renderInk();
      } else {
        liveCtx.clearRect(0, 0, curPg.width, curPg.height);
        var ex = Math.min(pointer.startX, pt.x), ey = Math.min(pointer.startY, pt.y);
        var ew = Math.abs(pt.x - pointer.startX), eh = Math.abs(pt.y - pointer.startY);
        liveCtx.save();
        liveCtx.fillStyle = 'rgba(124, 58, 237, 0.14)';
        liveCtx.strokeStyle = '#a78bfa';
        liveCtx.lineWidth = 2.5;
        liveCtx.setLineDash([7, 4]);
        liveCtx.fillRect(ex, ey, ew, eh);
        liveCtx.strokeRect(ex, ey, ew, eh);
        liveCtx.restore();
      }
      return;
    }

    var eff = pointer.effectiveTool;
    if (eff === 'eraser') {
      eraseAt(pt.x, pt.y);
    } else if (eff === 'laser') {
      var mNow = Date.now();
      if (!pointer.currentLaserStroke) {
        pointer.currentLaserStroke = { points: [], lastT: mNow, endT: 0, active: true };
        state.laserStrokes.push(pointer.currentLaserStroke);
      }
      pointer.currentLaserStroke.points.push({ x: pt.x, y: pt.y, t: mNow });
      pointer.currentLaserStroke.lastT = mNow;
      if (!state.laserRaf) state.laserRaf = requestAnimationFrame(renderLaserLoop);
    } else if ((eff === 'pen' || eff === 'highlighter') && pointer.currentStroke) {
      if (pointer.snapShape) {
        // 직선으로 변환된 뒤에도 손을 떼기 전까지 끝점 위치·수평/수직 각도를 자유롭게 조절할 수 있도록 유지
        if (pointer.snapShape.type === 'line') {
          var sx0 = pointer.snapShape.x1, sy0 = pointer.snapShape.y1;
          var nx2 = pt.x, ny2 = pt.y;
          var sang = Math.abs(Math.atan2(ny2 - sy0, nx2 - sx0) * 180 / Math.PI);
          if (sang < 7 || sang > 173) ny2 = sy0;
          else if (Math.abs(sang - 90) < 7) nx2 = sx0;
          pointer.snapShape.x2 = nx2;
          pointer.snapShape.y2 = ny2;
          liveCtx.clearRect(0, 0, curPg.width, curPg.height);
          drawInkObject(liveCtx, pointer.snapShape);
        }
        return;
      }
      armShapeHold(pt);
      var nowT = Date.now();
      var dt = Math.max(1, nowT - (pointer.lastMoveTime || nowT));
      var segDist = Math.hypot(pt.x - pointer.lastX, pt.y - pointer.lastY);
      var rawP;
      if (ev.pointerType === 'pen' && typeof ev.pressure === 'number' && ev.pressure > 0) {
        rawP = Math.max(0.15, Math.min(1, ev.pressure));
      } else {
        // 마우스/손가락: 속도 기반 자연스러운 필압 시뮬레이션 (빠를수록 가늘게, 천천히 쓸수록 또렷하게)
        var vel = segDist / dt; // px per ms
        rawP = Math.max(0.22, Math.min(0.95, 0.88 - vel * 0.32));
      }
      var smoothP = (pointer.lastPressure || 0.55) * 0.65 + rawP * 0.35;
      pointer.lastPressure = smoothP;
      pointer.lastMoveTime = nowT;
      pointer.currentStroke.points.push({ x: pt.x, y: pt.y, p: Math.round(smoothP * 100) / 100 });
      liveCtx.clearRect(0, 0, curPg.width, curPg.height);
      drawInkObject(liveCtx, pointer.currentStroke);
    } else if (eff === 'line' || eff === 'arrow' || eff === 'arrow2' || eff === 'rect' || eff === 'fillRect' || eff === 'cover' || eff === 'circle') {
      liveCtx.clearRect(0, 0, curPg.width, curPg.height);
      if (eff === 'line' || eff === 'arrow' || eff === 'arrow2') {
        drawInkObject(liveCtx, { type: 'line', arrow: eff === 'arrow' ? 1 : eff === 'arrow2' ? 2 : 0, x1: pointer.startX, y1: pointer.startY, x2: pt.x, y2: pt.y, color: state.color, size: state.size });
      } else {
        drawInkObject(liveCtx, { type: eff, x: pointer.startX, y: pointer.startY, w: pt.x - pointer.startX, h: pt.y - pointer.startY, color: state.color, size: state.size, revealed: false, perfect: !!ev.shiftKey });
      }
    }
    pointer.lastX = pt.x;
    pointer.lastY = pt.y;
  });

  // 밑줄(직선) · 원형 도형 좌측 상단에 번호 스탬프 자동 부착 헬퍼
  function syncAutoStampUi() {
    var on = !!state.autoStampShape;
    var topBtn = el('sbAutoStampToggleBtn');
    var topLbl = el('sbAutoStampToggleLabel');
    if (topBtn) {
      topBtn.classList.toggle('active', on);
      topBtn.setAttribute('aria-pressed', on ? 'true' : 'false');
    }
    if (topLbl) topLbl.textContent = on ? '자동① ON' : '자동① OFF';
    var menuBtn = el('sbAutoStampMenuBtn');
    if (menuBtn) {
      menuBtn.textContent = '밑줄·원 좌측상단 자동 스탬프: ' + (on ? '켜짐 (A)' : '꺼짐 (A)');
    }
  }

  function toggleAutoStampShape() {
    state.autoStampShape = !state.autoStampShape;
    try { localStorage.setItem('sbAutoStampShape_v1', state.autoStampShape ? 'on' : 'off'); } catch (e) {}
    syncAutoStampUi();
    showToast(state.autoStampShape
      ? '밑줄·원 좌측 상단 자동 스탬프를 켰습니다. (단축키 A)'
      : '밑줄·원 자동 스탬프를 껐습니다. (단축키 A)');
  }

  function pushShapeWithOptionalAutoStamp(pg, shapeObj) {
    if (!pg || !shapeObj) return;
    if (!state.autoStampShape || (shapeObj.type !== 'line' && shapeObj.type !== 'circle')) {
      pg.objects.push(shapeObj);
      return;
    }

    var sRad = 18;
    var sx = 0, sy = 0;
    if (shapeObj.type === 'line') {
      var dx = shapeObj.x2 - shapeObj.x1;
      var dy = shapeObj.y2 - shapeObj.y1;
      if (Math.hypot(dx, dy) < 22) {
        pg.objects.push(shapeObj);
        return;
      }
      if (Math.abs(dx) >= Math.abs(dy)) {
        var leftFirst = shapeObj.x1 <= shapeObj.x2;
        var lx = leftFirst ? shapeObj.x1 : shapeObj.x2;
        var ly = leftFirst ? shapeObj.y1 : shapeObj.y2;
        sx = lx - sRad - 6;
        sy = ly - sRad - 4;
        if (sx < sRad + 4) {
          sx = lx + Math.round(sRad * 0.4);
          sy = ly - sRad - 10;
        }
      } else {
        var topFirst = shapeObj.y1 <= shapeObj.y2;
        var tx = topFirst ? shapeObj.x1 : shapeObj.x2;
        var ty = topFirst ? shapeObj.y1 : shapeObj.y2;
        sx = tx - sRad - 6;
        sy = ty;
      }
    } else if (shapeObj.type === 'circle') {
      var cx = shapeObj.x + shapeObj.w / 2;
      var cy = shapeObj.y + shapeObj.h / 2;
      var rx = Math.max(1, Math.abs(shapeObj.w) / 2);
      var ry = Math.max(1, Math.abs(shapeObj.h) / 2);
      if (shapeObj.perfect) {
        var pr = Math.max(rx, ry);
        rx = pr; ry = pr;
      }
      if (Math.max(rx * 2, ry * 2) < 20) {
        pg.objects.push(shapeObj);
        return;
      }
      sx = cx - rx * 0.78 - sRad * 0.45;
      sy = cy - ry * 0.78 - sRad * 0.45;
    }

    sx = Math.max(sRad + 4, Math.min((pg.width || 1280) - sRad - 4, Math.round(sx)));
    sy = Math.max(sRad + 4, Math.min((pg.height || 720) - sRad - 4, Math.round(sy)));

    var gid = uid();
    shapeObj.groupId = gid;
    pg.objects.push(shapeObj);
    pg.objects.push({
      id: uid(),
      groupId: gid,
      type: 'stamp',
      shape: state.lastStampShape || 'circle',
      x: sx,
      y: sy,
      num: nextStampNumber(),
      series: getPageStampSeries(pg),
      color: shapeObj.color || state.color,
      radius: sRad
    });
  }

  function finishPointer(ev) {
    var wasGesture = !!pointer.pinchMap[ev.pointerId] && Object.keys(pointer.pinchMap).length >= 2;
    if (wasGesture && touchGestureRaf) moveTouchGesture();
    delete pointer.pinchMap[ev.pointerId];
    if (wasGesture) {
      // 남은 손가락으로 다시 기준을 잡습니다. 한 손가락만 남으면 제스처를 끝내고 판서하지 않습니다.
      if (Object.keys(pointer.pinchMap).length >= 2) startTouchGesture();
      else { if (touchGestureRaf) { cancelAnimationFrame(touchGestureRaf); touchGestureRaf = null; } pointer.gesture = null; }
      pointer.down = false;
      return;
    }
    if (pointer.gesture && !Object.keys(pointer.pinchMap).length) pointer.gesture = null;
    if (!pointer.down || ev.pointerId !== pointer.activeId) return;
    pointer.down = false;

    if (pointer.currentLaserStroke) {
      pointer.currentLaserStroke.active = false;
      pointer.currentLaserStroke.endT = Date.now();
      pointer.currentLaserStroke = null;
    }

    if (pointer.effectiveTool === 'pan') { if (state.tool === 'hand') liveCanvas.style.cursor = 'grab'; return; }

    var pt = clientToBoard(ev.clientX, ev.clientY);
    var curPg = curPage();

    if (state.mode === 'pin') {
      liveCtx.clearRect(0, 0, curPg.width, curPg.height);
      var px = Math.min(pointer.startX, pt.x), py = Math.min(pointer.startY, pt.y);
      var pw = Math.abs(pt.x - pointer.startX), ph = Math.abs(pt.y - pointer.startY);
      if (pw >= 18 && ph >= 18) createPinFromBoardRect(px, py, pw, ph);
      setMode('draw');
      return;
    }

    if (state.mode === 'edit' || pointer.effectiveTool === 'select') {
      if (!pointer.draggingSelection) {
        var ex = Math.min(pointer.startX, pt.x), ey = Math.min(pointer.startY, pt.y);
        var ew = Math.abs(pt.x - pointer.startX), eh = Math.abs(pt.y - pointer.startY);
        if (ew >= 12 && eh >= 12) {
          var box = { x: ex, y: ey, w: ew, h: eh };
          var hitIds = curPg.objects.filter(function (o) { return rectsIntersect(box, getObjBounds(o)); }).map(function (o) { return o.id; });
          state.selection = { x: ex, y: ey, w: ew, h: eh, ids: hitIds };
          renderSelectionOverlay();
          showEditBar();
        } else {
          clearSelection();
        }
      }
      pointer.draggingSelection = false;
      return;
    }

    if (!pointer.moved && (pointer.effectiveTool === 'pen' || pointer.effectiveTool === 'cover' || pointer.effectiveTool === 'laser')) {
      var hitCov = hitCoverAt(pt.x, pt.y);
      if (hitCov) {
        hitCov.revealed = !hitCov.revealed;
        pointer.currentStroke = null;
        liveCtx.clearRect(0, 0, curPg.width, curPg.height);
        renderInk();
        return;
      }
    }

    var eff = pointer.effectiveTool;
    clearTimeout(pointer.holdTimer);
    if ((eff === 'pen' || eff === 'highlighter') && pointer.currentStroke) {
      if (pointer.snapShape) {
        pushShapeWithOptionalAutoStamp(curPg, pointer.snapShape);
      } else {
        curPg.objects.push(pointer.currentStroke);
      }
      pointer.snapShape = null;
      curPg.redoStack = [];
      redoSnaps.delete(curPg);
      pointer.currentStroke = null;
      liveCtx.clearRect(0, 0, curPg.width, curPg.height);
      renderInk();
    } else if (eff === 'line' && pointer.moved) {
      pushShapeWithOptionalAutoStamp(curPg, { id: uid(), type: 'line', x1: pointer.startX, y1: pointer.startY, x2: pt.x, y2: pt.y, color: state.color, size: state.size });
      curPg.redoStack = [];
      redoSnaps.delete(curPg);
      liveCtx.clearRect(0, 0, curPg.width, curPg.height);
      renderInk();
    } else if ((eff === 'arrow' || eff === 'arrow2') && pointer.moved) {
      // 화살표는 자동 번호 스탬프를 붙이지 않습니다.
      curPg.objects.push({ id: uid(), type: 'line', arrow: eff === 'arrow' ? 1 : 2, x1: pointer.startX, y1: pointer.startY, x2: pt.x, y2: pt.y, color: state.color, size: Math.max(2, state.size) });
      curPg.redoStack = [];
      redoSnaps.delete(curPg);
      liveCtx.clearRect(0, 0, curPg.width, curPg.height);
      renderInk();
    } else if ((eff === 'rect' || eff === 'fillRect' || eff === 'cover' || eff === 'circle') && pointer.moved) {
      pushShapeWithOptionalAutoStamp(curPg, { id: uid(), type: eff, x: pointer.startX, y: pointer.startY, w: pt.x - pointer.startX, h: pt.y - pointer.startY, color: state.color, size: state.size, revealed: false, perfect: !!ev.shiftKey });
      curPg.redoStack = [];
      redoSnaps.delete(curPg);
      liveCtx.clearRect(0, 0, curPg.width, curPg.height);
      renderInk();
    }
  }

  liveCanvas.addEventListener('pointerup', finishPointer);
  liveCanvas.addEventListener('pointercancel', finishPointer);

  stageWrap.addEventListener('wheel', function (e) {
    if (state.webActive && state.webInteract) return;
    e.preventDefault();
    if (state.mode === 'zoom' || e.ctrlKey) {
      var nextZ = Math.round((state.zoomScale + (e.deltaY < 0 ? 0.2 : -0.2)) * 10) / 10;
      zoomAt(e.clientX, e.clientY, nextZ);
    } else if (state.mode === 'draw') {
      stepSize(e.deltaY < 0 ? 1 : -1);
    }
  }, { passive: false });

  // ===================================================================
  // PDF 불러오기 (pdf.js)
  // ===================================================================
  async function renderPdfPageToOffscreen(pageNum) {
    var page = await state.pdfDoc.getPage(pageNum);
    var baseVp = page.getViewport({ scale: 1 });
    var targetScale = Math.max(1.8, Math.min(2.6, 1600 / Math.max(baseVp.width, 1)));
    var vp = page.getViewport({ scale: targetScale });
    var off = document.createElement('canvas');
    off.width = Math.round(vp.width);
    off.height = Math.round(vp.height);
    await page.render({ canvasContext: off.getContext('2d'), viewport: vp }).promise;
    return off;
  }

  async function loadPdfFile(file) {
    if (typeof pdfjsLib === 'undefined') { showToast('pdf.js 라이브러리를 불러오지 못했습니다.'); return; }
    try {
      closeWebBrowser();
      showToast('PDF 문서를 불러오는 중...');
      if (pdfjsLib.GlobalWorkerOptions) pdfjsLib.GlobalWorkerOptions.workerSrc = './libs/pdf.worker.min.js';
      var buf = await file.arrayBuffer();
      var doc = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
      state.pdfDoc = doc;
      state.pdfCache = {};
      var newPages = [];
      for (var p = 1; p <= doc.numPages; p++) {
        var pg = await doc.getPage(p);
        var vp = pg.getViewport({ scale: 1 });
        var ratio = vp.height / vp.width;
        var bw = 1280;
        var bh = Math.round(bw * ratio);
        newPages.push({
          kind: 'pdf',
          label: file.name,
          width: bw,
          height: bh,
          pdfPageNum: p,
          objects: [],
          redoStack: []
        });
      }
      state.pages = newPages;
      state.pageIndex = 0;
      state.currentNoteId = null;
      forgetLastNote();
      state.currentNoteTitle = file.name.replace(/\.pdf$/i, '');
      updateNoteBadge();
      resetZoom();
      resizeBoardToPage();
      await renderBackground();
      renderInk();
      showToast(file.name + ' (' + doc.numPages + '쪽)을 불러왔습니다.');
    } catch (err) {
      showToast('PDF 열기 오류: ' + (err.message || err));
    }
  }

  // ===================================================================
  // 고호환 PPTX 파서 (그룹 도형 grpSp · 레이아웃 상속 ph · 문단별/도형별 액션)
  // ===================================================================
  function xmlAttr(node, name) { return node ? (node.getAttribute(name) || '') : ''; }
  function qAll(root, localName) {
    if (!root) return [];
    var out = [];
    function walk(n) {
      if (!n || (n.nodeType !== 1 && n.nodeType !== 9)) return;
      // mc:AlternateContent 중복 방지: Fallback 또는 첫 번째 분기만 탐색
      if (n.localName === 'AlternateContent') {
        var fb = Array.from(n.children || []).find(function (c) { return c.localName === 'Fallback'; });
        var ch = fb || (n.children && n.children[0]);
        if (ch) walk(ch);
        return;
      }
      if (n.nodeType === 1 && n.localName === localName) out.push(n);
      var kids = n.children || n.childNodes || [];
      for (var i = 0; i < kids.length; i++) walk(kids[i]);
    }
    walk(root);
    return out;
  }
  function qOne(root, localName) {
    var list = qAll(root, localName);
    return list.length ? list[0] : null;
  }
  function directChild(parent, localName) {
    if (!parent || !parent.children) return null;
    for (var i = 0; i < parent.children.length; i++) {
      if (parent.children[i].localName === localName) return parent.children[i];
    }
    return null;
  }

  var PRESET_COLORS = {
    black: '#000000', white: '#ffffff', red: '#ef4444', green: '#22c55e', blue: '#3b82f6',
    yellow: '#facc15', cyan: '#06b6d4', magenta: '#d946ef', gray: '#64748b', grey: '#64748b',
    dkBlue: '#1e3a8a', ltGray: '#e2e8f0', dkGray: '#334155', orange: '#f97316', purple: '#9333ea'
  };

  var activeThemeMap = {
    bg1: '#ffffff', lt1: '#ffffff', tx1: '#0f172a', dk1: '#0f172a',
    bg2: '#f1f5f9', lt2: '#e2e8f0', tx2: '#1e293b', dk2: '#1e293b',
    accent1: '#2563eb', accent2: '#ea580c', accent3: '#16a34a',
    accent4: '#d97706', accent5: '#0891b2', accent6: '#7c3aed',
    hlink: '#2563eb', folHlink: '#7c3aed'
  };

  function applyColorModifiers(hexStr, clrNode) {
    if (!hexStr || !clrNode || hexStr.charAt(0) !== '#' || hexStr.length < 7) return hexStr;
    var r = parseInt(hexStr.slice(1, 3), 16);
    var g = parseInt(hexStr.slice(3, 5), 16);
    var b = parseInt(hexStr.slice(5, 7), 16);
    if (isNaN(r) || isNaN(g) || isNaN(b)) return hexStr;

    var lumModNode = directChild(clrNode, 'lumMod');
    var lumOffNode = directChild(clrNode, 'lumOff');
    var tintNode = directChild(clrNode, 'tint');
    var shadeNode = directChild(clrNode, 'shade');

    if (lumModNode || lumOffNode) {
      var mod = lumModNode ? (parseInt(xmlAttr(lumModNode, 'val'), 10) / 100000) : 1;
      var off = lumOffNode ? (parseInt(xmlAttr(lumOffNode, 'val'), 10) / 100000) : 0;
      if (!isNaN(mod) && !isNaN(off)) {
        r = Math.round(Math.max(0, Math.min(255, r * mod + 255 * off)));
        g = Math.round(Math.max(0, Math.min(255, g * mod + 255 * off)));
        b = Math.round(Math.max(0, Math.min(255, b * mod + 255 * off)));
      }
    }
    if (tintNode) {
      var tv = parseInt(xmlAttr(tintNode, 'val'), 10) / 100000;
      if (!isNaN(tv)) {
        r = Math.round(r + (255 - r) * (1 - tv));
        g = Math.round(g + (255 - g) * (1 - tv));
        b = Math.round(b + (255 - b) * (1 - tv));
      }
    }
    if (shadeNode) {
      var sv = parseInt(xmlAttr(shadeNode, 'val'), 10) / 100000;
      if (!isNaN(sv)) {
        r = Math.round(r * sv);
        g = Math.round(g * sv);
        b = Math.round(b * sv);
      }
    }
    var toHex = function (n) { var s = Math.max(0, Math.min(255, n)).toString(16); return s.length < 2 ? '0' + s : s; };
    return '#' + toHex(r) + toHex(g) + toHex(b);
  }

  function extractFillColor(node) {
    if (!node) return null;
    if (directChild(node, 'noFill')) return null;
    var grad = node.localName === 'gradFill' ? node : qOne(node, 'gradFill');
    var target = grad ? (qOne(grad, 'gs') || grad) : node;

    var srgb = qOne(target, 'srgbClr');
    if (srgb && xmlAttr(srgb, 'val')) return applyColorModifiers('#' + xmlAttr(srgb, 'val'), srgb);
    var sys = qOne(target, 'sysClr');
    if (sys && (xmlAttr(sys, 'lastClr') || xmlAttr(sys, 'val'))) {
      var sc = xmlAttr(sys, 'lastClr') || (xmlAttr(sys, 'val') === 'windowText' ? '000000' : 'FFFFFF');
      return applyColorModifiers('#' + sc, sys);
    }
    var prst = qOne(target, 'prstClr');
    if (prst && PRESET_COLORS[xmlAttr(prst, 'val')]) return applyColorModifiers(PRESET_COLORS[xmlAttr(prst, 'val')], prst);
    var scheme = qOne(target, 'schemeClr');
    if (scheme) {
      var v = xmlAttr(scheme, 'val');
      var baseHex = activeThemeMap[v] || null;
      return baseHex ? applyColorModifiers(baseHex, scheme) : null;
    }
    return null;
  }

  // 실제 파워포인트 <p:timing> 클릭 시퀀스 파서 (mainSeq + interactiveSeq 트리거 + 문단 pRg + bldP 완벽 지원)
  function parseSlideTiming(doc) {
    var timing = qOne(doc, 'timing');
    if (!timing) return { steps: [], autoTargets: [] };

    // bldLst(<p:bldP spid="3" build="p"/>) 문단 빌드 대상 파악
    var bldParaMap = {};
    qAll(timing, 'bldP').forEach(function (bp) {
      var bSpid = xmlAttr(bp, 'spid');
      var bType = xmlAttr(bp, 'build');
      if (bSpid && (bType === 'p' || bType === 'allAtOnce')) {
        bldParaMap[bSpid] = bType;
      }
    });

    // mainSeq 뿐 아니라 interactiveSeq(특정 도형 클릭 트리거 시퀀스)도 모두 수집
    var seqNodes = qAll(timing, 'cTn').filter(function (n) {
      var nt = xmlAttr(n, 'nodeType');
      return nt === 'mainSeq' || nt === 'interactiveSeq';
    });
    var searchRoots = seqNodes.length ? seqNodes : [timing];

    var steps = [];
    var autoStep = { targets: [] };
    var seenInStep = {};

    function extractTargetsFromEffectNode(effectCtn, presetClass) {
      var targets = [];
      var bhvrs = qAll(effectCtn, 'tgtEl');
      bhvrs.forEach(function (te) {
        var spTgt = qOne(te, 'spTgt');
        if (!spTgt) return;
        var spid = xmlAttr(spTgt, 'spid');
        if (!spid) return;
        var pRg = qOne(spTgt, 'pRg');
        var pStart = pRg ? parseInt(xmlAttr(pRg, 'st'), 10) : -1;
        var pEnd = pRg ? parseInt(xmlAttr(pRg, 'end'), 10) : -1;
        targets.push({
          spid: spid,
          presetClass: presetClass,
          pStart: isNaN(pStart) ? -1 : pStart,
          pEnd: isNaN(pEnd) ? -1 : pEnd
        });
      });
      return targets;
    }

    searchRoots.forEach(function (rootNode) {
      var allCtns = qAll(rootNode, 'cTn');
      allCtns.forEach(function (ctn) {
        var nt = xmlAttr(ctn, 'nodeType');
        var pClass = xmlAttr(ctn, 'presetClass');
        var isEffectNode = (nt === 'clickEffect' || nt === 'withEffect' || nt === 'afterEffect') ||
                           (pClass === 'entr' || pClass === 'exit' || pClass === 'emph' || pClass === 'path');
        if (!isEffectNode) return;

        if (!pClass) {
          var animEf = qOne(ctn, 'animEffect');
          if (animEf) {
            pClass = xmlAttr(animEf, 'transition') === 'out' ? 'exit' : 'entr';
          } else if (qOne(ctn, 'set') || qOne(ctn, 'anim') || qOne(ctn, 'animScale') || qOne(ctn, 'animMotion')) {
            var toVal = qOne(ctn, 'to');
            var strVal = toVal ? qOne(toVal, 'strVal') : null;
            pClass = (strVal && (xmlAttr(strVal, 'val') === 'hidden' || xmlAttr(strVal, 'val') === 'false')) ? 'exit' : 'entr';
          } else {
            return;
          }
        }
        if (pClass === 'path') pClass = 'emph';
        if (pClass !== 'entr' && pClass !== 'exit' && pClass !== 'emph') return;

        var tgts = extractTargetsFromEffectNode(ctn, pClass);
        if (!tgts.length) return;

        // 첫 클릭 전에 '이전 효과와 함께/다음에'로 시작하는 효과는 슬라이드가 열릴 때 저절로 재생됩니다(클릭 단계가 아님).
        var curStep;
        if (nt === 'clickEffect') {
          steps.push({ targets: [] });
          seenInStep = {};
          curStep = steps[steps.length - 1];
        } else if (!steps.length) {
          curStep = autoStep;
        } else {
          curStep = steps[steps.length - 1];
        }
        tgts.forEach(function (t) {
          var key = t.spid + ':' + t.pStart + ':' + t.pEnd + ':' + t.presetClass;
          if (!seenInStep[key]) {
            seenInStep[key] = true;
            curStep.targets.push(t);
          }
        });
      });
    });

    return { steps: steps.filter(function (s) { return s.targets.length > 0; }), autoTargets: autoStep.targets };
  }

  function defaultPlaceholderBox(phType, phIdx, bw, bh) {
    if (phType === 'title' || phType === 'ctrTitle') {
      return { x: bw * 0.05, y: bh * 0.055, w: bw * 0.9, h: bh * 0.16 };
    }
    if (phType === 'subTitle') {
      return { x: bw * 0.08, y: bh * 0.24, w: bw * 0.84, h: bh * 0.14 };
    }
    if (phIdx === '1' || phType === 'body' || phType === 'obj') {
      return { x: bw * 0.06, y: bh * 0.23, w: bw * 0.88, h: bh * 0.68 };
    }
    return { x: bw * 0.06, y: bh * 0.18, w: bw * 0.88, h: bh * 0.7 };
  }

  async function loadPptxFile(file) {
    if (typeof JSZip === 'undefined') { showToast('JSZip 라이브러리를 불러오지 못했습니다.'); return; }
    try {
      closeWebBrowser();
      showToast('PPTX 슬라이드와 애니메이션 액션을 불러오는 중...');
      var zip = await JSZip.loadAsync(await file.arrayBuffer());
      var parser = new DOMParser();
      function parseXmlSafe(xmlStr) {
        return parser.parseFromString(String(xmlStr || '').replace(/^\uFEFF/, ''), 'application/xml');
      }

      var presFile = zip.file('ppt/presentation.xml') || zip.file('ppt/Presentation.xml');
      if (!presFile) throw new Error('유효한 PPTX 구조(ppt/presentation.xml)가 아닙니다.');
      var presDoc = parseXmlSafe(await presFile.async('string'));
      var sldSz = qOne(presDoc, 'sldSz');
      var slideCx = parseInt(xmlAttr(sldSz, 'cx'), 10) || 12192000;
      var slideCy = parseInt(xmlAttr(sldSz, 'cy'), 10) || 6858000;
      var bw = 1280;
      var bh = Math.round(bw * (slideCy / slideCx));
      var pxPerPt = bw / Math.max(200, slideCx / 12700);

      // 1) ppt/theme/theme1.xml 에서 실제 테마 색상표(dk1, lt1, accent1~6 등) 추출
      activeThemeMap = {
        bg1: '#ffffff', lt1: '#ffffff', tx1: '#0f172a', dk1: '#0f172a',
        bg2: '#f1f5f9', lt2: '#e2e8f0', tx2: '#1e293b', dk2: '#1e293b',
        accent1: '#2563eb', accent2: '#ea580c', accent3: '#16a34a',
        accent4: '#d97706', accent5: '#0891b2', accent6: '#7c3aed',
        hlink: '#2563eb', folHlink: '#7c3aed'
      };
      var themeFile = zip.file('ppt/theme/theme1.xml');
      if (themeFile) {
        var tDoc = parseXmlSafe(await themeFile.async('string'));
        var clrScheme = qOne(tDoc, 'clrScheme');
        if (clrScheme && clrScheme.children) {
          Array.from(clrScheme.children).forEach(function (cNode) {
            var slot = cNode.localName;
            var srgb = qOne(cNode, 'srgbClr');
            var sys = qOne(cNode, 'sysClr');
            if (srgb && xmlAttr(srgb, 'val')) activeThemeMap[slot] = '#' + xmlAttr(srgb, 'val');
            else if (sys && (xmlAttr(sys, 'lastClr') || xmlAttr(sys, 'val'))) {
              activeThemeMap[slot] = '#' + (xmlAttr(sys, 'lastClr') || (xmlAttr(sys, 'val') === 'windowText' ? '000000' : 'FFFFFF'));
            }
          });
          activeThemeMap.bg1 = activeThemeMap.lt1 || '#ffffff';
          activeThemeMap.tx1 = activeThemeMap.dk1 || '#0f172a';
          activeThemeMap.bg2 = activeThemeMap.lt2 || '#f1f5f9';
          activeThemeMap.tx2 = activeThemeMap.dk2 || '#1e293b';
        }
      }

      // 2) ppt/slideMasters/slideMaster1.xml 에서 clrMap 및 마스터 기본 배경·플레이스홀더 추출
      var masterPhBoxes = {};
      var masterBgColor = null;
      var masterFile = zip.file('ppt/slideMasters/slideMaster1.xml');
      if (masterFile) {
        var mDoc = parseXmlSafe(await masterFile.async('string'));
        var clrMap = qOne(mDoc, 'clrMap');
        if (clrMap) {
          ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].forEach(function (k) {
            var mapped = xmlAttr(clrMap, k);
            if (mapped && activeThemeMap[mapped]) activeThemeMap[k] = activeThemeMap[mapped];
          });
        }
        masterBgColor = extractFillColor(qOne(mDoc, 'bg'));
        qAll(mDoc, 'sp').forEach(function (sp) {
          var ph = qOne(sp, 'ph');
          if (!ph) return;
          var key = (xmlAttr(ph, 'type') || '') + ':' + (xmlAttr(ph, 'idx') || '');
          var xfrm = qOne(sp, 'xfrm');
          var off = xfrm ? qOne(xfrm, 'off') : null;
          var ext = xfrm ? qOne(xfrm, 'ext') : null;
          if (off && ext && xmlAttr(off, 'x') !== '') {
            var box = {
              x: (parseInt(xmlAttr(off, 'x'), 10) / slideCx) * bw,
              y: (parseInt(xmlAttr(off, 'y'), 10) / slideCy) * bh,
              w: (parseInt(xmlAttr(ext, 'cx'), 10) / slideCx) * bw,
              h: (parseInt(xmlAttr(ext, 'cy'), 10) / slideCy) * bh
            };
            masterPhBoxes[key] = box;
            if (xmlAttr(ph, 'type')) masterPhBoxes[xmlAttr(ph, 'type') + ':'] = box;
            if (xmlAttr(ph, 'idx')) masterPhBoxes[':' + xmlAttr(ph, 'idx')] = box;
          }
        });
      }

      var presRelsFile = zip.file('ppt/_rels/presentation.xml.rels');
      var relMap = {};
      if (presRelsFile) {
        var rDoc = parseXmlSafe(await presRelsFile.async('string'));
        qAll(rDoc, 'Relationship').forEach(function (rel) { relMap[xmlAttr(rel, 'Id')] = xmlAttr(rel, 'Target'); });
      }

      var slidePaths = [];
      qAll(presDoc, 'sldId').forEach(function (sid) {
        var rid = sid.getAttribute('r:id') || sid.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
        if (rid && relMap[rid]) {
          var t = relMap[rid].replace(/^\/+/, '').replace(/^\.\.\//, '');
          var full = t.indexOf('ppt/') === 0 ? t : ('ppt/' + t);
          if (zip.file(full)) slidePaths.push(full);
        }
      });
      if (!slidePaths.length) {
        slidePaths = Object.keys(zip.files).filter(function (k) { return /^ppt\/slides\/slide\d+\.xml$/i.test(k); })
          .sort(function (a, b) { return parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10); });
      }

      var mediaDataUrlCache = {};
      async function getMediaDataUrl(relTarget, baseDir) {
        if (!relTarget) return null;
        var norm = String(relTarget).replace(/\\/g, '/');
        if (norm.indexOf('../') === 0) norm = 'ppt/' + norm.replace(/^(\.\.\/)+/, '');
        else if (norm.indexOf('/') === 0) norm = norm.replace(/^\/+/, '');
        else if (norm.indexOf('ppt/') !== 0) norm = (baseDir || 'ppt/slides/') + norm;
        if (mediaDataUrlCache[norm]) return mediaDataUrlCache[norm];
        var mf = zip.file(norm);
        if (!mf) return null;
        var ext = norm.split('.').pop().toLowerCase();
        if (ext === 'emf' || ext === 'wmf') return null;
        var mime = ext === 'png' ? 'image/png' : (ext === 'gif' ? 'image/gif' : (ext === 'svg' ? 'image/svg+xml' : (ext === 'webp' ? 'image/webp' : 'image/jpeg')));
        var b64 = await mf.async('base64');
        var dUrl = 'data:' + mime + ';base64,' + b64;
        mediaDataUrlCache[norm] = dUrl;
        return dUrl;
      }

      // slideLayout 의 placeholder 좌표 추출 헬퍼
      async function parseLayoutPlaceholders(layoutPath) {
        var phBoxes = {};
        var lf = zip.file(layoutPath);
        if (!lf) return { phBoxes: phBoxes, bgColor: null };
        var lDoc = parseXmlSafe(await lf.async('string'));
        var lBg = extractFillColor(qOne(lDoc, 'bg'));
        qAll(lDoc, 'sp').forEach(function (sp) {
          var ph = qOne(sp, 'ph');
          if (!ph) return;
          var key = (xmlAttr(ph, 'type') || '') + ':' + (xmlAttr(ph, 'idx') || '');
          var xfrm = qOne(sp, 'xfrm');
          var off = xfrm ? qOne(xfrm, 'off') : null;
          var ext = xfrm ? qOne(xfrm, 'ext') : null;
          if (off && ext && xmlAttr(off, 'x') !== '') {
            var box = {
              x: (parseInt(xmlAttr(off, 'x'), 10) / slideCx) * bw,
              y: (parseInt(xmlAttr(off, 'y'), 10) / slideCy) * bh,
              w: (parseInt(xmlAttr(ext, 'cx'), 10) / slideCx) * bw,
              h: (parseInt(xmlAttr(ext, 'cy'), 10) / slideCy) * bh
            };
            phBoxes[key] = box;
            if (xmlAttr(ph, 'type')) phBoxes[xmlAttr(ph, 'type') + ':'] = box;
            if (xmlAttr(ph, 'idx')) phBoxes[':' + xmlAttr(ph, 'idx')] = box;
          }
        });
        return { phBoxes: phBoxes, bgColor: lBg };
      }

      var newPages = [];
      for (var si = 0; si < slidePaths.length; si++) {
        var spath = slidePaths[si];
        var sf = zip.file(spath);
        if (!sf) continue;
        var sDoc = parseXmlSafe(await sf.async('string'));

        var sRelsPath = spath.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels';
        var sRelMap = {};
        var layoutPath = '';
        if (zip.file(sRelsPath)) {
          var srDoc = parseXmlSafe(await zip.file(sRelsPath).async('string'));
          qAll(srDoc, 'Relationship').forEach(function (rel) {
            var tgt = xmlAttr(rel, 'Target');
            sRelMap[xmlAttr(rel, 'Id')] = tgt;
            if (xmlAttr(rel, 'Type').indexOf('/slideLayout') > 0) {
              var cleanLayout = tgt.replace(/^(\.\.\/)+/, '').replace(/^\/+/, '');
              layoutPath = cleanLayout.indexOf('ppt/') === 0 ? cleanLayout : ('ppt/' + cleanLayout);
            }
          });
        }

        var layoutInfo = layoutPath ? await parseLayoutPlaceholders(layoutPath) : { phBoxes: {}, bgColor: null };
        var bgNode = qOne(sDoc, 'bg');
        var bgColor = extractFillColor(bgNode) || layoutInfo.bgColor || masterBgColor || activeThemeMap.bg1 || '#ffffff';
        var isDarkSlideBg = false;
        if (bgColor && bgColor.charAt(0) === '#' && bgColor.length >= 7) {
          var br = parseInt(bgColor.slice(1, 3), 16), bgG = parseInt(bgColor.slice(3, 5), 16), bb = parseInt(bgColor.slice(5, 7), 16);
          if ((br * 0.299 + bgG * 0.587 + bb * 0.114) < 128) isDarkSlideBg = true;
        }
        var bgDataUrl = null;
        var bgBlip = bgNode ? qOne(bgNode, 'blip') : null;
        if (bgBlip) {
          var brId = bgBlip.getAttribute('r:embed') || bgBlip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');
          if (brId && sRelMap[brId]) bgDataUrl = await getMediaDataUrl(sRelMap[brId]);
        }

        var elements = [];

        // 그룹 도형(<p:grpSp>) 재귀 좌표 변환기
        function mapCoord(rawX, rawY, rawW, rawH, tf) {
          var gx = ((rawX - tf.chX) / Math.max(1, tf.chW)) * tf.w + tf.x;
          var gy = ((rawY - tf.chY) / Math.max(1, tf.chH)) * tf.h + tf.y;
          var gw = (rawW / Math.max(1, tf.chW)) * tf.w;
          var gh = (rawH / Math.max(1, tf.chH)) * tf.h;
          return { x: gx, y: gy, w: gw, h: gh };
        }

        async function parseShapeTree(containerNode, tf) {
          var kids = Array.from(containerNode.children || []);
          for (var ci = 0; ci < kids.length; ci++) {
            var ch = kids[ci];
            var ln = ch.localName;
            if (ln === 'AlternateContent') {
              var fb = Array.from(ch.children || []).find(function (c) { return c.localName === 'Fallback'; });
              var branch = fb || (ch.children && ch.children[0]);
              if (branch) await parseShapeTree(branch, tf);
              continue;
            }
            if (ln === 'grpSp') {
              var grpSpPr = directChild(ch, 'grpSpPr');
              var gxfrm = grpSpPr ? directChild(grpSpPr, 'xfrm') : null;
              var goff = gxfrm ? directChild(gxfrm, 'off') : null;
              var gext = gxfrm ? directChild(gxfrm, 'ext') : null;
              var gchOff = gxfrm ? directChild(gxfrm, 'chOff') : null;
              var gchExt = gxfrm ? directChild(gxfrm, 'chExt') : null;
              var nextTf = tf;
              if (goff && gext && gchOff && gchExt && parseInt(xmlAttr(gchExt, 'cx'), 10) > 0) {
                var mappedG = mapCoord(
                  parseInt(xmlAttr(goff, 'x'), 10) || 0,
                  parseInt(xmlAttr(goff, 'y'), 10) || 0,
                  parseInt(xmlAttr(gext, 'cx'), 10) || slideCx,
                  parseInt(xmlAttr(gext, 'cy'), 10) || slideCy,
                  tf
                );
                nextTf = {
                  x: mappedG.x, y: mappedG.y, w: mappedG.w, h: mappedG.h,
                  chX: parseInt(xmlAttr(gchOff, 'x'), 10) || 0,
                  chY: parseInt(xmlAttr(gchOff, 'y'), 10) || 0,
                  chW: parseInt(xmlAttr(gchExt, 'cx'), 10) || slideCx,
                  chH: parseInt(xmlAttr(gchExt, 'cy'), 10) || slideCy
                };
              }
              // 그룹에 건 애니메이션이 그룹 안 도형에도 적용되도록, 안쪽 도형에 그룹 id를 기록합니다.
              var grpNv = directChild(ch, 'nvGrpSpPr');
              var grpCNvPr = grpNv ? directChild(grpNv, 'cNvPr') : null;
              var grpId = grpCNvPr ? xmlAttr(grpCNvPr, 'id') : '';
              var grpStart = elements.length;
              await parseShapeTree(ch, nextTf);
              if (grpId) for (var gi = grpStart; gi < elements.length; gi++) elements[gi].groups = [grpId].concat(elements[gi].groups || []);
              continue;
            }

            if (ln !== 'sp' && ln !== 'pic' && ln !== 'cxnSp' && ln !== 'graphicFrame') continue;

            var cNvPr = qOne(ch, 'cNvPr');
            var spid = cNvPr ? xmlAttr(cNvPr, 'id') : ('e_' + elements.length);
            var ph = qOne(ch, 'ph');
            var phType = ph ? xmlAttr(ph, 'type') : '';
            var phIdx = ph ? xmlAttr(ph, 'idx') : '';

            var spPrForXfrm = directChild(ch, 'spPr');
            var xfrm = (spPrForXfrm && directChild(spPrForXfrm, 'xfrm')) || directChild(ch, 'xfrm') || qOne(ch, 'xfrm');
            var off = xfrm ? (directChild(xfrm, 'off') || qOne(xfrm, 'off')) : null;
            var ext = xfrm ? (directChild(xfrm, 'ext') || qOne(xfrm, 'ext')) : null;
            var box;
            if (off && ext && xmlAttr(off, 'x') !== '' && xmlAttr(ext, 'cx') !== '') {
              box = mapCoord(
                parseInt(xmlAttr(off, 'x'), 10) || 0,
                parseInt(xmlAttr(off, 'y'), 10) || 0,
                parseInt(xmlAttr(ext, 'cx'), 10) || 0,
                parseInt(xmlAttr(ext, 'cy'), 10) || 0,
                tf
              );
            } else {
              box = layoutInfo.phBoxes[phType + ':' + phIdx] ||
                    layoutInfo.phBoxes[phType + ':'] ||
                    layoutInfo.phBoxes[':' + phIdx] ||
                    masterPhBoxes[phType + ':' + phIdx] ||
                    masterPhBoxes[phType + ':'] ||
                    masterPhBoxes[':' + phIdx] ||
                    defaultPlaceholderBox(phType, phIdx, bw, bh);
            }

            if (ln === 'pic') {
              var blip = qOne(ch, 'blip');
              var rid = blip ? (blip.getAttribute('r:embed') || blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed')) : '';
              var imgDataUrl = (rid && sRelMap[rid]) ? await getMediaDataUrl(sRelMap[rid]) : null;
              if (imgDataUrl) elements.push({ kind: 'pic', spid: spid, x: box.x, y: box.y, w: box.w, h: box.h, imgDataUrl: imgDataUrl });
            } else if (ln === 'graphicFrame') {
              var tbl = qOne(ch, 'tbl');
              if (tbl) {
                var rows = qAll(tbl, 'tr').map(function (tr) {
                  return qAll(tr, 'tc').map(function (tc) {
                    return qAll(tc, 't').map(function (t) { return t.textContent; }).join('');
                  });
                });
                elements.push({ kind: 'table', spid: spid, x: box.x, y: box.y, w: box.w, h: box.h, rows: rows });
              }
            } else {
              var spPr = directChild(ch, 'spPr') || qOne(ch, 'spPr');
              var noFill = spPr ? directChild(spPr, 'noFill') : null;
              var fillNode = spPr ? (directChild(spPr, 'solidFill') || directChild(spPr, 'gradFill')) : null;
              var fill = noFill ? null : extractFillColor(fillNode);
              // 도형에 색을 직접 주지 않으면 파워포인트 기본 도형 스타일(<p:style>의 테마 색)을 따릅니다.
              var shStyle = directChild(ch, 'style');
              var styleRef = function (name) { var r = shStyle ? directChild(shStyle, name) : null; return r && xmlAttr(r, 'idx') !== '0' ? r : null; };
              if (!fill && !noFill && !fillNode && styleRef('fillRef')) fill = extractFillColor(styleRef('fillRef'));
              var styleFontColor = styleRef('fontRef') ? extractFillColor(styleRef('fontRef')) : null;
              // 도형 내부에 그림 채우기(<a:blipFill>)가 있는 경우 이미지 도형으로 추가
              var spBlipFill = spPr ? directChild(spPr, 'blipFill') : null;
              if (spBlipFill) {
                var spBlip = qOne(spBlipFill, 'blip');
                var spRid = spBlip ? (spBlip.getAttribute('r:embed') || spBlip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed')) : '';
                var spImgUrl = (spRid && sRelMap[spRid]) ? await getMediaDataUrl(sRelMap[spRid]) : null;
                if (spImgUrl) elements.push({ kind: 'pic', spid: spid, x: box.x, y: box.y, w: box.w, h: box.h, imgDataUrl: spImgUrl });
              }
              var lnNode = spPr ? directChild(spPr, 'ln') : null;
              var noLine = lnNode ? directChild(lnNode, 'noFill') : null;
              var stroke = (lnNode && !noLine) ? (extractFillColor(lnNode) || (ln === 'cxnSp' ? '#475569' : null)) : null;
              if (!stroke && !noLine && !(lnNode && extractFillColor(lnNode)) && styleRef('lnRef')) stroke = extractFillColor(styleRef('lnRef'));
              var prstGeom = spPr ? qOne(spPr, 'prstGeom') : null;
              var geom = prstGeom ? xmlAttr(prstGeom, 'prst') : (ln === 'cxnSp' ? 'line' : 'rect');

              var isDarkBox = isDarkSlideBg;
              if (fill && fill.charAt(0) === '#' && fill.length >= 7) {
                var fr = parseInt(fill.slice(1, 3), 16), fg = parseInt(fill.slice(3, 5), 16), fb2 = parseInt(fill.slice(5, 7), 16);
                isDarkBox = (fr * 0.299 + fg * 0.587 + fb2 * 0.114) < 135;
              }

              var paras = [];
              var txBody = directChild(ch, 'txBody') || qOne(ch, 'txBody');
              if (txBody) {
                var pNodes = Array.from(txBody.children || []).filter(function (n) { return n.localName === 'p'; });
                pNodes.forEach(function (pNode, pIdx) {
                  var pPr = directChild(pNode, 'pPr');
                  var align = pPr ? (xmlAttr(pPr, 'algn') || (phType === 'ctrTitle' ? 'ctr' : 'l')) : (phType === 'ctrTitle' ? 'ctr' : 'l');
                  var buChar = pPr ? qOne(pPr, 'buChar') : null;
                  var bullet = buChar ? (xmlAttr(buChar, 'char') + ' ') : '';
                  var lineTxt = '';
                  var fontSize = (phType === 'title' || phType === 'ctrTitle') ? 34 : 21;
                  var fontColor = styleFontColor || (isDarkBox ? '#f8fafc' : (activeThemeMap.tx1 || '#0f172a'));
                  var bold = (phType === 'title' || phType === 'ctrTitle');

                  qAll(pNode, 'r').forEach(function (rNode) {
                    var rPr = directChild(rNode, 'rPr');
                    if (rPr) {
                      var sz = parseInt(xmlAttr(rPr, 'sz'), 10);
                      if (sz > 0) fontSize = Math.max(11, Math.min(68, Math.round((sz / 100) * pxPerPt)));
                      if (xmlAttr(rPr, 'b') === '1') bold = true;
                      var rc = extractFillColor(rPr);
                      if (rc) fontColor = rc;
                    }
                    var tNode = directChild(rNode, 't');
                    if (tNode) lineTxt += tNode.textContent;
                  });
                  if (!lineTxt) {
                    qAll(pNode, 't').forEach(function (tn) { lineTxt += tn.textContent; });
                  }
                  if (lineTxt.trim()) {
                    paras.push({ pIdx: pIdx, text: bullet + lineTxt, fontSize: fontSize, color: fontColor, bold: bold, align: align });
                  }
                });
              }
              if (fill || stroke || paras.length) {
                elements.push({ kind: 'shape', spid: spid, geom: geom, x: box.x, y: box.y, w: box.w, h: box.h, fill: fill, stroke: stroke, paras: paras });
              }
            }
          }
        }

        var spTree = qOne(sDoc, 'spTree');
        var rootTf = { x: 0, y: 0, w: bw, h: bh, chX: 0, chY: 0, chW: slideCx, chH: slideCy };
        if (spTree) await parseShapeTree(spTree, rootTf);

        // 슬라이드 요소가 우측·하단 경계를 벗어나는 경우(4:3 헤더+16:9 본문 등) 자동 비율 축소 보정
        var maxR = bw, maxB = bh;
        elements.forEach(function (elm) {
          if (!elm || elm.w <= 0 || elm.h <= 0) return;
          var rEdge = elm.x + elm.w;
          var bEdge = elm.y + elm.h;
          if (elm.x >= -bw * 0.15 && rEdge > maxR && rEdge < bw * 2.5) maxR = rEdge;
          if (elm.y >= -bh * 0.15 && bEdge > maxB && bEdge < bh * 2.5) maxB = bEdge;
        });
        if (maxR > bw * 1.04 || maxB > bh * 1.04) {
          var ovScale = Math.min(bw / maxR, bh / maxB);
          elements.forEach(function (elm) {
            elm.x *= ovScale;
            elm.y *= ovScale;
            elm.w *= ovScale;
            elm.h *= ovScale;
            if (elm.paras && elm.paras.length) {
              elm.paras.forEach(function (p) {
                if (p.fontSize) p.fontSize = Math.max(10, Math.round(p.fontSize * ovScale));
              });
            }
          });
        }

        var timingInfo = parseSlideTiming(sDoc);
        var slideObj = { bgColor: bgColor, bgDataUrl: bgDataUrl, elements: elements, steps: timingInfo.steps, autoTargets: timingInfo.autoTargets };
        await ensurePptxSlideImages(slideObj);
        newPages.push({
          kind: 'pptx',
          label: file.name,
          width: bw,
          height: bh,
          pptxSlide: slideObj,
          pptxStep: 0,
          objects: [],
          redoStack: []
        });
      }

      if (!newPages.length) { showToast('PPTX에서 슬라이드를 찾지 못했습니다.'); return; }
      state.pages = newPages;
      state.pageIndex = 0;
      state.pptxShowAll = false;
      state.currentNoteId = null;
      forgetLastNote();
      state.currentNoteTitle = file.name.replace(/\.pptx$/i, '');
      updateNoteBadge();
      resetStampNumber(true);
      resetZoom();
      resizeBoardToPage();
      await renderBackground();
      renderInk();
      var totalActs = newPages.reduce(function (a, p) { return a + (p.pptxSlide.steps.length || 0); }, 0);
      showToast(file.name + ' (' + newPages.length + '슬라이드 · 클릭 액션 ' + totalActs + '단계)을 불러왔습니다.');
    } catch (err) {
      showToast('PPTX 열기 오류: ' + (err.message || err));
    }
  }

  async function ensurePptxSlideImages(slide) {
    if (!slide) return;
    if (slide.bgDataUrl && !slide.bgImg) slide.bgImg = await loadImgElement(slide.bgDataUrl);
    var elms = slide.elements || [];
    for (var i = 0; i < elms.length; i++) {
      if (elms[i].kind === 'pic' && elms[i].imgDataUrl && !elms[i].img) {
        elms[i].img = await loadImgElement(elms[i].imgDataUrl);
      }
    }
  }

  // 애니메이션 효과를 순서대로 다시 재생해 지금 단계에서 각 도형·문단이 보이는지 정합니다.
  // 첫 효과가 '나타내기'면 처음엔 숨기고, 나타내기/끝내기가 반복되면 마지막 상태를 따릅니다.
  function computePptxVisibility(slide, stepIdx) {
    var events = {}, emphSet = {};
    function add(key, step, cls) { (events[key] = events[key] || []).push({ step: step, cls: cls }); }
    function addTarget(tg, stepNum) {
      if (tg.presetClass === 'emph') { if (stepNum === stepIdx && stepNum > 0) emphSet[tg.spid] = true; return; }
      if (tg.pStart >= 0) {
        for (var p = tg.pStart; p <= Math.max(tg.pStart, tg.pEnd); p++) add(tg.spid + ':' + p, stepNum, tg.presetClass);
      } else add(tg.spid, stepNum, tg.presetClass);
    }
    (slide.autoTargets || []).forEach(function (tg) { addTarget(tg, 0); });
    (slide.steps || []).forEach(function (st, idx) { (st.targets || []).forEach(function (tg) { addTarget(tg, idx + 1); }); });
    function visible(key) {
      var ev = events[key];
      if (!ev || !ev.length) return true;
      var on = ev[0].cls !== 'entr';
      ev.forEach(function (e) { if (e.step <= stepIdx) on = e.cls === 'entr'; });
      return on;
    }
    return { visible: visible, emphSet: emphSet };
  }

  function wrapCanvasText(ctx, text, maxW) {
    var words = String(text || '');
    var lines = [], cur = '';
    for (var i = 0; i < words.length; i++) {
      var test = cur + words[i];
      if (ctx.measureText(test).width > maxW && cur) {
        lines.push(cur);
        cur = words[i];
      } else {
        cur = test;
      }
    }
    if (cur) lines.push(cur);
    return lines;
  }

  function renderPptxSlideToCtx(slide, stepIdx, ctx, w, h) {
    ctx.fillStyle = slide.bgColor || '#ffffff';
    ctx.fillRect(0, 0, w, h);
    if (slide.bgImg) ctx.drawImage(slide.bgImg, 0, 0, w, h);

    var vis = computePptxVisibility(slide, stepIdx);
    (slide.elements || []).forEach(function (elm) {
      if (!vis.visible(elm.spid)) return;
      if ((elm.groups || []).some(function (gid) { return !vis.visible(gid); })) return;   // 숨은 그룹 안의 도형

      ctx.save();
      if (elm.kind === 'pic' && elm.img) {
        ctx.drawImage(elm.img, elm.x, elm.y, elm.w, elm.h);
      } else if (elm.kind === 'table' && elm.rows && elm.rows.length) {
        var rCnt = elm.rows.length;
        var cCnt = Math.max(1, elm.rows[0].length);
        var cellW = elm.w / cCnt, cellH = elm.h / rCnt;
        for (var r = 0; r < rCnt; r++) {
          for (var c = 0; c < cCnt; c++) {
            var cx = elm.x + c * cellW, cy = elm.y + r * cellH;
            ctx.fillStyle = r === 0 ? '#e2e8f0' : '#ffffff';
            ctx.fillRect(cx, cy, cellW, cellH);
            ctx.strokeStyle = '#94a3b8';
            ctx.lineWidth = 1.2;
            ctx.strokeRect(cx, cy, cellW, cellH);
            ctx.fillStyle = '#0f172a';
            ctx.font = (r === 0 ? 'bold ' : '') + '15px Pretendard, "Noto Sans KR", sans-serif';
            ctx.textBaseline = 'middle';
            ctx.fillText(String(elm.rows[r][c] || '').slice(0, 32), cx + 6, cy + cellH / 2);
          }
        }
      } else if (elm.kind === 'shape') {
        if (elm.geom === 'line' || elm.geom === 'straightConnector1') {
          ctx.strokeStyle = elm.stroke || elm.fill || '#475569';
          ctx.lineWidth = 2.2;
          ctx.beginPath();
          ctx.moveTo(elm.x, elm.y);
          ctx.lineTo(elm.x + elm.w, elm.y + elm.h);
          ctx.stroke();
        } else {
          if (elm.fill) {
            ctx.fillStyle = elm.fill;
            if (elm.geom === 'ellipse') {
              ctx.beginPath();
              ctx.ellipse(elm.x + elm.w / 2, elm.y + elm.h / 2, Math.max(1, elm.w / 2), Math.max(1, elm.h / 2), 0, 0, Math.PI * 2);
              ctx.fill();
            } else {
              ctx.fillRect(elm.x, elm.y, elm.w, elm.h);
            }
          }
          if (elm.stroke) {
            ctx.strokeStyle = elm.stroke;
            ctx.lineWidth = 2;
            if (elm.geom === 'ellipse') {
              ctx.beginPath();
              ctx.ellipse(elm.x + elm.w / 2, elm.y + elm.h / 2, Math.max(1, elm.w / 2), Math.max(1, elm.h / 2), 0, 0, Math.PI * 2);
              ctx.stroke();
            } else {
              ctx.strokeRect(elm.x, elm.y, elm.w, elm.h);
            }
          }
        }
        if (elm.paras && elm.paras.length) {
          var py = elm.y + 10;
          var padX = 12;
          var maxW = Math.max(40, elm.w - padX * 2);
          elm.paras.forEach(function (p, pIdx) {
            var pk = elm.spid + ':' + (typeof p.pIdx === 'number' ? p.pIdx : pIdx);
            var pShown = vis.visible(pk);

            ctx.fillStyle = p.color || '#0f172a';
            ctx.font = (p.bold ? 'bold ' : '') + p.fontSize + 'px Pretendard, "Noto Sans KR", "Malgun Gothic", sans-serif';
            ctx.textBaseline = 'top';
            ctx.textAlign = p.align === 'ctr' ? 'center' : (p.align === 'r' ? 'right' : 'left');
            var tx = p.align === 'ctr' ? (elm.x + elm.w / 2) : (p.align === 'r' ? (elm.x + elm.w - padX) : (elm.x + padX));
            // 아직 나타나지 않은 문단도 자리는 차지해, 나중 문단이 위로 당겨지지 않게 합니다(파워포인트와 같은 배치).
            wrapCanvasText(ctx, p.text, maxW).forEach(function (ln) {
              if (pShown) ctx.fillText(ln, tx, py);
              py += p.fontSize * 1.32;
            });
          });
        }
      }
      if (vis.emphSet[elm.spid]) {
        ctx.strokeStyle = '#f59e0b';
        ctx.lineWidth = 4;
        ctx.strokeRect(elm.x - 4, elm.y - 4, elm.w + 8, elm.h + 8);
      }
      ctx.restore();
    });
  }

  function stepPptxAction(dir) {
    var pg = curPage();
    if (pg.kind !== 'pptx' || !pg.pptxSlide) return;
    state.pptxShowAll = false;
    var maxSteps = (pg.pptxSlide.steps || []).length;
    var next = (pg.pptxStep || 0) + dir;
    if (next < 0) {
      if (state.pageIndex > 0) goPage(state.pageIndex - 1);
      return;
    }
    if (next > maxSteps) {
      if (state.pageIndex < state.pages.length - 1) goPage(state.pageIndex + 1);
      else showToast('마지막 슬라이드의 마지막 액션입니다.');
      return;
    }
    pg.pptxStep = next;
    renderBackground();
  }

  // ===================================================================
  // 삼성노트 스타일 로컬 폴더 & 판서 노트 보관함 (IndexedDB + Fallback)
  // ===================================================================
  function loadFolders() {
    try {
      var list = JSON.parse(localStorage.getItem(FOLDERS_KEY) || 'null');
      if (Array.isArray(list) && list.length) return list;
    } catch (e) {}
    return ['1학년 수학', '2학년 확률과통계', '수행평가·기출'];
  }
  function saveFolders(list) {
    try { localStorage.setItem(FOLDERS_KEY, JSON.stringify(list)); } catch (e) {}
  }

  function openIdb() {
    return new Promise(function (resolve) {
      if (!window.indexedDB) { resolve(null); return; }
      try {
        var req = indexedDB.open(IDB_NAME, 1);
        req.onupgradeneeded = function (e) {
          var db = e.target.result;
          if (!db.objectStoreNames.contains(IDB_STORE)) {
            db.createObjectStore(IDB_STORE, { keyPath: 'id' });
          }
        };
        req.onsuccess = function (e) { resolve(e.target.result); };
        req.onerror = function () { resolve(null); };
      } catch (err) { resolve(null); }
    });
  }

  async function dbGetAllNotes() {
    var db = await openIdb();
    if (db) {
      return await new Promise(function (resolve) {
        var tx = db.transaction(IDB_STORE, 'readonly');
        var req = tx.objectStore(IDB_STORE).getAll();
        req.onsuccess = function () {
          var arr = Array.isArray(req.result) ? req.result : [];
          arr.sort(function (a, b) { return (b.updatedAt || 0) - (a.updatedAt || 0); });
          resolve(arr);
        };
        req.onerror = function () { resolve([]); };
      });
    }
    try { return JSON.parse(localStorage.getItem(FALLBACK_NOTES_KEY) || '[]') || []; } catch (e) { return []; }
  }

  async function dbPutNote(noteRecord) {
    var db = await openIdb();
    if (db) {
      return await new Promise(function (resolve) {
        try {
          var tx = db.transaction(IDB_STORE, 'readwrite');
          tx.objectStore(IDB_STORE).put(noteRecord);
          tx.oncomplete = function () { resolve({ ok: true, storage: 'indexedDB' }); };
          tx.onerror = function (ev) {
            resolve({ ok: false, error: (ev && ev.target && ev.target.error) || tx.error || new Error('IndexedDB 저장 오류') });
          };
          tx.onabort = function (ev) {
            resolve({ ok: false, error: (ev && ev.target && ev.target.error) || tx.error || new Error('IndexedDB 저장 중단(용량 부족 가능)') });
          };
        } catch (err) {
          resolve({ ok: false, error: err });
        }
      });
    }
    var list = await dbGetAllNotes();
    var idx = list.findIndex(function (n) { return n.id === noteRecord.id; });
    if (idx >= 0) list[idx] = noteRecord;
    else list.unshift(noteRecord);
    try {
      localStorage.setItem(FALLBACK_NOTES_KEY, JSON.stringify(list));
      return { ok: true, storage: 'localStorage' };
    } catch (e) {
      return { ok: false, error: e };
    }
  }

  async function dbDeleteNote(id) {
    var db = await openIdb();
    if (db) {
      await new Promise(function (resolve) {
        var tx = db.transaction(IDB_STORE, 'readwrite');
        tx.objectStore(IDB_STORE).delete(id);
        tx.oncomplete = function () { resolve(true); };
        tx.onerror = function () { resolve(false); };
      });
      return;
    }
    var list = (await dbGetAllNotes()).filter(function (n) { return n.id !== id; });
    try { localStorage.setItem(FALLBACK_NOTES_KEY, JSON.stringify(list)); } catch (e) {}
  }

  function updateNoteBadge() {
    if (el('sbCurrentNoteBadge')) el('sbCurrentNoteBadge').textContent = state.currentNoteTitle || '미저장 노트';
    if (el('sbOpenerNoteTitle')) el('sbOpenerNoteTitle').textContent = state.currentNoteTitle || '새 판서 노트';
    if (el('sbVaultNoteTitleInput')) el('sbVaultNoteTitleInput').value = state.currentNoteTitle || '새 판서 노트';
    if (el('sbVaultSaveCurBtn')) {
      el('sbVaultSaveCurBtn').textContent = state.currentNoteId ? '현재 노트 덮어쓰기 저장' : '이 노트를 보관함에 저장';
    }
  }

  async function serializeCurrentNote(asNew, opts) {
    opts = opts || {};
    if (!opts.silent) commitInlineText();
    var isOverwrite = Boolean(!asNew && state.currentNoteId);
    var noteId = isOverwrite
      ? state.currentNoteId
      : ('note_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7));
    var titleInput = el('sbVaultNoteTitleInput');
    var folderSel = el('sbVaultFolderSelect');
    var title = (titleInput && titleInput.value.trim()) || state.currentNoteTitle || '새 판서 노트';
    var folder = (folderSel && folderSel.value) || state.currentFolder || loadFolders()[0];

    // 1쪽 썸네일 생성 (320x180 JPEG)
    var thumbUrl = '';
    try {
      var thumbCv = document.createElement('canvas');
      thumbCv.width = 320; thumbCv.height = 180;
      thumbCv.getContext('2d').drawImage(createCompositeCanvas(state.pages[0]), 0, 0, 320, 180);
      thumbUrl = thumbCv.toDataURL('image/jpeg', 0.72);
    } catch (e) {}

    var serializedPages = [];
    for (var i = 0; i < state.pages.length; i++) {
      var pg = state.pages[i];
      if (pg.kind === 'pdf' && state.pdfDoc) {
        if (!state.pdfCache[pg.pdfPageNum]) {
          state.pdfCache[pg.pdfPageNum] = await renderPdfPageToOffscreen(pg.pdfPageNum);
        }
        var pdfImgUrl = state.pdfCache[pg.pdfPageNum].toDataURL('image/jpeg', 0.85);
        serializedPages.push({
          kind: 'image',
          label: pg.label || 'PDF',
          width: pg.width,
          height: pg.height,
          imageDataUrl: pdfImgUrl,
          stampSeries: pg.stampSeries || 0,
          objects: JSON.parse(JSON.stringify(pg.objects || []))
        });
      } else if (pg.kind === 'pptx' && pg.pptxSlide) {
        var cleanSlide = {
          bgColor: pg.pptxSlide.bgColor,
          bgDataUrl: pg.pptxSlide.bgDataUrl || null,
          steps: pg.pptxSlide.steps || [],
          autoTargets: pg.pptxSlide.autoTargets || [],
          elements: (pg.pptxSlide.elements || []).map(function (e) {
            var copy = Object.assign({}, e);
            delete copy.img;
            return copy;
          })
        };
        serializedPages.push({
          kind: 'pptx',
          label: pg.label || 'PPTX',
          width: pg.width,
          height: pg.height,
          pptxSlide: cleanSlide,
          pptxStep: pg.pptxStep || 0,
          stampSeries: pg.stampSeries || 0,
          objects: JSON.parse(JSON.stringify(pg.objects || []))
        });
      } else if (pg.kind === 'image') {
        var imgUrl = pg.imageDataUrl;
        if (!imgUrl && pg.imageObj) {
          var tmp = document.createElement('canvas');
          tmp.width = pg.width; tmp.height = pg.height;
          tmp.getContext('2d').drawImage(pg.imageObj, 0, 0, pg.width, pg.height);
          imgUrl = tmp.toDataURL('image/jpeg', 0.85);
        }
        serializedPages.push({
          kind: 'image',
          label: pg.label || '이미지',
          width: pg.width,
          height: pg.height,
          imageDataUrl: imgUrl,
          stampSeries: pg.stampSeries || 0,
          objects: JSON.parse(JSON.stringify(pg.objects || []))
        });
      } else {
        serializedPages.push({
          kind: 'blank',
          label: pg.label || '칠판',
          width: pg.width || 1280,
          height: pg.height || 720,
          bgMode: pg.bgMode || 'dark',
          stampSeries: pg.stampSeries || 0,
          objects: JSON.parse(JSON.stringify(pg.objects || []))
        });
      }
    }

    var record = {
      id: noteId,
      title: title,
      folder: folder,
      pageCount: serializedPages.length,
      thumbUrl: thumbUrl,
      updatedAt: Date.now(),
      updatedStr: new Date().toLocaleString('ko-KR'),
      pages: serializedPages
    };
    var saveRes = await dbPutNote(record);
    if (!saveRes || !saveRes.ok) {
      if (opts.silent) { showToast('자동 저장에 실패했습니다 (저장 공간 부족 가능). 노트함에서 오래된 노트를 정리해 주세요.'); return false; }
      var errMsg = (saveRes && saveRes.error && (saveRes.error.message || saveRes.error.name)) || '저장 공간 부족';
      showToast('노트 저장 실패 (' + errMsg + ') — 불필요한 대용량 노트를 정리하거나 백업(.json) 후 다시 시도하세요.');
      return;
    }
    state.currentNoteId = noteId;
    state.currentNoteTitle = title;
    state.currentFolder = folder;
    markNoteOpened(noteId);
    autoSaveSig = boardSignature();
    updateNoteBadge();
    if (opts.silent) {
      if (el('sbVaultModal') && el('sbVaultModal').classList.contains('show')) await renderVaultUi();
      return true;
    }
    await renderVaultUi();
    showToast('[' + folder + '] 폴더에 "' + title + '" 노트를 ' + (isOverwrite ? '덮어쓰기 저장했습니다.' : '새 노트로 저장했습니다.'));
  }

  async function loadNoteFromVault(noteId, opts) {
    var notes = await dbGetAllNotes();
    var target = notes.find(function (n) { return n.id === noteId; });
    if (!target || !Array.isArray(target.pages) || !target.pages.length) return false;
    closeWebBrowser();
    state.currentNoteId = target.id;
    state.currentNoteTitle = target.title;
    state.currentFolder = target.folder;
    state.pages = target.pages.map(function (p) {
      return Object.assign({}, p, { redoStack: [] });
    });
    state.pageIndex = Math.max(0, Math.min(state.pages.length - 1, (opts && opts.pageIndex) || 0));
    markNoteOpened(target.id);
    autoSaveSig = boardSignature();
    updateNoteBadge();
    resetZoom();
    resizeBoardToPage();
    await renderBackground();
    renderInk();
    el('sbVaultModal').classList.remove('show');
    if (opts && opts.resumed) { showToast('지난번 판서로 돌아왔습니다: "' + target.title + '"' + (state.pageIndex ? ' · ' + (state.pageIndex + 1) + '쪽' : '')); return true; }
    showToast('"' + target.title + '" (' + state.pages.length + '쪽) 노트를 불러왔습니다.');
  }

  // 최근에 연 노트 기록 — 노트함을 최근 연 순서로 정렬하고, 다시 들어오면 마지막 판서로 돌아갑니다.
  var NOTE_OPENED_KEY = 'sbVaultNoteOpened_v1';
  var LAST_NOTE_KEY = 'sbLastNote_v1';
  var NOTE_SORT_KEY = 'sbVaultNoteSort_v1';
  function loadOpenedMap() { try { var m = JSON.parse(localStorage.getItem(NOTE_OPENED_KEY) || '{}'); return m && typeof m === 'object' ? m : {}; } catch (e) { return {}; } }
  function markNoteOpened(id) {
    if (!id) return;
    var m = loadOpenedMap();
    m[id] = Date.now();
    try {
      localStorage.setItem(NOTE_OPENED_KEY, JSON.stringify(m));
      localStorage.setItem(LAST_NOTE_KEY, JSON.stringify({ id: id, page: state.pageIndex || 0 }));
    } catch (e) {}
  }
  function rememberLastPage() {
    if (!state.currentNoteId) return;
    try { localStorage.setItem(LAST_NOTE_KEY, JSON.stringify({ id: state.currentNoteId, page: state.pageIndex || 0 })); } catch (e) {}
  }
  function forgetLastNote() { try { localStorage.removeItem(LAST_NOTE_KEY); } catch (e) {} }
  function sortNotesByRecent(notes) {
    var m = loadOpenedMap();
    function t(n) { return Math.max(m[n.id] || 0, n.updatedAt || 0); }
    return notes.slice().sort(function (a, b) { return t(b) - t(a); });
  }
  function noteSortMode() { try { return localStorage.getItem(NOTE_SORT_KEY) === 'manual' ? 'manual' : 'recent'; } catch (e) { return 'recent'; } }

  // 1분마다 자동 저장 — 바뀐 내용이 있을 때만 조용히 저장합니다. 아직 저장하지 않은 새 노트는 현재 폴더에 새 노트로 만들어 둡니다.
  var AUTO_SAVE_MS = 60000;
  var autoSaveSig = '';
  var autoSaving = false;
  function boardSignature() {
    try {
      return state.pages.length + '|' + state.pages.map(function (p) {
        return (p.kind || '') + ':' + (p.label || '') + ':' + JSON.stringify(p.objects || []).length + ':' + hashStr(JSON.stringify(p.objects || []));
      }).join('|');
    } catch (e) { return String(Date.now()); }
  }
  function hashStr(str) {
    var h = 0;
    for (var i = 0; i < str.length; i++) { h = ((h << 5) - h + str.charCodeAt(i)) | 0; }
    return h;
  }
  function boardHasContent() {
    return state.pages.some(function (p) { return (p.objects && p.objects.length) || (p.kind && p.kind !== 'blank'); });
  }
  async function autoSaveNow(reason) {
    if (autoSaving) return false;
    rememberLastPage();
    var sig = boardSignature();
    if (sig === autoSaveSig) return true;
    if (!boardHasContent()) return true;
    if (pointer && pointer.down) return false;
    autoSaving = true;
    try {
      if (!state.currentNoteId && (!state.currentNoteTitle || state.currentNoteTitle === '새 판서 노트')) {
        var d = new Date();
        state.currentNoteTitle = '판서 ' + (d.getMonth() + 1) + '/' + d.getDate() + ' ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
        updateNoteBadge();
      }
      var ok = await serializeCurrentNote(false, { silent: true, reason: reason });
      return ok !== false;
    } catch (e) { return false; }
    finally { autoSaving = false; }
  }
  async function resumeLastNote() {
    var last = null;
    try { last = JSON.parse(localStorage.getItem(LAST_NOTE_KEY) || 'null'); } catch (e) {}
    if (!last || !last.id) return false;
    return await loadNoteFromVault(last.id, { resumed: true, pageIndex: last.page || 0 });
  }
  async function goHomeSaved() {
    showToast('판서를 저장하고 홈으로 이동합니다…');
    await autoSaveNow('home');
    location.href = './index.html';
  }

  // 노트함 순서 — 내가 정한 순서를 기기에 저장합니다. 순서를 정하지 않은 새 노트는 맨 앞(최근순)에 둡니다.
  var NOTE_ORDER_KEY = 'sbVaultNoteOrder_v1';
  function loadNoteOrder() { try { var a = JSON.parse(localStorage.getItem(NOTE_ORDER_KEY) || '[]'); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
  function applyNoteOrder(notes) {
    var order = loadNoteOrder(), pos = {};
    order.forEach(function (id, i) { pos[id] = i; });
    var fresh = notes.filter(function (n) { return pos[n.id] == null; });
    var placed = notes.filter(function (n) { return pos[n.id] != null; }).sort(function (a, b) { return pos[a.id] - pos[b.id]; });
    return fresh.concat(placed);
  }
  // 보이는 목록(폴더) 안에서 바로 옆 노트와 자리를 바꾸고, 전체 순서를 저장합니다.
  function moveNoteOrder(allIds, visibleIds, id, dir) {
    var vi = visibleIds.indexOf(id), other = visibleIds[vi + dir];
    if (vi < 0 || !other) { showToast(dir < 0 ? '이미 맨 앞입니다.' : '이미 맨 뒤입니다.'); return; }
    var a = allIds.indexOf(id), b = allIds.indexOf(other);
    allIds[a] = other; allIds[b] = id;
    try { localStorage.setItem(NOTE_ORDER_KEY, JSON.stringify(allIds)); } catch (e) {}
  }

  // 노트 복사 — 같은 수업을 다른 학급에서 쓸 수 있게, 자료만(판서 비움) 또는 판서까지 복사합니다.
  var copyTarget = null;
  function openNoteCopyBox(note) {
    if (!note) return;
    copyTarget = note;
    var box = el('sbVaultCopyBox');
    var folders = loadFolders();
    var sel = el('sbVaultCopyFolder');
    sel.innerHTML = folders.map(function (f) { return '<option value="' + f.replace(/"/g, '&quot;') + '">' + f + '</option>'; }).join('');
    sel.value = state.activeVaultFolder !== 'ALL' && folders.indexOf(state.activeVaultFolder) >= 0 ? state.activeVaultFolder : (note.folder || folders[0]);
    el('sbVaultCopyName').textContent = note.title || '판서 노트';
    el('sbVaultCopyTitle').value = (note.title || '판서 노트') + ' (복사)';
    box.hidden = false;
    box.scrollIntoView({ block: 'nearest' });
  }
  async function copyNote() {
    if (!copyTarget) return;
    var material = (document.querySelector('input[name="sbVaultCopyMode"]:checked') || {}).value !== 'all';
    var folder = el('sbVaultCopyFolder').value || copyTarget.folder;
    var copy = JSON.parse(JSON.stringify(copyTarget));
    copy.id = 'note_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7);
    copy.title = String(el('sbVaultCopyTitle').value || '').trim() || (copyTarget.title || '판서 노트') + ' (복사)';
    copy.folder = folder;
    copy.updatedAt = Date.now();
    copy.updatedStr = new Date().toLocaleString('ko-KR');
    if (material) {
      (copy.pages || []).forEach(function (pg) { pg.objects = []; });
      copy.thumbUrl = '';   // 미리보기에는 원래 학급의 판서가 담겨 있어 지웁니다. 열어서 저장하면 새로 만들어집니다.
    }
    var res = await dbPutNote(copy);
    if (!res || !res.ok) { showToast('노트를 복사하지 못했습니다 (저장 공간 부족일 수 있습니다).'); return; }
    el('sbVaultCopyBox').hidden = true;
    copyTarget = null;
    state.activeVaultFolder = folder;
    await renderVaultUi();
    showToast('[' + folder + '] 폴더에 "' + copy.title + '"을(를) ' + (material ? '자료만' : '판서까지') + ' 복사했습니다.');
  }

  // 판서 노트 백업(.json) 및 복원 — 전체 노트함 백업 + 개별 노트 백업 + 폴더 비파괴 병합 복원
  function safeFilePart(s) {
    return String(s || '판서노트').replace(/[\\/:*?"<>|]+/g, '_').replace(/\s+/g, '_').slice(0, 48) || '판서노트';
  }
  async function exportAllVaultJson() {
    await autoSaveNow('backup');
    var notes = await dbGetAllNotes();
    if (!notes.length && boardHasContent()) {
      await serializeCurrentNote(false, { silent: true, reason: 'backup' });
      notes = await dbGetAllNotes();
    }
    var folders = loadFolders();
    if (!notes.length) {
      showToast('저장된 판서 노트가 없습니다. 먼저 필기하거나 노트를 저장한 뒤 백업해 주세요.');
      return false;
    }
    var payload = {
      kind: 'smartboard-vault-backup',
      version: 2,
      exportedAt: new Date().toISOString(),
      origin: location.origin,
      folders: folders,
      notes: notes,
      noteOrder: loadNoteOrder(),
      openedMap: loadOpenedMap(),
      sortMode: noteSortMode(),
      lastNoteId: state.currentNoteId || null
    };
    var dateStr = new Date().toISOString().slice(0, 10);
    var blob = new Blob([JSON.stringify(payload)], { type: 'application/json' });
    downloadBlob(blob, '스마트판서_노트보관함_전체백업_' + dateStr + '.json');
    showToast('전체 백업 완료: 폴더 ' + folders.length + '개 · 판서 노트 ' + notes.length + '개를 다운로드 폴더에 저장했습니다.');
    return true;
  }
  async function exportSingleNoteJson(noteId) {
    var target = null;
    if (noteId) {
      var all = await dbGetAllNotes();
      target = all.find(function (n) { return n.id === noteId; }) || null;
    } else {
      await autoSaveNow('backup-single');
      if (!state.currentNoteId) {
        await serializeCurrentNote(false, { silent: true, reason: 'backup-single' });
      }
      if (state.currentNoteId) {
        var allCur = await dbGetAllNotes();
        target = allCur.find(function (n) { return n.id === state.currentNoteId; }) || null;
      }
    }
    if (!target) {
      showToast('백업할 판서 노트를 찾지 못했습니다.');
      return false;
    }
    var payload = {
      kind: 'smartboard-single-note',
      version: 2,
      exportedAt: new Date().toISOString(),
      origin: location.origin,
      folders: [target.folder || loadFolders()[0] || '기본'],
      notes: [target]
    };
    var dateStr = new Date().toISOString().slice(0, 10);
    var blob = new Blob([JSON.stringify(payload)], { type: 'application/json' });
    downloadBlob(blob, '스마트판서_노트_' + safeFilePart(target.title) + '_' + dateStr + '.json');
    showToast('"' + (target.title || '판서 노트') + '" 노트 백업(.json)을 다운로드했습니다.');
    return true;
  }
  async function importVaultJsonFile(f) {
    if (!f) return false;
    return await new Promise(function (resolve) {
      var fr = new FileReader();
      fr.onload = async function () {
        try {
          var data = JSON.parse(fr.result);
          var rawNotes = [];
          if (data && Array.isArray(data.notes)) {
            rawNotes = data.notes;
          } else if (data && data.id && Array.isArray(data.pages)) {
            rawNotes = [data];
          } else if (Array.isArray(data)) {
            rawNotes = data;
          } else {
            showToast('스마트 판서 백업 파일(.json) 형식이 아닙니다.');
            resolve(false);
            return;
          }
          var validNotes = rawNotes.filter(function (n) { return n && n.id && Array.isArray(n.pages) && n.pages.length; });
          if (!validNotes.length) {
            showToast('백업 파일 안에 복원할 판서 노트가 없습니다.');
            resolve(false);
            return;
          }

          // 기존 폴더를 지우지 않고 백업 폴더 + 노트 소속 폴더를 중복 없이 병합
          var mergedFolders = loadFolders().slice();
          function ensureFolder(fname) {
            var clean = String(fname || '').trim();
            if (clean && mergedFolders.indexOf(clean) < 0) mergedFolders.push(clean);
          }
          if (Array.isArray(data.folders)) data.folders.forEach(ensureFolder);
          validNotes.forEach(function (n) {
            if (!n.folder) n.folder = mergedFolders[0] || '1학년 수학';
            ensureFolder(n.folder);
          });
          if (mergedFolders.length) saveFolders(mergedFolders);

          var savedCount = 0;
          for (var i = 0; i < validNotes.length; i++) {
            var res = await dbPutNote(validNotes[i]);
            if (res && res.ok) savedCount++;
          }

          if (data.openedMap && typeof data.openedMap === 'object') {
            var curOpened = loadOpenedMap();
            Object.keys(data.openedMap).forEach(function (k) {
              var v = Number(data.openedMap[k]) || 0;
              if (v > (curOpened[k] || 0)) curOpened[k] = v;
            });
            try { localStorage.setItem(NOTE_OPENED_KEY, JSON.stringify(curOpened)); } catch (e) {}
          }
          if (Array.isArray(data.noteOrder) && data.noteOrder.length) {
            var curOrder = loadNoteOrder();
            data.noteOrder.forEach(function (id) {
              if (id && curOrder.indexOf(id) < 0) curOrder.push(id);
            });
            try { localStorage.setItem(NOTE_ORDER_KEY, JSON.stringify(curOrder)); } catch (e) {}
          }

          state.activeVaultFolder = 'ALL';
          await renderVaultUi();

          var wasBlank = !boardHasContent();
          if (validNotes.length === 1 || wasBlank) {
            var sorted = validNotes.slice().sort(function (a, b) { return (b.updatedAt || 0) - (a.updatedAt || 0); });
            var pickId = (data.lastNoteId && validNotes.some(function (n) { return n.id === data.lastNoteId; }))
              ? data.lastNoteId
              : sorted[0].id;
            await loadNoteFromVault(pickId);
            showToast('백업 복원 완료: 판서 노트 ' + savedCount + '개를 불러와 화면에 열었습니다.');
          } else {
            if (el('sbVaultModal')) el('sbVaultModal').classList.add('show');
            showToast('백업 복원 완료: 판서 노트 ' + savedCount + '개를 노트함에 복원했습니다.');
          }
          resolve(true);
        } catch (err) {
          showToast('백업 파일 형식이 올바르지 않습니다.');
          resolve(false);
        }
      };
      fr.onerror = function () {
        showToast('백업 파일을 읽는 중 오류가 발생했습니다.');
        resolve(false);
      };
      fr.readAsText(f);
    });
  }

  async function renderVaultUi() {
    var folders = loadFolders();
    var sel = el('sbVaultFolderSelect');
    if (sel) {
      sel.innerHTML = folders.map(function (f) {
        return '<option value="' + f.replace(/"/g, '&quot;') + '">' + f + '</option>';
      }).join('');
      if (folders.indexOf(state.currentFolder) >= 0) sel.value = state.currentFolder;
    }

    var notes = await dbGetAllNotes();
    var fList = el('sbVaultFolderList');
    if (fList) {
      var html = '<div class="sb-folder-item ' + (state.activeVaultFolder === 'ALL' ? 'active' : '') + '" data-vfolder="ALL">' +
        '<span>전체 노트 (' + notes.length + ')</span></div>';
      folders.forEach(function (f) {
        var cnt = notes.filter(function (n) { return n.folder === f; }).length;
        html += '<div class="sb-folder-item ' + (state.activeVaultFolder === f ? 'active' : '') + '" data-vfolder="' + f.replace(/"/g, '&quot;') + '">' +
          '<span>' + f + ' (' + cnt + ')</span>' +
          (folders.length > 1 ? '<button type="button" class="sb-folder-del" data-del-folder="' + f.replace(/"/g, '&quot;') + '" title="폴더 삭제">&times;</button>' : '') +
        '</div>';
      });
      fList.innerHTML = html;

      fList.querySelectorAll('.sb-folder-item').forEach(function (item) {
        item.addEventListener('click', function (e) {
          if (e.target.closest('.sb-folder-del')) return;
          state.activeVaultFolder = item.getAttribute('data-vfolder') || 'ALL';
          if (state.activeVaultFolder !== 'ALL') state.currentFolder = state.activeVaultFolder;
          renderVaultUi();
        });
      });
      fList.querySelectorAll('.sb-folder-del').forEach(function (db) {
        db.addEventListener('click', function (e) {
          e.stopPropagation();
          var targetF = db.getAttribute('data-del-folder');
          var nextF = loadFolders().filter(function (x) { return x !== targetF; });
          if (!nextF.length) return;
          saveFolders(nextF);
          if (state.activeVaultFolder === targetF) state.activeVaultFolder = 'ALL';
          renderVaultUi();
        });
      });
    }

    var sortMode = noteSortMode();
    notes = sortMode === 'manual' ? applyNoteOrder(notes) : sortNotesByRecent(notes);
    if (el('sbVaultSortSelect')) el('sbVaultSortSelect').value = sortMode;
    var filtered = state.activeVaultFolder === 'ALL'
      ? notes
      : notes.filter(function (n) { return n.folder === state.activeVaultFolder; });
    if (el('sbVaultActiveFolderTitle')) el('sbVaultActiveFolderTitle').textContent = state.activeVaultFolder === 'ALL' ? '전체 노트' : state.activeVaultFolder;
    if (el('sbVaultNoteCount')) el('sbVaultNoteCount').textContent = filtered.length + '개 노트';
    var openedMap = loadOpenedMap();

    var grid = el('sbVaultNotesGrid');
    if (!grid) return;
    if (!filtered.length) {
      grid.innerHTML = '<div style="grid-column:1/-1;padding:36px 14px;text-align:center;color:#64748b;font-size:0.86rem;">이 폴더에 저장된 판서 노트가 없습니다. 상단의 <b>[현재 노트 덮어쓰기 저장]</b> 또는 <b>[+ 새 노트로 저장]</b>을 눌러 저장해 보세요.</div>';
      return;
    }
    grid.innerHTML = filtered.map(function (n) {
      return '<div class="sb-note-card">' +
        '<div class="sb-note-thumb">' + (n.thumbUrl ? '<img src="' + n.thumbUrl + '" alt="">' : '<span class="sb-note-thumb-empty">판서 없음 · 자료만</span>') + '</div>' +
        '<div class="sb-note-meta">' +
          '<strong>' + (n.title || '판서 노트') + '</strong>' +
          '<small>' + (n.folder || '기본') + ' · ' + (n.pageCount || 1) + '쪽</small>' +
          '<small>' + (n.id === state.currentNoteId ? '<b style="color:#2563eb;">지금 여는 중</b> · ' : '') + (openedMap[n.id] ? '최근 열람 ' + new Date(openedMap[n.id]).toLocaleString('ko-KR', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }) : (n.updatedStr || '')) + '</small>' +
          '<div class="sb-note-actions">' +
            '<button type="button" class="sb-btn sb-btn-primary sb-btn-xs" style="flex:1;" data-load-note="' + n.id + '">열기</button>' +
            '<button type="button" class="sb-btn sb-btn-ghost sb-btn-xs" data-copy-note="' + n.id + '" title="다른 학급(폴더)으로 복사">복사</button>' +
            '<button type="button" class="sb-btn sb-btn-ghost sb-btn-xs" data-export-note="' + n.id + '" title="이 노트 1개만 백업 파일(.json)로 내려받기">백업</button>' +
            '<button type="button" class="sb-btn sb-btn-danger sb-btn-xs" data-del-note="' + n.id + '">삭제</button>' +
          '</div>' +
          '<div class="sb-note-order"' + (sortMode === 'manual' ? '' : ' style="display:none;"') + '>' +
            '<button type="button" class="sb-mini-btn" data-move-note="' + n.id + '" data-dir="-1" title="앞으로">◀ 앞으로</button>' +
            '<button type="button" class="sb-mini-btn" data-move-note="' + n.id + '" data-dir="1" title="뒤로">뒤로 ▶</button>' +
          '</div>' +
        '</div>' +
      '</div>';
    }).join('');

    grid.querySelectorAll('[data-load-note]').forEach(function (lb) {
      lb.addEventListener('click', function () { loadNoteFromVault(lb.getAttribute('data-load-note')); });
    });
    var visibleIds = filtered.map(function (n) { return n.id; });
    grid.querySelectorAll('[data-move-note]').forEach(function (mb) {
      mb.addEventListener('click', function () {
        moveNoteOrder(notes.map(function (n) { return n.id; }), visibleIds, mb.getAttribute('data-move-note'), parseInt(mb.getAttribute('data-dir'), 10));
        renderVaultUi();
      });
    });
    grid.querySelectorAll('[data-copy-note]').forEach(function (cb) {
      cb.addEventListener('click', function () { openNoteCopyBox(notes.find(function (n) { return n.id === cb.getAttribute('data-copy-note'); })); });
    });
    grid.querySelectorAll('[data-export-note]').forEach(function (eb) {
      eb.addEventListener('click', function () { exportSingleNoteJson(eb.getAttribute('data-export-note')); });
    });
    grid.querySelectorAll('[data-del-note]').forEach(function (db) {
      db.addEventListener('click', async function () {
        await dbDeleteNote(db.getAttribute('data-del-note'));
        renderVaultUi();
        showToast('노트를 삭제했습니다.');
      });
    });
  }

  // ===================================================================
  // 이미지 불러오기 & 페이지 이동
  // ===================================================================
  async function loadImageFromDataUrl(dataUrl, label) {
    closeWebBrowser();
    var img = await loadImgElement(dataUrl);
    if (!img) return;
    var bw = 1280;
    var bh = Math.round(bw * (img.naturalHeight / Math.max(1, img.naturalWidth)));
    bh = Math.max(480, Math.min(1800, bh));
    state.pages.push({
      kind: 'image',
      label: label || '이미지',
      width: bw,
      height: bh,
      imageDataUrl: dataUrl,
      imageObj: img,
      objects: [],
      redoStack: []
    });
    state.pageIndex = state.pages.length - 1;
    resetZoom();
    resizeBoardToPage();
    await renderBackground();
    renderInk();
    showToast((label || '이미지') + '를 새 페이지로 불러왔습니다.');
  }

  async function goPage(idx) {
    if (idx < 0 || idx >= state.pages.length) return;
    var pageChanged = (idx !== state.pageIndex);
    commitInlineText();
    clearSelection();
    state.pageIndex = idx;
    if (pageChanged) resetStampNumber(true);
    rememberLastPage();
    resetZoom();
    resizeBoardToPage();
    await renderBackground();
    renderInk();
  }

  function updateFooterUi() {
    var pg = curPage();
    var pageStr = (state.pageIndex + 1) + ' / ' + state.pages.length;
    if (el('sbPageIndicator')) {
      el('sbPageIndicator').textContent = pageStr + ' (' + (pg.label || '보드').slice(0, 12) + ')';
    }
    var pptxBar = el('sbPptxActionBar');
    var miniLbl = el('sbMiniNavLabel');
    if (pg.kind === 'pptx' && pg.pptxSlide) {
      if (pptxBar) pptxBar.style.display = 'flex';
      var tot = (pg.pptxSlide.steps || []).length;
      var actStr = state.pptxShowAll ? ('전체/' + tot) : ((pg.pptxStep || 0) + '/' + tot);
      if (el('sbPptxActIndicator')) el('sbPptxActIndicator').textContent = actStr;
      if (miniLbl) miniLbl.textContent = pageStr + '쪽 · 액션 ' + actStr;
    } else {
      if (pptxBar) pptxBar.style.display = 'none';
      if (miniLbl) miniLbl.textContent = pageStr + '쪽';
    }
  }

  // ===================================================================
  // 무료 OCR 손글씨 → 글씨 변환 (키 없이도 OCR.space 자동 폴백 지원!)
  // ===================================================================
  function loadOcrCfg() {
    try { return JSON.parse(localStorage.getItem(OCR_CFG_KEY) || '{}') || {}; } catch (e) { return {}; }
  }
  function saveOcrCfg(cfg) {
    try { localStorage.setItem(OCR_CFG_KEY, JSON.stringify(cfg)); } catch (e) {}
  }

  function buildOcrCropBase64(sel) {
    var pg = curPage();
    var pad = 16;
    var sx = Math.max(0, Math.round(sel.x - pad));
    var sy = Math.max(0, Math.round(sel.y - pad));
    var sw = Math.min(pg.width - sx, Math.round(sel.w + pad * 2));
    var sh = Math.min(pg.height - sy, Math.round(sel.h + pad * 2));
    var cv = document.createElement('canvas');
    cv.width = Math.max(60, sw);
    cv.height = Math.max(60, sh);
    var c = cv.getContext('2d');
    c.fillStyle = '#ffffff';
    c.fillRect(0, 0, cv.width, cv.height);
    if (sel.ids && sel.ids.length) {
      c.save();
      c.translate(-sx, -sy);
      pg.objects.forEach(function (o) {
        if (sel.ids.indexOf(o.id) >= 0) {
          var clone = Object.assign({}, o, { color: '#0f172a' });
          drawInkObject(c, clone);
        }
      });
      c.restore();
    } else {
      c.drawImage(createCompositeCanvas(), sx, sy, sw, sh, 0, 0, sw, sh);
    }
    return cv.toDataURL('image/png');
  }

  async function runOcrOnSelection() {
    if (!state.selection) return;
    var cfg = loadOcrCfg();
    var engine = cfg.engine || 'gemini';
    // Gemini 키가 비어 있으면 자동으로 무료 OCR.space 공용 엔진으로 폴백!
    if (engine === 'gemini' && !cfg.geminiKey) {
      engine = 'ocrspace';
    }

    state.ocrPendingBox = Object.assign({}, state.selection);
    var dataUrl = buildOcrCropBase64(state.selection);
    var b64 = dataUrl.split(',')[1];

    el('sbOcrResultText').value = '';
    el('sbOcrLoadingMsg').style.display = 'block';
    el('sbOcrResultModal').classList.add('show');

    try {
      var recognized = '';
      if (engine === 'gemini') {
        var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + encodeURIComponent(cfg.geminiKey);
        var resp = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [{
              parts: [
                { text: '이 이미지에 손으로 쓴 판서(한글, 영어, 숫자, 수학 수식)를 정확히 읽어 깔끔한 텍스트만 출력해 주세요. 불필요한 설명은 쓰지 마세요.' },
                { inline_data: { mime_type: 'image/png', data: b64 } }
              ]
            }]
          })
        });
        var j = await resp.json();
        if (j.error) throw new Error(j.error.message || 'Gemini API 오류');
        recognized = (j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts || [])
          .map(function (p) { return p.text || ''; }).join('').trim();
      } else {
        var fd = new FormData();
        fd.append('base64Image', dataUrl);
        fd.append('language', 'kor');
        fd.append('OCREngine', '2');
        fd.append('scale', 'true');
        var resp2 = await fetch('https://api.ocr.space/parse/image', {
          method: 'POST',
          headers: { apikey: (cfg.ocrSpaceKey || '').trim() || 'helloworld' },
          body: fd
        });
        var j2 = await resp2.json();
        if (j2.IsErroredOnProcessing) throw new Error((j2.ErrorMessage && j2.ErrorMessage[0]) || 'OCR.space 오류');
        recognized = (j2.ParsedResults || []).map(function (r) { return r.ParsedText || ''; }).join('\n').trim();
      }
      el('sbOcrResultText').value = recognized || '(인식된 글자가 없습니다. 직접 입력해 수정하거나, OCR 설정에서 Gemini 무료 키를 등록하면 한글 손글씨 인식률이 크게 높아집니다.)';
    } catch (err) {
      el('sbOcrResultText').value = '변환 안내: ' + (err.message || err) + '\n(이 칸에 직접 변환할 텍스트를 입력해 [선택 판서를 글씨로 교체]를 누를 수도 있습니다.)';
    } finally {
      el('sbOcrLoadingMsg').style.display = 'none';
    }
  }

  // ===================================================================
  // PNG 저장 & 전체 페이지 합본 PDF 저장
  // ===================================================================
  function exportBaseName() {
    var t = (state.currentNoteTitle || '스마트판서').replace(/[\\/:*?"<>|]+/g, ' ').trim() || '스마트판서';
    return t + '_' + new Date().toISOString().slice(0, 10);
  }
  function pngName() { return exportBaseName() + '_' + (state.pageIndex + 1) + '쪽.png'; }
  function pdfName() { return exportBaseName() + '.pdf'; }
  function zipName() { return exportBaseName() + '_전체PNG(' + state.pages.length + '쪽).zip'; }
  function canvasToPngBlob(cv) {
    return new Promise(function (resolve, reject) {
      try { cv.toBlob(function (b) { b ? resolve(b) : reject(new Error('PNG 생성 실패')); }, 'image/png'); }
      catch (e) { reject(e); }
    });
  }
  function makePngBlob() {
    return canvasToPngBlob(createCompositeCanvas());
  }
  async function makeAllPagesPngZipBlob() {
    if (typeof JSZip === 'undefined') {
      showToast('JSZip 라이브러리를 불러오지 못했습니다.');
      return null;
    }
    try {
      var total = state.pages.length;
      showToast('전체 ' + total + '쪽 PNG 압축 파일(ZIP)을 만드는 중...');
      var zip = new JSZip();
      var digits = Math.max(2, String(total).length);
      for (var i = 0; i < total; i++) {
        var pg = state.pages[i];
        if (pg.kind === 'pdf' && state.pdfDoc && !state.pdfCache[pg.pdfPageNum]) {
          state.pdfCache[pg.pdfPageNum] = await renderPdfPageToOffscreen(pg.pdfPageNum);
        }
        if (pg.kind === 'pptx' && pg.pptxSlide) {
          await ensurePptxSlideImages(pg.pptxSlide);
        }
        var comp = createCompositeCanvas(pg);
        var pBlob = await canvasToPngBlob(comp);
        var numStr = String(i + 1);
        while (numStr.length < digits) numStr = '0' + numStr;
        var cleanLbl = String(pg.label || '페이지').replace(/[\\/:*?"<>|]+/g, '_').trim().slice(0, 24) || '페이지';
        zip.file(numStr + '쪽_' + cleanLbl + '.png', pBlob);
      }
      return await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    } catch (err) {
      showToast('PNG 압축(ZIP) 생성 오류: ' + (err.message || err));
      return null;
    }
  }
  async function saveAllPngZip() {
    var blob = await makeAllPagesPngZipBlob();
    if (!blob) return;
    downloadBlob(blob, zipName());
    showToast('전체 ' + state.pages.length + '쪽 PNG 압축 파일(ZIP)을 저장했습니다.');
  }
  async function shareAllPngZip() {
    var blob = await makeAllPagesPngZipBlob();
    if (!blob) return;
    var file = new File([blob], zipName(), { type: 'application/zip' });
    if (canShareFile(file) && await shareFile(file, '전체 쪽 PNG 압축(ZIP)')) return;
    downloadBlob(blob, file.name);
    showToast('전체 ' + state.pages.length + '쪽 PNG 압축(ZIP) 파일을 내려받았습니다. 노션·드라이브·메신저에 바로 올려 공유하세요.');
  }
  function downloadBlob(blob, name) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { a.remove(); URL.revokeObjectURL(a.href); }, 5000);
  }
  async function saveCurrentPng() {
    try {
      downloadBlob(await makePngBlob(), pngName());
      showToast('현재 화면을 PNG 이미지로 저장했습니다. (Ctrl+P)');
    } catch (e) { showToast('PNG 저장 오류: ' + (e.message || e)); }
  }

  // 공유 — 휴대폰·태블릿은 공유 창(노션·카톡·드라이브…)으로, 컴퓨터는 이미지 복사(노션에 Ctrl+V) 또는 내려받기
  var pendingShare = null;
  function canShareFile(file) {
    try { return !!(navigator.share && navigator.canShare && navigator.canShare({ files: [file] })); } catch (e) { return false; }
  }
  async function shareFile(file, label) {
    try {
      await navigator.share({ files: [file], title: file.name });
      showToast(label + ' 공유 창을 열었습니다.');
      return true;
    } catch (e) {
      if (e && e.name === 'AbortError') return true; // 사용자가 공유 창을 닫음
      if (e && e.name === 'NotAllowedError') {
        // 파일 만드는 동안 시간이 지나 브라우저가 막은 경우 — 한 번 더 눌러 공유합니다.
        pendingShare = { file: file, label: label };
        if (el('sbShareReadyTitle')) el('sbShareReadyTitle').textContent = file.name;
        el('sbShareReadyModal').classList.add('show');
        return true;
      }
      return false;
    }
  }
  async function copyPngToClipboard(silentFail) {
    if (!navigator.clipboard || !window.ClipboardItem) { if (!silentFail) showToast('이 브라우저는 이미지 복사를 지원하지 않습니다. PNG로 저장한 뒤 노션에 끌어다 놓으세요.'); return false; }
    try {
      await navigator.clipboard.write([new ClipboardItem({ 'image/png': makePngBlob() })]);
      showToast('현재 화면을 이미지로 복사했습니다. 노션·구글 문서 등에서 Ctrl+V로 붙여넣으세요.');
      return true;
    } catch (e) {
      if (!silentFail) showToast('이미지 복사에 실패했습니다. PNG로 저장한 뒤 노션에 끌어다 놓으세요.');
      return false;
    }
  }
  async function sharePng() {
    var blob;
    try { blob = await makePngBlob(); } catch (e) { showToast('PNG 생성 오류: ' + (e.message || e)); return; }
    var file = new File([blob], pngName(), { type: 'image/png' });
    if (canShareFile(file) && await shareFile(file, 'PNG')) return;
    // 공유 창이 없는 컴퓨터: 이미지 복사로 대신합니다(노션에 바로 붙여넣기).
    if (await copyPngToClipboard(true)) return;
    downloadBlob(blob, file.name);
    showToast('이 브라우저는 공유를 지원하지 않아 PNG로 내려받았습니다. 노션에 끌어다 놓으면 됩니다.');
  }
  async function sharePdf() {
    var blob = await makePdfBlob();
    if (!blob) return;
    var file = new File([blob], pdfName(), { type: 'application/pdf' });
    if (canShareFile(file) && await shareFile(file, 'PDF')) return;
    downloadBlob(blob, file.name);
    showToast('이 브라우저는 파일 공유를 지원하지 않아 PDF로 내려받았습니다. 노션에 끌어다 놓거나 첨부하세요.');
  }

  async function saveMergedPdf() {
    var blob = await makePdfBlob();
    if (!blob) return;
    downloadBlob(blob, pdfName());
    showToast('전체 ' + state.pages.length + '페이지 판서 합본 PDF를 저장했습니다.');
  }

  async function makePdfBlob() {
    if (typeof PDFLib === 'undefined') { showToast('pdf-lib 라이브러리를 불러오지 못했습니다.'); return null; }
    try {
      showToast('전체 ' + state.pages.length + '페이지 판서 합본 PDF를 만드는 중...');
      var pdfDoc = await PDFLib.PDFDocument.create();
      for (var i = 0; i < state.pages.length; i++) {
        var pg = state.pages[i];
        if (pg.kind === 'pdf' && state.pdfDoc && !state.pdfCache[pg.pdfPageNum]) {
          state.pdfCache[pg.pdfPageNum] = await renderPdfPageToOffscreen(pg.pdfPageNum);
        }
        if (pg.kind === 'pptx' && pg.pptxSlide) {
          await ensurePptxSlideImages(pg.pptxSlide);
        }
        var comp = createCompositeCanvas(pg);
        var pngImg = await pdfDoc.embedPng(comp.toDataURL('image/png'));
        var pdfPage = pdfDoc.addPage([pg.width * 0.6, pg.height * 0.6]);
        pdfPage.drawImage(pngImg, { x: 0, y: 0, width: pg.width * 0.6, height: pg.height * 0.6 });
      }
      var bytes = await pdfDoc.save();
      return new Blob([bytes], { type: 'application/pdf' });
    } catch (err) {
      showToast('합본 PDF 만들기 오류: ' + (err.message || err));
      return null;
    }
  }

  // ===================================================================
  // 이벤트 바인딩 & 단축키
  // ===================================================================
  function init() {
    state.pages = [createBlankPage('dark')];
    updateNoteBadge();
    resizeBoardToPage();
    renderBackground();
    renderInk();

    // 마지막 판서로 돌아오기 + 1분 자동 저장 + 화면을 떠날 때 저장
    resumeLastNote().then(function () { autoSaveSig = boardSignature(); });
    setInterval(function () { autoSaveNow('interval'); }, AUTO_SAVE_MS);
    document.addEventListener('visibilitychange', function () { if (document.visibilityState === 'hidden') autoSaveNow('hidden'); });
    window.addEventListener('pagehide', function () { rememberLastPage(); autoSaveNow('pagehide'); });
    document.querySelectorAll('a[href="./index.html"]').forEach(function (a) {
      a.addEventListener('click', function (e) { e.preventDefault(); goHomeSaved(); });
    });

    function handleViewportChange() {
      resizeBoardToPage();
      renderBackground();
      renderInk();
      setTimeout(fitViewport, 80);
      setTimeout(fitViewport, 260);
      setTimeout(fitViewport, 520);
    }

    window.addEventListener('resize', handleViewportChange);
    window.addEventListener('orientationchange', handleViewportChange);
    document.addEventListener('fullscreenchange', handleViewportChange);
    if (window.visualViewport) {
      window.visualViewport.addEventListener('resize', fitViewport);
    }

    // 상단 팔레트 접기/펴기 & 자동 숨김 센서
    if (el('sbFoldPaletteBtn')) el('sbFoldPaletteBtn').addEventListener('click', function () { togglePalette(true); });
    if (el('sbOpenPaletteBtn')) el('sbOpenPaletteBtn').addEventListener('click', function () { togglePalette(false); });
    if (el('sbTopHoverSensor')) {
      el('sbTopHoverSensor').addEventListener('pointerenter', function () {
        if (el('sbPalette') && el('sbPalette').classList.contains('collapsed')) {
          togglePalette(false);
        }
      });
    }
    if (el('sbAutoHideBtn')) {
      el('sbAutoHideBtn').addEventListener('click', function () {
        state.autoHide = !state.autoHide;
        this.classList.toggle('active', state.autoHide);
        this.textContent = state.autoHide ? '자동숨김 ON' : '자동숨김 OFF';
        showToast(state.autoHide ? '상단 팔레트 자동 숨김 ON (판서 시작 시 자동으로 위로 숨겨집니다)' : '상단 팔레트 자동 숨김 OFF (항상 상단에 고정)');
      });
    }
    if (el('sbFillScreenModeBtn')) {
      el('sbFillScreenModeBtn').addEventListener('click', function () {
        state.fillScreen = !state.fillScreen;
        this.classList.toggle('active', state.fillScreen);
        this.textContent = state.fillScreen ? '꽉찬화면: ON' : '꽉찬화면: OFF';
        resizeBoardToPage();
        renderBackground();
        renderInk();
        showToast(state.fillScreen ? '화면 송출용 100% 꽉 찬 화면 모드 ON' : '기본 비율 화면 모드');
      });
    }
    if (el('sbFloatingHelpBtn')) {
      el('sbFloatingHelpBtn').addEventListener('click', function () {
        openHelpGuide();
      });
    }
    if (el('sbFloatingFsBtn')) {
      el('sbFloatingFsBtn').addEventListener('click', function () {
        if (el('sbFullscreenBtn')) el('sbFullscreenBtn').click();
      });
    }

    // 삼성노트형 폴더 보관함
    if (el('sbOpenVaultBtn')) el('sbOpenVaultBtn').addEventListener('click', async function () {
      updateNoteBadge();
      await renderVaultUi();
      el('sbVaultModal').classList.add('show');
    });
    if (el('sbQuickSaveNoteBtn')) el('sbQuickSaveNoteBtn').addEventListener('click', function () { serializeCurrentNote(false); });
    if (el('sbVaultSaveCurBtn')) el('sbVaultSaveCurBtn').addEventListener('click', function () { serializeCurrentNote(false); });
    if (el('sbVaultSortSelect')) el('sbVaultSortSelect').addEventListener('change', function () {
      try { localStorage.setItem(NOTE_SORT_KEY, this.value === 'manual' ? 'manual' : 'recent'); } catch (e) {}
      renderVaultUi();
    });
    if (el('sbVaultSaveNewBtn')) el('sbVaultSaveNewBtn').addEventListener('click', function () { serializeCurrentNote(true); });
    if (el('sbVaultCopyOkBtn')) el('sbVaultCopyOkBtn').addEventListener('click', copyNote);
    if (el('sbVaultCopyCancelBtn')) el('sbVaultCopyCancelBtn').addEventListener('click', function () { el('sbVaultCopyBox').hidden = true; copyTarget = null; });
    if (el('sbAddFolderBtn')) el('sbAddFolderBtn').addEventListener('click', function () {
      var row = el('sbNewFolderRow');
      if (row) row.style.display = row.style.display === 'none' ? 'flex' : 'none';
      if (el('sbNewFolderNameInput')) el('sbNewFolderNameInput').focus();
    });
    if (el('sbConfirmAddFolderBtn')) el('sbConfirmAddFolderBtn').addEventListener('click', function () {
      var inp = el('sbNewFolderNameInput');
      var name = (inp && inp.value.trim()) || '';
      if (!name) return;
      var list = loadFolders();
      if (list.indexOf(name) < 0) {
        list.push(name);
        saveFolders(list);
      }
      state.currentFolder = name;
      state.activeVaultFolder = name;
      inp.value = '';
      if (el('sbNewFolderRow')) el('sbNewFolderRow').style.display = 'none';
      renderVaultUi();
      showToast('새 폴더 [' + name + ']를 추가했습니다.');
    });
    if (el('sbVaultExportJsonBtn')) el('sbVaultExportJsonBtn').addEventListener('click', function () {
      exportAllVaultJson();
    });
    if (el('sbVaultExportCurJsonBtn')) el('sbVaultExportCurJsonBtn').addEventListener('click', function () {
      exportSingleNoteJson(null);
    });
    if (el('sbVaultImportJsonInput')) el('sbVaultImportJsonInput').addEventListener('change', function () {
      var f = this.files && this.files[0];
      if (f) importVaultJsonFile(f);
      this.value = '';
    });
    // 기존 경로(/smart-board.html)에서 접속한 경우 새 독립 판서보드(board-tong) 이동 버튼 자동 노출
    (function () {
      var isLegacyBoard = /smart-board\.html/i.test(location.pathname || '') || /^(t-ong|neis-help|reponeishelp)\./i.test(location.hostname || '');
      if (isLegacyBoard) {
        if (el('sbVaultGoBoardTongWrap')) el('sbVaultGoBoardTongWrap').style.display = 'flex';
        if (el('sbMenuBoardTongItem')) el('sbMenuBoardTongItem').style.display = 'flex';
      }
    })();

    // 내장 웹 브라우저 + 유튜브 + 광고제거 안내 + 웹 위 스탬프/필기 빠른 전환
    if (el('sbOpenWebBtn')) el('sbOpenWebBtn').addEventListener('click', function () { openWebBrowser(); });
    if (el('sbWebGoBtn')) el('sbWebGoBtn').addEventListener('click', function () { openWebBrowser(el('sbWebUrlInput').value); });
    if (el('sbWebUrlInput')) el('sbWebUrlInput').addEventListener('keydown', function (e) { if (e.key === 'Enter') openWebBrowser(this.value); });
    document.querySelectorAll('[data-web-url]').forEach(function (wb) {
      wb.addEventListener('click', function () { openWebBrowser(wb.getAttribute('data-web-url')); });
    });
    if (el('sbWebYoutubePresetBtn')) {
      el('sbWebYoutubePresetBtn').addEventListener('click', function () {
        var inp = el('sbWebUrlInput');
        var curVal = inp ? inp.value.trim() : '';
        if (convertYoutubeToEmbedUrl(curVal)) {
          openWebBrowser(curVal);
        } else {
          if (inp) {
            inp.value = 'https://www.youtube.com/watch?v=aircAruvnKk';
            inp.focus();
            inp.select();
          }
          openWebBrowser('https://www.youtube.com/watch?v=aircAruvnKk');
          showToast('유튜브 샘플 영상을 열었습니다. 원하는 유튜브 링크를 주소창에 붙여넣고 [열기]를 누르세요.');
        }
      });
    }
    var adSlideIdx = 0;
    var adShowAll = false;
    function goAdSlide(nextIdx) {
      var panes = document.querySelectorAll('[data-ad-slide-pane]');
      if (!panes.length) return;
      adShowAll = false;
      adSlideIdx = (nextIdx + panes.length) % panes.length;
      panes.forEach(function (p, i) { p.style.display = i === adSlideIdx ? 'flex' : 'none'; });
      document.querySelectorAll('[data-ad-slide]').forEach(function (tb) {
        tb.classList.toggle('active', parseInt(tb.getAttribute('data-ad-slide'), 10) === adSlideIdx);
      });
      if (el('sbAdSlideCounter')) el('sbAdSlideCounter').textContent = (adSlideIdx + 1) + ' / ' + panes.length;
      if (el('sbAdShowAllToggleBtn')) el('sbAdShowAllToggleBtn').textContent = '전체 보기';
    }
    function openAdBlockGuide() {
      goAdSlide(0);
      if (el('sbAdBlockModal')) el('sbAdBlockModal').classList.add('show');
    }
    if (el('sbAdBlockGuideBtn')) el('sbAdBlockGuideBtn').addEventListener('click', openAdBlockGuide);
    if (el('sbTopAdBlockGuideBtn')) el('sbTopAdBlockGuideBtn').addEventListener('click', openAdBlockGuide);
    if (el('sbAdPrevSlideBtn')) el('sbAdPrevSlideBtn').addEventListener('click', function () { goAdSlide(adSlideIdx - 1); });
    if (el('sbAdNextSlideBtn')) el('sbAdNextSlideBtn').addEventListener('click', function () { goAdSlide(adSlideIdx + 1); });
    document.querySelectorAll('[data-ad-slide]').forEach(function (tb) {
      tb.addEventListener('click', function () { goAdSlide(parseInt(tb.getAttribute('data-ad-slide'), 10) || 0); });
    });
    if (el('sbAdShowAllToggleBtn')) {
      el('sbAdShowAllToggleBtn').addEventListener('click', function () {
        var panes = document.querySelectorAll('[data-ad-slide-pane]');
        adShowAll = !adShowAll;
        if (adShowAll) {
          panes.forEach(function (p) { p.style.display = 'flex'; });
          this.textContent = '1장씩 슬라이드로 보기';
        } else {
          goAdSlide(adSlideIdx);
        }
      });
    }
    if (el('sbWebInteractToggleBtn')) {
      el('sbWebInteractToggleBtn').addEventListener('click', function () {
        setWebInteractMode(true);
        if (webFrame) webFrame.focus();
        showToast(state.webSlides ? '슬라이드 조작: 화면 왼쪽 아래 ◀ ▶ 또는 키보드 ← →로 넘기세요. 필기는 [펜 필기]로 전환합니다.' : '웹 조작 모드: 웹페이지 스크롤 및 유튜브 재생 버튼 클릭이 가능합니다.');
      });
    }
    if (el('sbWebPenBtn')) {
      el('sbWebPenBtn').addEventListener('click', function () {
        setWebInteractMode(false);
        setTool('pen');
        showToast('웹 위 [펜 필기] 모드로 전환했습니다.');
      });
    }
    if (el('sbWebCircleStampBtn')) {
      el('sbWebCircleStampBtn').addEventListener('click', function () {
        setWebInteractMode(false);
        setTool('stamp');
        showToast('웹 위 [① 원형 스탬프] 모드입니다. 웹 화면을 클릭하면 ①, ②, ③이 찍힙니다.');
      });
    }
    if (el('sbWebSquareStampBtn')) {
      el('sbWebSquareStampBtn').addEventListener('click', function () {
        setWebInteractMode(false);
        setTool('sqStamp');
        showToast('웹 위 [▣ 사각 스탬프] 모드입니다. 웹 화면을 클릭하면 [1], [2], [3]이 찍힙니다.');
      });
    }
    if (el('sbWebEraserBtn')) {
      el('sbWebEraserBtn').addEventListener('click', function () {
        setWebInteractMode(false);
        setTool('eraser');
        showToast('웹 위 [지우개] 모드로 전환했습니다.');
      });
    }
    if (el('sbWebSnapBtn')) el('sbWebSnapBtn').addEventListener('click', snapshotWebToBoard);
    if (el('sbWebOpenTabBtn')) el('sbWebOpenTabBtn').addEventListener('click', function () {
      var u = (el('sbWebUrlInput') && el('sbWebUrlInput').value.trim()) || 'https://www.geogebra.org';
      window.open(u, '_blank', 'noopener');
    });
    if (el('sbWebCloseBtn')) el('sbWebCloseBtn').addEventListener('click', closeWebBrowser);

    // 모드 전환
    document.querySelectorAll('.sb-mode-btn').forEach(function (b) {
      b.addEventListener('click', function () {
        var m = b.getAttribute('data-sb-mode');
        if (m === 'zoom') toggleZoom(); else setMode(m);
      });
    });

    // 도구 선택
    document.querySelectorAll('.sb-tool-btn[data-sb-tool]').forEach(function (b) {
      b.addEventListener('click', function () { setTool(b.getAttribute('data-sb-tool')); });
    });

    // 색상 & 굵기
    document.querySelectorAll('#sbColorGroup .sb-color-dot').forEach(function (d) {
      d.addEventListener('click', function () { setColor(d.getAttribute('data-color')); });
    });
    if (el('sbSizeDecBtn')) el('sbSizeDecBtn').addEventListener('click', function () { stepSize(-1); });
    if (el('sbSizeIncBtn')) el('sbSizeIncBtn').addEventListener('click', function () { stepSize(1); });
    document.querySelectorAll('.sb-size-preset-btn[data-size-preset]').forEach(function (b) {
      b.addEventListener('click', function () {
        setSize(parseFloat(b.getAttribute('data-size-preset')) || 2);
        closeAllPopovers();
      });
    });
    document.querySelectorAll('.sb-pen-opt[data-pen-style]').forEach(function (b) {
      b.addEventListener('click', function () {
        setPenStyle(b.getAttribute('data-pen-style'));
        closeAllPopovers();
      });
    });

    // 배경 전환
    document.querySelectorAll('[data-sb-bg]').forEach(function (b) {
      b.addEventListener('click', function () { setBlankBgMode(b.getAttribute('data-sb-bg')); });
    });

    // S펜 옆 버튼 & 에어액션 → [지우개] <-> [펜] 양방향 스위치 토글
    if (el('sbSpenEraserModeBtn')) {
      el('sbSpenEraserModeBtn').addEventListener('click', function () {
        state.spenEraserOnBtn = !state.spenEraserOnBtn;
        this.classList.toggle('active', state.spenEraserOnBtn);
        this.textContent = state.spenEraserOnBtn ? 'S펜버튼: 지우개↔펜 스위치' : 'S펜버튼: 꺼짐';
        showToast(state.spenEraserOnBtn
          ? 'S펜 버튼 양방향 스위치 ON (펜 버튼을 누르면 [지우개], 다시 누르면 [펜]으로 전환됩니다)'
          : 'S펜 옆 버튼 스위치 OFF');
      });
    }

    // 갤럭시탭(안드로이드 크롬/삼성 인터넷)에서 S펜 옆 버튼을 누를 때 발생하는 contextmenu / auxclick 가로채기 → [지우개] <-> [펜] 스위치
    liveCanvas.addEventListener('contextmenu', function (ev) {
      ev.preventDefault();
      if (state.spenEraserOnBtn) {
        activateSpenEraserButton(true, 'S펜 버튼 스위치');
      }
    });
    liveCanvas.addEventListener('auxclick', function (ev) {
      if (ev.button === 2 || ev.button === 5) {
        ev.preventDefault();
        if (state.spenEraserOnBtn) {
          activateSpenEraserButton(true, 'S펜 버튼 스위치');
        }
      }
    });

    // S펜 전용 토글
    if (el('sbPalmRejectBtn')) {
      el('sbPalmRejectBtn').addEventListener('click', function () {
        setPalmReject(!state.palmReject);
        state.spenManualOff = !state.palmReject;
        showToast(state.palmReject ? 'S펜 전용 모드 ON (손바닥 무시 · 손가락으로 확대 이동)' : 'S펜 전용 모드 OFF (손가락·펜 모두 필기)');
      });
    }

    if (el('sbSpenAutoBtn')) {
      var syncSpenAutoBtn = function () {
        el('sbSpenAutoBtn').classList.toggle('active', state.spenAuto);
        el('sbSpenAutoBtn').textContent = state.spenAuto ? 'S펜자동: 켜짐' : 'S펜자동: 꺼짐';
      };
      syncSpenAutoBtn();
      el('sbSpenAutoBtn').addEventListener('click', function () {
        state.spenAuto = !state.spenAuto;
        try { localStorage.setItem('sbSpenAuto_v1', state.spenAuto ? 'on' : 'off'); } catch (e) {}
        syncSpenAutoBtn();
        showToast(state.spenAuto ? 'S펜 자동 ON — S펜이 닿으면 S펜 전용 모드가 자동으로 켜집니다.' : 'S펜 자동 OFF — S펜 전용 모드는 직접 켜고 끕니다.');
      });
    }

    // 버전 · 수정 이력 · 이용 현황
    var SB_VERSION = '2.8.2';
    var SB_CHANGELOG = [
      { v: '2.8.2', d: '2026-11-27', items: [
        '사용설명서 새 단장: 왼쪽 목차에서 고르기 · 큰 영상 아래 설명 · 처음 사용자 안내',
        '기능별 음성 해설 영상 16편 (v2.8 단축키 · 백업·복원 · 공유 · S펜 설정까지)',
        '단축키 핸드아웃(그림 · A4 PDF) — 휴대폰에 저장하거나 인쇄해서 교탁에'
      ] },
      { v: '2.8.1', d: '2026-11-27', items: [
        '위쪽 첫 줄에 오늘 방문자 · 총 이용 표시 (누르면 버전 · 이용 현황)'
      ] },
      { v: '2.8.0', d: '2026-11-27', items: [
        '상단 [백업·복원 ▾] 메뉴 및 노트함 전체·개별 판서 노트 백업(.json) / 비파괴 병합 복원 기능 추가',
        '단축키 개선: 뒤로 취소(Z · Esc · ,) / 앞으로 다시 실행(X · .), 방향키(← → · ↑ ↓) 슬라이드·페이지 이동, 레이저(D)',
        '밑줄(직선)·원 도형 좌측 상단 자동 번호 스탬프 부착(자동① ON · 단축키 A) 및 슬라이드 전환·SS/0 스탬프 1번 초기화',
        '전체 쪽 PNG 압축(ZIP) 저장 및 공유 메뉴 추가 · PDF 가로 맞춤(Ctrl+1) 중앙 정렬 및 터치 이동 떨림 제거',
        '일반 펜 기본 2pt · 형광펜 기본 6pt · 삼성노트 질감 연필 모드 · 형광펜 0.5초 멈춤 도형화 · 3초 다중 레이저 포인터',
        '독립 도메인(board-tong / boardtong.vercel.app) 분리 배포 및 시각 사용설명서 업그레이드'
      ] },
      { v: '2.7.2', d: '2026-11-25', items: [
        'PNG 저장 단축키 Ctrl+P',
        '공유 ▾ 버튼: 현재 화면 PNG 공유 · 이미지 복사(노션에 Ctrl+V) · 전체 쪽 PDF 공유',
        '저장·공유 파일 이름에 노트 제목 사용'
      ] },
      { v: '2.7.1', d: '2026-11-25', items: [
        '스탬프 전체 지우기 단축키를 Ctrl+S로 변경 (PNG 저장은 PNG·PDF ▾ 버튼)',
        'Ctrl+1을 두 번 누르면 확대 전 모드·도구(형광펜·레이저 등) 그대로 돌아오기'
      ] },
      { v: '2.7.0', d: '2026-11-25', items: [
        '판서 1분마다 자동 저장 · 다른 앱으로 갔다 오거나 홈으로 나갈 때도 저장',
        '스마트 판서를 다시 열면 마지막으로 쓰던 노트·쪽으로 바로 돌아오기',
        '노트함을 최근에 연 순서로 정렬 (직접 정한 순서도 선택 가능)',
        'S펜 자동: S펜이 닿으면 S펜 전용 모드가 자동으로 켜짐',
        '단축키 변경: ESC = 되돌리기(열린 창이 있으면 닫기), Backspace = 저장 후 홈으로',
        '모드 단축키를 Ctrl로: Ctrl+` 게시 · Ctrl+1 확대 · Ctrl+2 판서 · Ctrl+3 편집 (Alt도 계속 동작)',
        '그리다가 꾸욱: 별 모양도 반듯한 별로 바뀜',
        '설정에 버전 · 수정 이력 · 방문자 수 표시'
      ] },
      { v: '2.6.6', d: '2026-11-24', items: [
        '지우개 ▾ 메뉴: 펜 글씨 전체(Shift+E) · 스탬프 전체(Shift+S) · 전체 필기(Ctrl+E) 지우기',
        '그리다가 꾸욱 멈추면 직선 · 삼각형 · 사각형 · 원으로 자동 변환',
        'PPTX 애니메이션 호환 개선: 그룹 · 자동 재생 · 나타남/사라짐 반복 · 문단 자리 · 테마 도형 색'
      ] },
      { v: '2.6.5', d: '2026-11-24', items: [
        '이동 도구(M · Move)와 단축키 영어 설명(Pen, Eraser …) 표시',
        '노트함 순서 바꾸기 · 다른 학급 폴더로 노트 복사(자료만/판서까지)',
        '사용설명서 10장 확대 · 이동 추가'
      ] },
      { v: '2.6.4', d: '2026-11-24', items: [
        '확대 화면 검은 여백 수정 · 두 손가락 확대/이동 · 세 손가락 이동',
        '펜 글씨 · 스탬프 전체 지우기와 한 번에 되돌리기'
      ] },
      { v: '2.6.3', d: '2026-11-24', items: [
        '영상 사용설명서 12장 · 단축키 배지 표시',
        '번호 스탬프 빈 번호 채우기 · 편집 막대 가림 수정'
      ] }
    ];
    if (el('sbVersionLabel')) el('sbVersionLabel').textContent = 'v' + SB_VERSION;
    if (el('sbChangelogVersion')) el('sbChangelogVersion').textContent = 'v' + SB_VERSION;
    if (el('sbChangelogBody')) {
      el('sbChangelogBody').innerHTML = SB_CHANGELOG.map(function (r) {
        return '<div class="sb-changelog-item"><h4>v' + r.v + '<small>' + r.d + '</small></h4><ul>' +
          r.items.map(function (t) { return '<li>' + t.replace(/</g, '&lt;') + '</li>'; }).join('') + '</ul></div>';
      }).join('');
    }
    if (el('sbChangelogBtn')) el('sbChangelogBtn').addEventListener('click', function () {
      closeAllPopovers();
      el('sbChangelogModal').classList.add('show');
    });
    if (el('sbTopVisitorPill')) el('sbTopVisitorPill').addEventListener('click', function (e) {
      e.stopPropagation();
      var moreBtn = el('sbBoardMoreBtn'), moreMenu = el('sbBoardMoreMenu');
      if (moreBtn && moreMenu) {
        var willOpen = !moreMenu.classList.contains('show');
        closeAllPopovers();
        if (willOpen) showAnchoredPopover(moreBtn, moreMenu, true);
      }
    });
    (async function recordSbVisit() {
      var h0 = location.hostname;
      if (h0 === 'localhost' || h0 === '127.0.0.1' || /^192\.168\./.test(h0) || /^10\./.test(h0)) {
        if (el('sbUsageStatus')) el('sbUsageStatus').textContent = '로컬 실행 중에는 방문 수를 집계하지 않습니다.';
        return;
      }
      var apiBase = /^t-ong\.vercel\.app$/i.test(h0) ? '' : 'https://t-ong.vercel.app';
      var SB_CACHE_STATS_KEY = 'sbLastVisitCounts_v1';
      function applySbCounts(data, fromCache) {
        var rawTd = Number(data && (data.todayVisitors != null ? data.todayVisitors : data.d) || 0);
        var rawTt = Number(data && (data.totalVisits != null ? data.totalVisits : data.v) || 0);
        if (!rawTd && !rawTt) return false;
        var td = rawTd.toLocaleString('ko-KR');
        var tt = rawTt.toLocaleString('ko-KR');
        if (el('sbTodayVisitors')) el('sbTodayVisitors').textContent = td;
        if (el('sbTotalVisits')) el('sbTotalVisits').textContent = tt;
        if (el('sbTopTodayVisitors')) el('sbTopTodayVisitors').textContent = td;
        if (el('sbTopTotalVisits')) el('sbTopTotalVisits').textContent = tt;
        if (el('sbUsageStatus')) el('sbUsageStatus').textContent = '방문 수만 익명으로 집계 · 판서 내용은 이 브라우저에만 저장';
        if (!fromCache) {
          try { localStorage.setItem(SB_CACHE_STATS_KEY, JSON.stringify({ d: rawTd, v: rawTt })); } catch (e) {}
        }
        return true;
      }
      try {
        var cachedSb = JSON.parse(localStorage.getItem(SB_CACHE_STATS_KEY) || 'null');
        if (cachedSb) applySbCounts(cachedSb, true);
      } catch (e0) {}
      async function attemptRecord(isRetry) {
        try {
          var vid = localStorage.getItem('notetongAnonymousVisitor_v1');
          if (!vid) {
            vid = (window.crypto && crypto.randomUUID)
              ? crypto.randomUUID().replace(/-/g, '')
              : String(Date.now()) + Math.random().toString(36).slice(2);
            localStorage.setItem('notetongAnonymousVisitor_v1', vid);
          }
          var parts = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Seoul', year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(new Date());
          var mp = {}; parts.forEach(function (p) { mp[p.type] = p.value; });
          var raw = 'smartboard:' + mp.year + '-' + mp.month + '-' + mp.day + ':' + vid;
          var hash = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(raw));
          var token = Array.from(new Uint8Array(hash)).map(function (x) { return x.toString(16).padStart(2, '0'); }).join('');
          var urls = [
            { u: apiBase + '/api/classroom-hello', opts: { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify({ t: token, m: 'smartboard' }), cache: 'no-store' } },
            { u: apiBase + '/api/classroom-hello?m=smartboard&t=' + encodeURIComponent(token) + '&_=' + Date.now(), opts: { method: 'GET', cache: 'no-store' } },
            { u: apiBase + '/api/usage', opts: { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify({ visitorToken: token, scope: 'smartboard' }), cache: 'no-store' } }
          ];
          for (var i = 0; i < urls.length; i++) {
            try {
              var res = await fetch(urls[i].u, urls[i].opts);
              if (res && res.ok && applySbCounts(await res.json(), false)) return;
            } catch (inner) {}
          }
          throw new Error('sb_stats_failed');
        } catch (e) {
          if (!isRetry) setTimeout(function () { attemptRecord(true); }, 2500);
          else if (el('sbUsageStatus')) el('sbUsageStatus').textContent = '통계 서버 연결 대기 중 · 판서 내용은 전송하지 않습니다.';
        }
      }
      attemptRecord(false);
    })();

    // 통합 파일·링크 열기 모달 (#sbFileHubModal)
    function closeFileHubModal() {
      if (el('sbFileHubModal')) el('sbFileHubModal').classList.remove('show');
    }
    function handleUnifiedFile(f) {
      if (!f) return;
      closeFileHubModal();
      var name = String(f.name || '').toLowerCase();
      if (name.endsWith('.json') || f.type === 'application/json') {
        importVaultJsonFile(f);
      } else if (name.endsWith('.pdf') || f.type === 'application/pdf') {
        loadPdfFile(f);
      } else if (name.endsWith('.pptx')) {
        loadPptxFile(f);
      } else if (f.type && f.type.indexOf('image/') === 0 || /\.(png|jpe?g|gif|webp|bmp)$/i.test(name)) {
        var fr = new FileReader();
        fr.onload = function () { loadImageFromDataUrl(fr.result, f.name); };
        fr.readAsDataURL(f);
      } else {
        showToast('지원 형식은 PDF, PPTX, 이미지(PNG/JPG), 판서 백업(.json)입니다.');
      }
    }
    if (el('sbHubBackupJsonInput')) {
      el('sbHubBackupJsonInput').addEventListener('change', function () {
        if (this.files && this.files[0]) {
          closeFileHubModal();
          importVaultJsonFile(this.files[0]);
        }
        this.value = '';
      });
    }
    if (stageWrap) {
      stageWrap.addEventListener('dragover', function (e) {
        if (e.dataTransfer && e.dataTransfer.types && Array.prototype.indexOf.call(e.dataTransfer.types, 'Files') >= 0) {
          e.preventDefault();
        }
      });
      stageWrap.addEventListener('drop', function (e) {
        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]) {
          e.preventDefault();
          handleUnifiedFile(e.dataTransfer.files[0]);
        }
      });
    }
    if (el('sbOpenFileHubBtn')) {
      el('sbOpenFileHubBtn').addEventListener('click', function () {
        if (el('sbFileHubModal')) el('sbFileHubModal').classList.add('show');
      });
    }
    if (el('sbUnifiedFileInput')) {
      el('sbUnifiedFileInput').addEventListener('change', function () {
        if (this.files && this.files[0]) handleUnifiedFile(this.files[0]);
        this.value = '';
      });
    }
    if (el('sbFileHubDropzone')) {
      var dz = el('sbFileHubDropzone');
      dz.addEventListener('dragover', function (e) { e.preventDefault(); dz.classList.add('dragover'); });
      dz.addEventListener('dragleave', function () { dz.classList.remove('dragover'); });
      dz.addEventListener('drop', function (e) {
        e.preventDefault();
        dz.classList.remove('dragover');
        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]) {
          handleUnifiedFile(e.dataTransfer.files[0]);
        }
      });
    }
    if (el('sbFileHubOpenDriveUrlBtn')) {
      el('sbFileHubOpenDriveUrlBtn').addEventListener('click', function () {
        var u = (el('sbFileHubDriveUrlInput') && el('sbFileHubDriveUrlInput').value || '').trim();
        if (!u) {
          showToast('구글 드라이브 공유 링크를 먼저 붙여넣어 주세요.');
          if (el('sbFileHubDriveUrlInput')) el('sbFileHubDriveUrlInput').focus();
          return;
        }
        closeFileHubModal();
        openWebBrowser(u);
      });
    }
    if (el('sbFileHubDriveUrlInput')) {
      el('sbFileHubDriveUrlInput').addEventListener('keydown', function (e) {
        if (e.key === 'Enter') {
          e.preventDefault();
          if (el('sbFileHubOpenDriveUrlBtn')) el('sbFileHubOpenDriveUrlBtn').click();
        }
      });
    }
    if (el('sbFileHubOpenUrlBtn')) {
      el('sbFileHubOpenUrlBtn').addEventListener('click', function () {
        var u = (el('sbFileHubUrlInput') && el('sbFileHubUrlInput').value || '').trim();
        closeFileHubModal();
        openWebBrowser(u || 'https://www.youtube.com/watch?v=aircAruvnKk');
      });
    }
    if (el('sbFileHubUrlInput')) {
      el('sbFileHubUrlInput').addEventListener('keydown', function (e) {
        if (e.key === 'Enter') {
          e.preventDefault();
          var u = (this.value || '').trim();
          closeFileHubModal();
          openWebBrowser(u);
        }
      });
    }
    document.querySelectorAll('[data-filehub-url]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        closeFileHubModal();
        openWebBrowser(btn.getAttribute('data-filehub-url'));
      });
    });
    document.querySelectorAll('[data-filehub-yt]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        closeFileHubModal();
        var u = (el('sbFileHubUrlInput') && el('sbFileHubUrlInput').value || '').trim();
        if (u && /youtube\.com|youtu\.be/i.test(u)) openWebBrowser(u);
        else openWebBrowser('https://www.youtube.com/watch?v=aircAruvnKk');
      });
    });

    // 파일 입력
    if (el('sbPdfInput')) el('sbPdfInput').addEventListener('change', function () { if (this.files && this.files[0]) { closeFileHubModal(); loadPdfFile(this.files[0]); } this.value = ''; });
    if (el('sbPptxInput')) el('sbPptxInput').addEventListener('change', function () { if (this.files && this.files[0]) { closeFileHubModal(); loadPptxFile(this.files[0]); } this.value = ''; });
    if (el('sbImgInput')) el('sbImgInput').addEventListener('change', function () {
      var f = this.files && this.files[0];
      if (!f) return;
      closeFileHubModal();
      var fr = new FileReader();
      fr.onload = function () { loadImageFromDataUrl(fr.result, f.name); };
      fr.readAsDataURL(f);
      this.value = '';
    });

    document.addEventListener('paste', function (e) {
      if (document.activeElement && (document.activeElement.tagName === 'TEXTAREA' || document.activeElement.tagName === 'INPUT')) return;
      var items = (e.clipboardData && e.clipboardData.items) || [];
      for (var i = 0; i < items.length; i++) {
        if (items[i].type && items[i].type.indexOf('image/') === 0) {
          var blob = items[i].getAsFile();
          var fr = new FileReader();
          fr.onload = function () { loadImageFromDataUrl(fr.result, '붙여넣은 이미지'); };
          fr.readAsDataURL(blob);
          break;
        }
      }
    });

    // 되돌리기 / 다시실행 / 전체삭제
    // '전체 지우기' 직전 상태 — 노트 저장에 섞이지 않도록 쪽 객체 밖(WeakMap)에 둡니다.

    function performUndo(keyLabel) {
      var pg = curPage();
      // 바로 앞이 지우개/전체 지우기였다면 지운 것을 한꺼번에 되살립니다.
      var st = undoSnaps.get(pg) || [];
      while (st.length && st[st.length - 1].after !== pg.objects) st.pop();
      var bulk = st[st.length - 1];
      if (bulk && pg.objects.length === bulk.afterLen) {
        var rst = redoSnaps.get(pg);
        if (!rst) { rst = []; redoSnaps.set(pg, rst); }
        rst.push({ before: bulk.before, beforeLen: bulk.before.length, after: bulk.after, afterLen: bulk.afterLen });
        pg.objects = bulk.before;
        st.pop();
        clearSelection();
        renderInk();
        showToast(keyLabel ? ('지운 판서를 되살렸습니다. (' + keyLabel + ')') : '지운 판서를 되살렸습니다.');
        return true;
      }
      if (!pg.objects.length) {
        if (keyLabel) showToast('취소할 직전 필기가 없습니다.');
        return false;
      }
      var popped = pg.objects.pop();
      pg.redoStack.push(popped);
      if (popped && popped.groupId) {
        while (pg.objects.length && pg.objects[pg.objects.length - 1].groupId === popped.groupId) {
          pg.redoStack.push(pg.objects.pop());
        }
      }
      clearSelection();
      renderInk();
      if (keyLabel) showToast('직전 필기를 취소했습니다. (' + keyLabel + ')');
      return true;
    }

    function performRedo(keyLabel) {
      var pg = curPage();
      var rst = redoSnaps.get(pg) || [];
      while (rst.length && (rst[rst.length - 1].before !== pg.objects || rst[rst.length - 1].beforeLen !== pg.objects.length)) {
        rst.pop();
      }
      var rBulk = rst[rst.length - 1];
      if (rBulk) {
        var st = undoSnaps.get(pg);
        if (!st) { st = []; undoSnaps.set(pg, st); }
        st.push({ before: rBulk.before, after: rBulk.after, afterLen: rBulk.afterLen });
        pg.objects = rBulk.after;
        rst.pop();
        clearSelection();
        renderInk();
        if (keyLabel) showToast('취소했던 지우기를 다시 적용했습니다. (' + keyLabel + ')');
        return true;
      }
      if (!pg.redoStack.length) {
        if (keyLabel) showToast('다시 실행할 필기가 없습니다.');
        return false;
      }
      var restored = pg.redoStack.pop();
      pg.objects.push(restored);
      if (restored && restored.groupId) {
        while (pg.redoStack.length && pg.redoStack[pg.redoStack.length - 1].groupId === restored.groupId) {
          pg.objects.push(pg.redoStack.pop());
        }
      }
      clearSelection();
      renderInk();
      if (keyLabel) showToast('앞으로 다시 실행했습니다. (' + keyLabel + ')');
      return true;
    }

    if (el('sbUndoBtn')) el('sbUndoBtn').addEventListener('click', function () {
      performUndo('');
    });
    if (el('sbRedoBtn')) el('sbRedoBtn').addEventListener('click', function () {
      performRedo('');
    });
    // 현재 쪽에서 조건에 맞는 판서를 한꺼번에 지웁니다. 바로 되돌리기(Ctrl+Z)하면 모두 되살아납니다.
    function bulkClear(keep, doneMsg, emptyMsg) {
      var pg = curPage();
      var after = pg.objects.filter(keep);
      var removed = pg.objects.length - after.length;
      if (!removed) { showToast(emptyMsg); return; }
      pushUndoSnap(pg, pg.objects, after);
      pg.objects = after;
      pg.redoStack = [];
      clearSelection();
      renderInk();
      closeAllPopovers();
      showToast(doneMsg.replace('{n}', removed) + ' 되돌리기(Ctrl+Z)로 복구할 수 있습니다.');
    }
    if (el('sbClearInkBtn')) el('sbClearInkBtn').addEventListener('click', function () {
      bulkClear(function () { return false; }, '현재 쪽의 필기 {n}개를 모두 지웠습니다.', '지울 필기가 없습니다.');
    });
    if (el('sbClearPenBtn')) el('sbClearPenBtn').addEventListener('click', function () {
      bulkClear(function (o) { return o.type !== 'stroke' && o.type !== 'highlighter' && !o.isHighlighter; }, '펜 글씨 {n}개를 지웠습니다.', '지울 펜 글씨가 없습니다.');
    });
    ['sbClearStampsBtn', 'sbClearStampsBtn2'].forEach(function (id) {
      if (el(id)) el(id).addEventListener('click', function () {
        bulkClear(function (o) { return o.type !== 'stamp'; }, '번호 스탬프 {n}개를 지웠습니다. 다음 스탬프는 1번부터 찍힙니다.', '지울 스탬프가 없습니다.');
      });
    });

    if (el('sbZoomResetBtn')) el('sbZoomResetBtn').addEventListener('click', function () { resetZoom(); if (state.preZoom) restorePreZoom(); else { if (state.mode === 'zoom') setMode('draw'); if (state.tool === 'hand') setTool(state.lastDrawTool || 'pen'); } });
    if (el('sbZoomHandBtn')) el('sbZoomHandBtn').addEventListener('click', function () {
      if (state.mode === 'zoom') setMode('draw');
      setTool(state.tool === 'hand' ? (state.lastDrawTool || 'pen') : 'hand');
    });
    inlineTextEl.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); commitInlineText(); }
      else if (e.key === 'Escape') { e.stopPropagation(); inlineTextEl.style.display = 'none'; pendingTextPos = null; }
    });

    // 편집 모드 액션 바
    if (el('sbEditOcrBtn')) el('sbEditOcrBtn').addEventListener('click', runOcrOnSelection);
    if (el('sbQuickPinBtn')) el('sbQuickPinBtn').addEventListener('click', function () {
      commitInlineText();
      if (state.selection) {
        var selection = state.selection;
        clearSelection();
        createPinFromBoardRect(selection.x, selection.y, selection.w, selection.h);
      } else {
        setMode('pin');
        showToast('고정할 영역을 드래그하세요. 선택한 부분을 조각으로 게시합니다.');
      }
    });
    if (el('sbEditPinBtn')) el('sbEditPinBtn').addEventListener('click', function () {
      if (!state.selection) return;
      var s = state.selection;
      clearSelection();
      createPinFromBoardRect(s.x, s.y, s.w, s.h);
    });
    if (el('sbEditDelBtn')) el('sbEditDelBtn').addEventListener('click', function () {
      if (!state.selection) return;
      var ids = state.selection.ids;
      var pg = curPage();
      pg.objects = pg.objects.filter(function (o) { return ids.indexOf(o.id) < 0; });
      clearSelection();
      renderInk();
      showToast('선택한 판서 ' + ids.length + '개를 삭제했습니다.');
    });
    if (el('sbEditCancelBtn')) el('sbEditCancelBtn').addEventListener('click', clearSelection);

    // 페이지 & PPTX 액션 버튼 + 우측 하단 미니 네비게이터
    if (el('sbPrevPageBtn')) el('sbPrevPageBtn').addEventListener('click', function () { goPage(state.pageIndex - 1); });
    if (el('sbNextPageBtn')) el('sbNextPageBtn').addEventListener('click', function () { goPage(state.pageIndex + 1); });
    if (el('sbAddPageBtn')) el('sbAddPageBtn').addEventListener('click', function () {
      state.pages.push(createBlankPage(state.bgMode));
      goPage(state.pages.length - 1);
    });
    if (el('sbPptxPrevActBtn')) el('sbPptxPrevActBtn').addEventListener('click', function () { stepPptxAction(-1); });
    if (el('sbPptxNextActBtn')) el('sbPptxNextActBtn').addEventListener('click', function () { stepPptxAction(1); });
    if (el('sbPptxShowAllBtn')) el('sbPptxShowAllBtn').addEventListener('click', function () {
      state.pptxShowAll = !state.pptxShowAll;
      this.textContent = state.pptxShowAll ? '액션 모드' : '전체 보기';
      renderBackground();
    });
    if (el('sbMiniPrevBtn')) el('sbMiniPrevBtn').addEventListener('click', function () {
      if (curPage().kind === 'pptx') stepPptxAction(-1);
      else goPage(state.pageIndex - 1);
    });
    if (el('sbMiniNextBtn')) el('sbMiniNextBtn').addEventListener('click', function () {
      if (curPage().kind === 'pptx') stepPptxAction(1);
      else goPage(state.pageIndex + 1);
    });

    if (el('sbStampResetBtn')) el('sbStampResetBtn').addEventListener('click', function () {
      resetStampNumber(false);
    });
    if (el('sbClearPinsBtn')) el('sbClearPinsBtn').addEventListener('click', function () {
      pinLayer.innerHTML = '';
      showToast('모든 게시 조각을 닫았습니다.');
    });

    // 저장 / 단축키 전체창(좌우 슬라이드) / 사용설명서 2단계 뷰(큰 글씨 요약 <-> 전체화면 GIF 좌우 슬라이드 과정)
    // PNG · PDF 저장 / 펜 종류 / 도형·도구 / 보드 설정 팝오버 통합 제어
    function closeAllPopovers() {
      ['sbBackupMenu', 'sbSaveMenu', 'sbShareMenu', 'sbPenMenu', 'sbMoreToolsMenu', 'sbBoardMoreMenu', 'sbEraserMenu', 'sbLaserMenu'].forEach(function (id) {
        if (el(id)) el(id).classList.remove('show');
      });
    }
    function bindAnchoredPopover(btnId, menuId, alignRight) {
      var btn = el(btnId);
      var menu = el(menuId);
      if (!btn || !menu) return;
      btn.addEventListener('click', function (e) {
        e.stopPropagation();
        var willOpen = !menu.classList.contains('show');
        closeAllPopovers();
        if (willOpen) showAnchoredPopover(btn, menu, alignRight);
      });
      menu.addEventListener('click', function (e) { e.stopPropagation(); });
    }
    function showAnchoredPopover(btn, menu, alignRight) {
      var rect = btn.getBoundingClientRect();
      menu.style.top = Math.round(rect.bottom + 6) + 'px';
      if (alignRight) {
        menu.style.left = 'auto';
        menu.style.right = Math.max(10, Math.round(window.innerWidth - rect.right)) + 'px';
      } else {
        menu.style.right = 'auto';
        var maxLeft = Math.max(10, window.innerWidth - 330);
        menu.style.left = Math.min(maxLeft, Math.max(10, Math.round(rect.left))) + 'px';
      }
      menu.classList.add('show');
    }
    // 지우개 버튼: 처음 누르면 지우개, 지우개가 켜진 상태에서 한 번 더 누르거나 ▾를 누르면 전체 지우기 메뉴
    (function () {
      var btn = el('sbEraserToolBtn'), menu = el('sbEraserMenu');
      if (!btn || !menu) return;
      var wasEraser = false;
      btn.addEventListener('pointerdown', function () { wasEraser = state.tool === 'eraser'; }, true);
      btn.addEventListener('click', function (e) {
        e.stopPropagation();
        var open = menu.classList.contains('show');
        closeAllPopovers();
        if (!open && (wasEraser || e.target.closest('.sb-caret'))) showAnchoredPopover(btn, menu, false);
      });
      menu.addEventListener('click', function (e) { e.stopPropagation(); });
      menu.querySelectorAll('[data-sb-tool]').forEach(function (b) { b.addEventListener('click', closeAllPopovers); });
      [['sbEraseAllPenBtn', 'sbClearPenBtn'], ['sbEraseAllStampBtn', 'sbClearStampsBtn'], ['sbEraseAllInkBtn', 'sbClearInkBtn']].forEach(function (pair) {
        if (el(pair[0])) el(pair[0]).addEventListener('click', function () { closeAllPopovers(); if (el(pair[1])) el(pair[1]).click(); });
      });
    })();
    // 레이저 버튼: 처음 누르면 레이저, 레이저가 켜진 상태에서 한 번 더 누르거나 ▾를 누르면 유지 시간 메뉴
    (function () {
      var btn = document.querySelector('.sb-tool-laser-btn'), menu = el('sbLaserMenu');
      if (!btn || !menu) return;
      var wasLaser = false;
      btn.addEventListener('pointerdown', function () { wasLaser = state.tool === 'laser'; }, true);
      btn.addEventListener('click', function (e) {
        e.stopPropagation();
        var open = menu.classList.contains('show');
        closeAllPopovers();
        if (!open && (wasLaser || e.target.closest('.sb-caret'))) showAnchoredPopover(btn, menu, false);
      });
      menu.addEventListener('click', function (e) { e.stopPropagation(); });
      menu.querySelectorAll('[data-laser-hold]').forEach(function (b) {
        b.addEventListener('click', function () { setLaserHold(+b.getAttribute('data-laser-hold')); closeAllPopovers(); if (state.tool !== 'laser') setTool('laser'); });
      });
      if (el('sbLaserClearBtn')) el('sbLaserClearBtn').addEventListener('click', function () { clearLaserStrokes(); closeAllPopovers(); showToast('레이저 선을 모두 지웠습니다.'); });
      syncLaserHoldUi();
    })();
    bindAnchoredPopover('sbBackupMenuBtn', 'sbBackupMenu', false);
    bindAnchoredPopover('sbSaveMenuBtn', 'sbSaveMenu', true);
    bindAnchoredPopover('sbShareMenuBtn', 'sbShareMenu', true);
    if (el('sbMenuBackupAllBtn')) el('sbMenuBackupAllBtn').addEventListener('click', function () { closeAllPopovers(); exportAllVaultJson(); });
    if (el('sbMenuBackupCurBtn')) el('sbMenuBackupCurBtn').addEventListener('click', function () { closeAllPopovers(); exportSingleNoteJson(null); });
    if (el('sbMenuRestoreBtn')) el('sbMenuRestoreBtn').addEventListener('click', function () {
      closeAllPopovers();
      if (el('sbVaultImportJsonInput')) el('sbVaultImportJsonInput').click();
    });
    if (el('sbSaveJsonBackupBtn')) el('sbSaveJsonBackupBtn').addEventListener('click', function () { closeAllPopovers(); exportAllVaultJson(); });
    if (el('sbSharePngBtn')) el('sbSharePngBtn').addEventListener('click', function () { closeAllPopovers(); sharePng(); });
    if (el('sbCopyPngBtn')) el('sbCopyPngBtn').addEventListener('click', function () { closeAllPopovers(); copyPngToClipboard(false); });
    if (el('sbSharePngZipBtn')) el('sbSharePngZipBtn').addEventListener('click', function () { closeAllPopovers(); shareAllPngZip(); });
    if (el('sbSharePdfBtn')) el('sbSharePdfBtn').addEventListener('click', function () { closeAllPopovers(); sharePdf(); });
    if (el('sbShareReadyGo')) el('sbShareReadyGo').addEventListener('click', function () {
      el('sbShareReadyModal').classList.remove('show');
      if (pendingShare) { var ps = pendingShare; pendingShare = null; navigator.share({ files: [ps.file], title: ps.file.name }).catch(function (e) { if (!e || e.name !== 'AbortError') { downloadBlob(ps.file, ps.file.name); showToast('공유하지 못해 파일로 내려받았습니다.'); } }); }
    });
    if (el('sbShareReadyDl')) el('sbShareReadyDl').addEventListener('click', function () {
      el('sbShareReadyModal').classList.remove('show');
      if (pendingShare) { downloadBlob(pendingShare.file, pendingShare.file.name); pendingShare = null; }
    });
    bindAnchoredPopover('sbPenMenuBtn', 'sbPenMenu', false);
    bindAnchoredPopover('sbMoreToolsBtn', 'sbMoreToolsMenu', false);
    bindAnchoredPopover('sbBoardMoreBtn', 'sbBoardMoreMenu', true);
    document.addEventListener('click', closeAllPopovers);
    liveCanvas.addEventListener('pointerdown', closeAllPopovers);
    if (el('sbMoreToolsMenu')) {
      el('sbMoreToolsMenu').querySelectorAll('[data-sb-tool]').forEach(function (b) {
        b.addEventListener('click', closeAllPopovers);
      });
    }
    if (el('sbSavePngBtn')) el('sbSavePngBtn').addEventListener('click', function () { closeAllPopovers(); saveCurrentPng(); });
    if (el('sbSavePngZipBtn')) el('sbSavePngZipBtn').addEventListener('click', function () { closeAllPopovers(); saveAllPngZip(); });
    if (el('sbAutoStampToggleBtn')) el('sbAutoStampToggleBtn').addEventListener('click', toggleAutoStampShape);
    if (el('sbAutoStampMenuBtn')) el('sbAutoStampMenuBtn').addEventListener('click', toggleAutoStampShape);
    syncAutoStampUi();
    if (el('sbSavePdfBtn')) el('sbSavePdfBtn').addEventListener('click', function () { closeAllPopovers(); saveMergedPdf(); });
    if (el('sbShortcutBtn')) el('sbShortcutBtn').addEventListener('click', function () {
      goScSlide(0);
      el('sbShortcutModal').classList.add('show');
    });
    if (el('sbShortcutToggleFsBtn')) el('sbShortcutToggleFsBtn').addEventListener('click', function () {
      var card = el('sbShortcutModalCard');
      if (!card) return;
      var fs = card.classList.toggle('is-fullscreen');
      this.textContent = fs ? '창 크기로 보기' : '전체화면으로 보기';
    });

    // 단축키 안내 모달 내부 스크린샷 좌우 슬라이드 (1 / 2)
    var scSlideIdx = 0;
    function goScSlide(idx) {
      var panes = document.querySelectorAll('[data-sc-slide-pane]');
      if (!panes.length) return;
      scSlideIdx = (idx + panes.length) % panes.length;
      panes.forEach(function (p, i) {
        p.style.display = i === scSlideIdx ? 'flex' : 'none';
        p.classList.toggle('active', i === scSlideIdx);
      });
      document.querySelectorAll('[data-sc-slide]').forEach(function (t, i) {
        t.classList.toggle('active', i === scSlideIdx);
      });
      if (el('sbScSlideCounter')) el('sbScSlideCounter').textContent = (scSlideIdx + 1) + ' / ' + panes.length;
    }
    if (el('sbScPrevSlideBtn')) el('sbScPrevSlideBtn').addEventListener('click', function () { goScSlide(scSlideIdx - 1); });
    if (el('sbScNextSlideBtn')) el('sbScNextSlideBtn').addEventListener('click', function () { goScSlide(scSlideIdx + 1); });
    document.querySelectorAll('[data-sc-slide]').forEach(function (tb) {
      tb.addEventListener('click', function () { goScSlide(parseInt(tb.getAttribute('data-sc-slide'), 10) || 0); });
    });

    // 사용설명서 창 — 영상 매뉴얼(smart-board-guide.html)은 처음 열 때 불러옵니다
    function openHelpGuide() {
      var frame = el('sbHelpFrame');
      if (frame && !frame.getAttribute('src')) frame.setAttribute('src', frame.getAttribute('data-src'));
      if (el('sbHelpModal')) el('sbHelpModal').classList.add('show');
    }

    // 사용설명서 전체화면 GIF & 스크린샷 좌우 슬라이드 (1 / 6)
    var visSlideIdx = 0;
    var visShowAll = false;
    function goVisSlide(idx) {
      var panes = document.querySelectorAll('[data-vis-slide-pane]');
      if (!panes.length) return;
      visShowAll = false;
      if (el('sbVisShowAllToggleBtn')) el('sbVisShowAllToggleBtn').textContent = '연속 스크롤로 보기';
      visSlideIdx = (idx + panes.length) % panes.length;
      panes.forEach(function (p, i) {
        p.style.display = i === visSlideIdx ? 'flex' : 'none';
        p.classList.toggle('active', i === visSlideIdx);
      });
      document.querySelectorAll('[data-vis-slide]').forEach(function (t, i) {
        t.classList.toggle('active', i === visSlideIdx);
      });
      if (el('sbVisSlideCounter')) el('sbVisSlideCounter').textContent = (visSlideIdx + 1) + ' / ' + panes.length;
    }
    if (el('sbVisPrevSlideBtn')) el('sbVisPrevSlideBtn').addEventListener('click', function () { goVisSlide(visSlideIdx - 1); });
    if (el('sbVisNextSlideBtn')) el('sbVisNextSlideBtn').addEventListener('click', function () { goVisSlide(visSlideIdx + 1); });
    document.querySelectorAll('[data-vis-slide]').forEach(function (tb) {
      tb.addEventListener('click', function () { goVisSlide(parseInt(tb.getAttribute('data-vis-slide'), 10) || 0); });
    });
    if (el('sbVisShowAllToggleBtn')) el('sbVisShowAllToggleBtn').addEventListener('click', function () {
      var panes = document.querySelectorAll('[data-vis-slide-pane]');
      visShowAll = !visShowAll;
      if (visShowAll) {
        panes.forEach(function (p) { p.style.display = 'flex'; });
        this.textContent = '1장씩 슬라이드로 보기';
      } else {
        goVisSlide(visSlideIdx);
      }
    });

    function switchHelpVisualMode(showVisualGif) {
      var quickPane = el('sbGuideQuickTextPane');
      var visualPane = el('sbGuideFullVisualPane');
      var toggleBtn = el('sbToggleVisualGuideBtn');
      var card = el('sbHelpModalCard');
      if (!quickPane || !visualPane) return;
      quickPane.style.display = showVisualGif ? 'none' : 'block';
      visualPane.style.display = showVisualGif ? 'block' : 'none';
      if (showVisualGif) goVisSlide(0);
      if (card) card.classList.add('is-fullscreen');
      if (toggleBtn) {
        toggleBtn.textContent = showVisualGif ? '← 큰 글씨 핵심 요약으로 보기' : '전체화면으로 GIF & 상세 과정 보기';
      }
    }

    if (el('sbHelpBtn')) el('sbHelpBtn').addEventListener('click', openHelpGuide);
    if (el('sbToggleVisualGuideBtn')) el('sbToggleVisualGuideBtn').addEventListener('click', function () {
      var visualPane = el('sbGuideFullVisualPane');
      var isCurrentlyVisual = visualPane && visualPane.style.display !== 'none';
      switchHelpVisualMode(!isCurrentlyVisual);
    });
    if (el('sbEnterFullGifViewBtn')) el('sbEnterFullGifViewBtn').addEventListener('click', function () {
      switchHelpVisualMode(true);
    });
    if (el('sbBackToQuickTextBtn')) el('sbBackToQuickTextBtn').addEventListener('click', function () {
      switchHelpVisualMode(false);
    });

    if (el('sbFullscreenBtn')) el('sbFullscreenBtn').addEventListener('click', function () {
      if (!document.fullscreenElement) {
        if (document.documentElement.requestFullscreen) document.documentElement.requestFullscreen().catch(function () {});
      } else if (document.exitFullscreen) {
        document.exitFullscreen().catch(function () {});
      }
      setTimeout(function () {
        resizeBoardToPage();
        renderBackground();
        renderInk();
      }, 150);
    });

    // OCR 설정 & 결과 모달
    var ocrCfg = loadOcrCfg();
    if (el('sbOcrEngineSelect')) {
      el('sbOcrEngineSelect').value = ocrCfg.engine || 'gemini';
      var syncOcrFields = function () {
        var v = el('sbOcrEngineSelect').value;
        el('sbOcrGeminiFields').style.display = v === 'gemini' ? 'block' : 'none';
        el('sbOcrSpaceFields').style.display = v === 'ocrspace' ? 'block' : 'none';
      };
      syncOcrFields();
      el('sbOcrEngineSelect').addEventListener('change', syncOcrFields);
    }
    if (el('sbGeminiKeyInput')) el('sbGeminiKeyInput').value = ocrCfg.geminiKey || '';
    if (el('sbOcrSpaceKeyInput')) el('sbOcrSpaceKeyInput').value = ocrCfg.ocrSpaceKey || '';
    if (el('sbOcrCfgBtn')) el('sbOcrCfgBtn').addEventListener('click', function () { el('sbOcrCfgModal').classList.add('show'); });
    if (el('sbOcrSaveCfgBtn')) el('sbOcrSaveCfgBtn').addEventListener('click', function () {
      saveOcrCfg({
        engine: el('sbOcrEngineSelect').value,
        geminiKey: (el('sbGeminiKeyInput').value || '').trim(),
        ocrSpaceKey: (el('sbOcrSpaceKeyInput').value || '').trim()
      });
      el('sbOcrCfgModal').classList.remove('show');
      showToast('무료 OCR 설정을 저장했습니다.');
    });

    if (el('sbOcrReplaceBtn')) el('sbOcrReplaceBtn').addEventListener('click', function () {
      var txt = (el('sbOcrResultText').value || '').trim();
      if (!txt || !state.ocrPendingBox) return;
      var pg = curPage();
      var ids = state.ocrPendingBox.ids || [];
      pg.objects = pg.objects.filter(function (o) { return ids.indexOf(o.id) < 0; });
      pg.objects.push({
        id: uid(),
        type: 'text',
        x: state.ocrPendingBox.x + 6,
        y: state.ocrPendingBox.y + 30,
        text: txt,
        color: pg.kind === 'blank' && pg.bgMode === 'dark' ? '#f8fafc' : '#0f172a',
        fontSize: 26
      });
      pg.redoStack = [];
      clearSelection();
      renderInk();
      el('sbOcrResultModal').classList.remove('show');
      showToast('손글씨 판서를 깔끔한 텍스트로 교체했습니다.');
    });
    if (el('sbOcrCopyBtn')) el('sbOcrCopyBtn').addEventListener('click', function () {
      var txt = (el('sbOcrResultText').value || '').trim();
      if (txt && navigator.clipboard) navigator.clipboard.writeText(txt).then(function () { showToast('변환 텍스트를 복사했습니다.'); });
    });
    if (el('sbOcrToMemoBtn')) el('sbOcrToMemoBtn').addEventListener('click', function () {
      var txt = (el('sbOcrResultText').value || '').trim();
      if (!txt) return;
      try {
        var list = JSON.parse(localStorage.getItem(MEMO_NOTES_KEY) || '[]');
        if (!Array.isArray(list)) list = [];
        list.push({
          id: 'm_' + Date.now(),
          room: 'local',
          group: '공통',
          author: '판서 OCR',
          color: '#e0f2fe',
          text: txt,
          likes: 0,
          createdAt: new Date().toISOString().replace('T', ' ').slice(0, 19)
        });
        localStorage.setItem(MEMO_NOTES_KEY, JSON.stringify(list));
        showToast('변환된 글씨를 T-ong 수업 메모보드로 보냈습니다!');
      } catch (e) { showToast('메모보드 전송에 실패했습니다.'); }
    });

    document.querySelectorAll('[data-close]').forEach(function (cb) {
      cb.addEventListener('click', function () {
        var m = el(cb.getAttribute('data-close'));
        if (m) m.classList.remove('show');
      });
    });
    // 부분 화면 모달에서 닫기 버튼 없이 빈 배경을 눌러도 바로 닫기
    document.querySelectorAll('.sb-modal').forEach(function (m) {
      m.addEventListener('click', function (ev) {
        if (ev.target === m) m.classList.remove('show');
      });
    });

    // ===================================================================
    // 단축키 내 마음대로 설정 (위젯처럼 카드에서 키를 바꾸거나 끄기 · 초기화)
    // ===================================================================
    var SC_ACTIONS = [
      { g: '필기 도구', k: 'p', n: '펜', i: '✏️' },
      { g: '필기 도구', k: 'h', n: '형광펜', i: '🖍️' },
      { g: '필기 도구', k: 'e', n: '지우개', i: '🧽', sh: 'Shift: 펜 글씨 전체 지우기' },
      { g: '필기 도구', k: 'd', n: '레이저', i: '🔴', sh: 'Shift: 레이저 유지 시간 바꾸기' },
      { g: '필기 도구', k: 'v', n: '선택·옮기기', i: '🖱️' },
      { g: '필기 도구', k: 'm', n: '화면 이동(손)', i: '✋' },
      { g: '필기 도구', k: 't', n: '텍스트', i: '🔤' },
      { g: '도형', k: 'l', n: '직선', i: '📏', sh: 'Shift: 화살표 → 양쪽화살표' },
      { g: '도형', k: 'r', n: '사각형', i: '▭', sh: 'Shift: 채움 사각형' },
      { g: '도형', k: 'c', n: '원', i: '◯' },
      { g: '번호 스탬프', k: 's', n: '원 스탬프 ①', i: '①', sh: '두 번: 번호 초기화 · Shift: 스탬프 전체 지우기' },
      { g: '번호 스탬프', k: 'q', n: '사각 스탬프 [1]', i: '▣', sh: '두 번: 번호 초기화' },
      { g: '번호 스탬프', k: '0', n: '스탬프 번호 초기화', i: '↺' },
      { g: '번호 스탬프', k: 'a', n: '자동① 켜기/끄기', i: '🅰️' },
      { g: '실행', k: 'z', n: '실행 취소', i: '↶' },
      { g: '실행', k: 'x', n: '다시 실행', i: '↷' },
      { g: '실행', k: 'f', n: '전체화면', i: '⛶' },
      { g: '배경', k: 'k', n: '칠판 배경', i: '⬛' },
      { g: '배경', k: 'w', n: '백지 배경', i: '⬜' },
      { g: '펜 색상', k: '1', n: '빨강', i: '🟥' },
      { g: '펜 색상', k: '2', n: '초록', i: '🟩' },
      { g: '펜 색상', k: '3', n: '파랑', i: '🟦' },
      { g: '펜 색상', k: '4', n: '노랑', i: '🟨' },
      { g: '펜 색상', k: '5', n: '흰색', i: '⬜' },
      { g: '펜 색상', k: '6', n: '검정', i: '⬛' }
    ];
    var SC_STORE = 'sbShortcutMap_v1';
    var scMap = {};   // 기본키 → 사용자 키('' = 끔). 없으면 기본키 그대로
    try { scMap = JSON.parse(localStorage.getItem(SC_STORE) || '{}') || {}; } catch (e) { scMap = {}; }
    function scKeyOf(defKey) { return Object.prototype.hasOwnProperty.call(scMap, defKey) ? scMap[defKey] : defKey; }
    function remapShortcutKey(pressed) {
      if (!pressed || pressed.length !== 1) return pressed;
      for (var i = 0; i < SC_ACTIONS.length; i++) { if (scKeyOf(SC_ACTIONS[i].k) === pressed) return SC_ACTIONS[i].k; }
      // 다른 키로 옮겼거나 끈 기능의 원래 키는 무시합니다.
      for (var j = 0; j < SC_ACTIONS.length; j++) { if (SC_ACTIONS[j].k === pressed) return ''; }
      return pressed;
    }
    function scLabel(key) { return key ? key.toUpperCase() : '끔'; }
    scDisplayKey = function (defKey) { var c = scKeyOf(defKey); return c ? c.toUpperCase() : ''; };
    function saveScMap() {
      // 기본값과 같은 항목은 지워서 저장을 가볍게 합니다.
      Object.keys(scMap).forEach(function (k) { if (scMap[k] === k) delete scMap[k]; });
      try { localStorage.setItem(SC_STORE, JSON.stringify(scMap)); } catch (e) {}
      applyScToButtons();
    }
    // 도구 버튼 옆 작은 글자(kbd)도 바꾼 키로 보이게 합니다.
    function applyScToButtons() {
      document.querySelectorAll('.sb-top-group kbd, .sb-tool-popover kbd').forEach(function (kb) {
        var t = (kb.textContent || '').trim();
        if (!kb.hasAttribute('data-sc-def')) {
          var m = /^(⇧)?([A-Z0-9])(²)?$/.exec(t);
          if (!m) return;
          kb.setAttribute('data-sc-def', m[2].toLowerCase());
          kb.setAttribute('data-sc-pre', m[1] || ''); kb.setAttribute('data-sc-post', m[3] || '');
        }
        var d = kb.getAttribute('data-sc-def');
        var isAction = SC_ACTIONS.some(function (a) { return a.k === d; });
        if (!isAction) return;
        var cur = scKeyOf(d);
        if (kb.id === 'sbMoreToolsKey' && !t) return;   // 도형 버튼: 도형 도구가 선택되지 않았으면 숨김 유지
        kb.textContent = cur ? (kb.getAttribute('data-sc-pre') || '') + cur.toUpperCase() + (kb.getAttribute('data-sc-post') || '') : '';
        kb.hidden = !cur;
      });
    }
    var scCapture = null;  // 키 입력을 기다리는 기본키
    function renderScGrid() {
      var box = el('sbScGrid'); if (!box) return;
      var groups = [];
      SC_ACTIONS.forEach(function (a) { if (groups.indexOf(a.g) < 0) groups.push(a.g); });
      box.innerHTML = groups.map(function (g) {
        return '<div class="sb-sc-group"><div class="sb-sc-gtitle">' + g + '</div><div class="sb-sc-cards">' +
          SC_ACTIONS.filter(function (a) { return a.g === g; }).map(function (a) {
            var cur = scKeyOf(a.k), changed = cur !== a.k, off = !cur, wait = scCapture === a.k;
            return '<div class="sb-sc-card' + (off ? ' off' : '') + (changed ? ' changed' : '') + (wait ? ' wait' : '') + '" data-sc="' + a.k + '">' +
              '<span class="sb-sc-ico">' + a.i + '</span>' +
              '<span class="sb-sc-name"><b>' + a.n + '</b>' + (a.sh ? '<small>' + a.sh + '</small>' : '') + (changed ? '<small class="sb-sc-def">기본 ' + a.k.toUpperCase() + '</small>' : '') + '</span>' +
              '<button type="button" class="sb-sc-key" data-sc-set="' + a.k + '" title="눌러서 새 키 정하기">' + (wait ? '키를 누르세요…' : scLabel(cur)) + '</button>' +
              '<label class="sb-sc-sw" title="이 단축키 켜기/끄기"><input type="checkbox" data-sc-on="' + a.k + '"' + (off ? '' : ' checked') + '><i></i></label>' +
            '</div>';
          }).join('') + '</div></div>';
      }).join('');
      box.querySelectorAll('[data-sc-set]').forEach(function (b) {
        b.onclick = function () { scCapture = scCapture === b.getAttribute('data-sc-set') ? null : b.getAttribute('data-sc-set'); renderScGrid(); };
      });
      box.querySelectorAll('[data-sc-on]').forEach(function (c) {
        c.onchange = function () {
          var d = c.getAttribute('data-sc-on');
          if (c.checked) {
            // 다시 켤 때 기본키가 다른 기능에 쓰이고 있으면 키를 새로 고르게 합니다.
            var taken = SC_ACTIONS.some(function (a) { return a.k !== d && scKeyOf(a.k) === d; });
            if (taken) { scMap[d] = ''; scCapture = d; showToast('기본 키 ' + d.toUpperCase() + '는 다른 기능에 쓰이고 있어요. 새 키를 눌러 주세요.'); }
            else delete scMap[d];
          } else { scMap[d] = ''; if (scCapture === d) scCapture = null; }
          saveScMap(); renderScGrid();
        };
      });
    }
    function setScKey(defKey, newKey) {
      var msg = '';
      SC_ACTIONS.forEach(function (a) {
        if (a.k !== defKey && scKeyOf(a.k) === newKey) {
          // 이미 쓰는 키면 서로 맞바꿉니다.
          scMap[a.k] = scKeyOf(defKey);
          msg = ' · [' + a.n + ']는 ' + scLabel(scMap[a.k]) + '로 바뀌었습니다';
        }
      });
      scMap[defKey] = newKey;
      saveScMap();
      var act = SC_ACTIONS.filter(function (a) { return a.k === defKey; })[0];
      showToast('[' + (act ? act.n : '') + '] 단축키를 ' + newKey.toUpperCase() + '로 바꿨습니다' + msg + '.');
    }
    document.addEventListener('keydown', function (e) {
      if (!scCapture) return;
      if (!el('sbScModal') || !el('sbScModal').classList.contains('show')) { scCapture = null; return; }
      e.preventDefault(); e.stopImmediatePropagation();
      if (e.key === 'Escape') { scCapture = null; renderScGrid(); return; }
      var key = (e.key || '').toLowerCase();
      if (e.ctrlKey || e.altKey || e.metaKey || !/^[a-z0-9]$/.test(key)) { showToast('영문자(A~Z) 또는 숫자(0~9) 한 개를 눌러 주세요. (Esc: 취소)'); return; }
      var d = scCapture; scCapture = null;
      setScKey(d, key);
      renderScGrid();
    }, true);
    function openScModal() { scCapture = null; renderScGrid(); el('sbScModal').classList.add('show'); }
    if (el('sbScOpenBtn')) el('sbScOpenBtn').addEventListener('click', function () { closeAllPopovers(); openScModal(); });
    if (el('sbScOpenBtn2')) el('sbScOpenBtn2').addEventListener('click', function () { el('sbShortcutModal').classList.remove('show'); openScModal(); });
    if (el('sbScResetBtn')) el('sbScResetBtn').addEventListener('click', function () {
      if (!confirm('모든 단축키를 처음(기본값)으로 되돌릴까요?')) return;
      scMap = {}; scCapture = null; saveScMap(); renderScGrid(); showToast('단축키를 기본값으로 초기화했습니다.');
    });
    applyScToButtons();

    // 전역 단축키 (스탬프 동일키 2연타 SS/QQ 감지 포함)
    var lastStampKeyTap = { key: '', time: 0 };
    var STAMP_DOUBLE_TAP_MS = 450;

    document.addEventListener('keydown', function (e) {
      var inInput = document.activeElement && (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'TEXTAREA' || document.activeElement.tagName === 'SELECT');
      if (inInput) return;
      if (state.webActive && state.webInteract && document.activeElement === webFrame) return;

      var rawLower = (e.key || '').toLowerCase();
      // 사용자가 바꾼 단축키를 원래(기본) 키로 바꿔 아래 처리 코드가 그대로 동작하게 합니다.
      if (!e.ctrlKey && !e.altKey && !e.metaKey) rawLower = remapShortcutKey(rawLower);
      if (rawLower !== 's' && rawLower !== 'q') {
        lastStampKeyTap.key = '';
      }

      if (e.key === 'Tab') { e.preventDefault(); togglePalette(); return; }
      if (e.key === '?' || e.key === 'F1') { e.preventDefault(); el('sbShortcutModal').classList.toggle('show'); return; }
      // 백스페이스(Backspace): 선택한 영역이 있으면 삭제, 없으면 판서를 저장하고 홈으로 / Delete: 선택 영역 삭제
      if (e.key === 'Backspace' || e.key === 'Delete') {
        e.preventDefault();
        if ((state.mode === 'edit' || state.tool === 'select') && state.selection && el('sbEditDelBtn')) {
          el('sbEditDelBtn').click();
        } else if (e.key === 'Backspace' && !e.repeat) {
          goHomeSaved();
        }
        return;
      }
      // 모드 전환: Ctrl+` 게시 / Ctrl+1 확대 / Ctrl+2 판서 / Ctrl+3 편집 (예전 Alt 조합도 그대로 동작)
      var modKey = (e.ctrlKey || e.metaKey || e.altKey) && !e.shiftKey;
      if (modKey && (e.code === 'Backquote' || e.key === '`')) { e.preventDefault(); setMode(state.mode === 'pin' ? 'draw' : 'pin'); return; }
      if (modKey && (e.key === '1' || e.code === 'Digit1')) { e.preventDefault(); toggleZoom(); return; }
      if (modKey && (e.key === '2' || e.code === 'Digit2')) { e.preventDefault(); setMode('draw'); return; }
      if (modKey && (e.key === '3' || e.code === 'Digit3')) { e.preventDefault(); setMode(state.mode === 'edit' ? 'draw' : 'edit'); return; }
      if ((e.ctrlKey || e.metaKey) && e.shiftKey && rawLower === 's') { e.preventDefault(); serializeCurrentNote(false); return; }
      if ((e.ctrlKey || e.metaKey) && rawLower === 's') { e.preventDefault(); if (el('sbClearStampsBtn')) el('sbClearStampsBtn').click(); return; }
      if ((e.ctrlKey || e.metaKey) && !e.shiftKey && rawLower === 'p') { e.preventDefault(); saveCurrentPng(); return; }
      if ((e.ctrlKey || e.metaKey) && e.shiftKey && rawLower === 'z') { e.preventDefault(); performRedo('Ctrl+Shift+Z'); return; }
      if ((e.ctrlKey || e.metaKey) && rawLower === 'z') { e.preventDefault(); performUndo('Ctrl+Z'); return; }
      if ((e.ctrlKey || e.metaKey) && rawLower === 'y') { e.preventDefault(); performRedo('Ctrl+Y'); return; }
      if ((e.ctrlKey || e.metaKey) && rawLower === 'e') { e.preventDefault(); el('sbClearInkBtn').click(); return; }

      // ESC: 열린 창·메뉴가 있으면 먼저 닫고, 없으면 직전 필기를 취소합니다(실행 취소).
      if (e.key === 'Escape') {
        var openModals = document.querySelectorAll('.sb-modal.show');
        var openPops = ['sbBackupMenu', 'sbSaveMenu', 'sbShareMenu', 'sbPenMenu', 'sbMoreToolsMenu', 'sbBoardMoreMenu', 'sbEraserMenu', 'sbLaserMenu'].filter(function (id) { return el(id) && el(id).classList.contains('show'); });
        if (openModals.length || openPops.length) {
          openModals.forEach(function (m) { m.classList.remove('show'); });
          closeAllPopovers();
          return;
        }
        if (state.selection) { clearSelection(); return; }
        if (state.tool === 'laser' && state.laserStrokes && state.laserStrokes.some(function (x) { return !x.active; })) { clearLaserStrokes(); return; }
        if (state.webActive && state.webInteract) { setWebInteractMode(false); return; }
        e.preventDefault();
        performUndo('ESC');
        return;
      }

      if (!e.ctrlKey && !e.altKey && !e.metaKey) {
        var k = rawLower;
        var now = Date.now();
        // 도형·도구·취소 단축키: 뒤로취소(Z / ,), 앞으로다시실행(X / .), 레이저(D), 선택(V), 지우개(E), 스탬프(S/Q, 초기화 SS/QQ/0), 자동스탬프(A)
        if (e.shiftKey && k === 'e') { if (el('sbClearPenBtn')) el('sbClearPenBtn').click(); }
        else if (e.shiftKey && k === 's') { if (el('sbClearStampsBtn')) el('sbClearStampsBtn').click(); }
        else if (k === 'z' || k === ',' || e.key === '<') {
          e.preventDefault();
          performUndo(k === 'z' ? 'Z' : '<');
        }
        else if (k === 'x' || k === '.' || e.key === '>') {
          e.preventDefault();
          performRedo(k === 'x' ? 'X' : '>');
        }
        else if (e.shiftKey && k === 'd') { if (!e.repeat) cycleLaserHold(); }
        else if (k === 'd') setTool(state.tool === 'laser' ? (state.lastDrawTool || 'pen') : 'laser');
        else if (k === 'v') setTool('select');
        else if (k === 'm') {
          setTool(state.tool === 'hand' ? (state.lastDrawTool || 'pen') : 'hand');
          if (state.tool === 'hand') showToast(state.zoomScale > 1.02 ? '이동(M): 끌어서 확대한 화면을 옮기세요. 다시 M을 누르면 펜으로 돌아갑니다.' : '이동(M): 확대(Ctrl+1)한 뒤 끌어서 화면을 옮길 수 있습니다.');
        }
        else if (k === 'e') setTool(state.tool === 'eraser' ? (state.lastDrawTool || 'pen') : 'eraser');
        else if (k === 's') {
          if (e.repeat) return;
          if (lastStampKeyTap.key === 's' && (now - lastStampKeyTap.time) <= STAMP_DOUBLE_TAP_MS) {
            lastStampKeyTap = { key: '', time: 0 };
            if (state.tool !== 'stamp') setTool('stamp');
            resetStampNumber(false, '원형 스탬프 번호를 ①번부터 다시 시작합니다. (SS · 0)');
          } else {
            lastStampKeyTap = { key: 's', time: now };
            setTool(state.tool === 'stamp' ? (state.lastDrawTool || 'pen') : 'stamp');
          }
        }
        else if (k === 'q') {
          if (e.repeat) return;
          if (lastStampKeyTap.key === 'q' && (now - lastStampKeyTap.time) <= STAMP_DOUBLE_TAP_MS) {
            lastStampKeyTap = { key: '', time: 0 };
            if (state.tool !== 'sqStamp') setTool('sqStamp');
            resetStampNumber(false, '사각 스탬프 번호를 [1]번부터 다시 시작합니다. (QQ · 0)');
          } else {
            lastStampKeyTap = { key: 'q', time: now };
            setTool(state.tool === 'sqStamp' ? (state.lastDrawTool || 'pen') : 'sqStamp');
          }
        }
        else if (k === '0') {
          if (!e.repeat) resetStampNumber(false, '스탬프 번호를 1번부터 다시 시작합니다. (단축키 0 · SS)');
        }
        else if (k === 'a') {
          if (!e.repeat) toggleAutoStampShape();
        }
        else if (k === 't') setTool(state.tool === 'text' ? (state.lastDrawTool || 'pen') : 'text');
        else if (k === 'r') setTool(e.shiftKey ? 'fillRect' : 'rect');
        else if (k === 'c') setTool('circle');
        else if (e.shiftKey && k === 'l') setTool(state.tool === 'arrow' ? 'arrow2' : 'arrow');
        else if (k === 'l') setTool('line');
        else if (k === 'p') setTool('pen');
        else if (k === 'h') setTool('highlighter');
        else if (k === 'f') { if (el('sbFullscreenBtn')) el('sbFullscreenBtn').click(); }
        // 숫자 1~6: 펜 색상 빠른 변경
        else if (k === '1') setColor('#ef4444');
        else if (k === '2') setColor('#22c55e');
        else if (k === '3') setColor('#3b82f6');
        else if (k === '4') setColor('#facc15');
        else if (k === '5') setColor('#f8fafc');
        else if (k === '6') setColor('#0f172a');
        else if (k === 'k') setBlankBgMode('dark');
        else if (k === 'w') setBlankBgMode('light');
        // 갤럭시 S펜 블루투스 에어액션(Air Action) → [지우개] <-> [펜] 양방향 스위치 토글
        else if (state.spenEraserOnBtn && (e.key === 'PageDown' || e.key === 'PageUp' || e.key === 'MediaPlayPause' || /^f(1[3-9]|2[0-4])$/i.test(e.key))) {
          e.preventDefault();
          activateSpenEraserButton(true, 'S펜 에어액션 스위치');
        }
        // 방향키(← / → · ↑ / ↓) · Space · PageUp/PageDown: 슬라이드·페이지 넘기기 및 PPTX 액션 진행
        else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp' || (!state.spenEraserOnBtn && e.key === 'PageUp')) {
          if (e.key === 'ArrowLeft') {
            if (el('sbHelpModal') && el('sbHelpModal').classList.contains('show') && el('sbGuideFullVisualPane') && el('sbGuideFullVisualPane').style.display !== 'none') {
              e.preventDefault();
              goVisSlide(visSlideIdx - 1);
              return;
            }
            if (el('sbShortcutModal') && el('sbShortcutModal').classList.contains('show')) {
              e.preventDefault();
              goScSlide(scSlideIdx - 1);
              return;
            }
          }
          e.preventDefault();
          if (curPage().kind === 'pptx') stepPptxAction(-1);
          else goPage(state.pageIndex - 1);
        }
        else if (e.key === ' ' || e.key === 'ArrowRight' || e.key === 'ArrowDown' || (!state.spenEraserOnBtn && e.key === 'PageDown')) {
          if (el('sbHelpModal') && el('sbHelpModal').classList.contains('show') && el('sbGuideFullVisualPane') && el('sbGuideFullVisualPane').style.display !== 'none') {
            e.preventDefault();
            goVisSlide(visSlideIdx + 1);
            return;
          }
          if (el('sbShortcutModal') && el('sbShortcutModal').classList.contains('show')) {
            e.preventDefault();
            goScSlide(scSlideIdx + 1);
            return;
          }
          if (curPage().kind === 'pptx') { e.preventDefault(); stepPptxAction(1); }
          else if (e.key !== ' ') { e.preventDefault(); goPage(state.pageIndex + 1); }
        }
      }
    });

    // 단독 안드로이드 PWA / Vercel 링크 실행을 위한 서비스 워커 등록
    var h = location.hostname;
    var isLocal = h === 'localhost' || h === '127.0.0.1' || /^192\.168\./.test(h);
    if (!isLocal && 'serviceWorker' in navigator) {
      navigator.serviceWorker.register('/sw.js').catch(function () {});
    }
  }

  init();
  // 매뉴얼 영상·기능 점검용 읽기 전용 연결 (?sbdebug=1 일 때만)
  if (/[?&]sbdebug=1/.test(location.search)) window.__sbDebug = { state: state, pointer: pointer };
})();
