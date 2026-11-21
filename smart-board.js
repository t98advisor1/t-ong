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
    color: '#ef4444',
    size: 5,
    bgMode: 'dark',        // 'dark' | 'light' | 'grid'
    palmReject: false,
    spenEraserOnBtn: true, // 갤럭시 S펜 옆 버튼 누르기 / 에어액션 시 [지우개] <-> [펜] 양방향 스위치
    lastSpenSwitchTime: 0,
    lastDrawTool: 'pen',
    autoHide: true,        // 상단 팔레트 자동 숨김 (판서 시작 시 자동으로 위로 숨김)
    fillScreen: true,      // 화면 송출 시 여백 없는 100% 꽉 찬 화면 모드
    stampCounter: 1,
    zoomScale: 1,
    panX: 0,
    panY: 0,
    baseScale: 1,
    pages: [],
    pageIndex: 0,
    pdfDoc: null,
    pdfCache: {},
    selection: null,
    laserPoints: [],
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
    var rw = Math.max(100, vp.w);
    var rh = Math.max(100, vp.h);
    var w = pg.width || 1280;
    var h = pg.height || 720;
    var sx = rw / w;
    var sy = rh / h;
    // 1:1 정방형 픽셀 비율 유지 (원형 스탬프가 절대 찌그러지지 않으면서 화면을 꽉 채움)
    state.baseScale = state.fillScreen ? Math.max(sx, sy) : Math.min(sx, sy);
    if (pg.kind === 'pdf' || pg.kind === 'pptx' || pg.kind === 'image') {
      // 문서/슬라이드는 우측·하단 잘림 없이 100% 화면 안에 들어오도록 비율 맞춤
      state.baseScale = Math.min(sx, sy);
    }
    // 1280px 고정 DOM 박스로 인한 태블릿(960px 등) 레이아웃 뷰포트 확장·우측 1/4 잘림을 원천 차단하기 위해
    // boardFrame 의 CSS 크기를 실제 화면 맞춤 픽셀(fitW x fitH)로 직접 지정하고 내부 캔버스(1280x720)만 GPU 스케일링
    var fitW = Math.max(100, Math.round(w * state.baseScale));
    var fitH = Math.max(100, Math.round(h * state.baseScale));
    boardFrame.style.width = fitW + 'px';
    boardFrame.style.height = fitH + 'px';
    applyTransform();
  }

  function applyTransform() {
    boardFrame.style.transform = 'translate(' + state.panX + 'px, ' + state.panY + 'px) scale(' + state.zoomScale + ')';
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
      return { x: lx - 6, y: ly - 6, w: Math.abs(o.x2 - o.x1) + 12, h: Math.abs(o.y2 - o.y1) + 12 };
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

  function drawInkObject(ctx, o) {
    ctx.save();
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    if (o.type === 'stroke' || o.type === 'highlighter') {
      var pts = o.points || [];
      if (!pts.length) { ctx.restore(); return; }
      ctx.strokeStyle = o.color || '#ef4444';
      ctx.lineWidth = o.type === 'highlighter' ? Math.max(14, (o.size || 5) * 3.2) : (o.size || 5);
      if (o.type === 'highlighter') ctx.globalAlpha = 0.34;
      ctx.beginPath();
      ctx.moveTo(pts[0].x, pts[0].y);
      if (pts.length === 1) {
        ctx.lineTo(pts[0].x + 0.1, pts[0].y + 0.1);
      } else {
        for (var i = 1; i < pts.length; i++) {
          var midX = (pts[i - 1].x + pts[i].x) / 2;
          var midY = (pts[i - 1].y + pts[i].y) / 2;
          ctx.quadraticCurveTo(pts[i - 1].x, pts[i - 1].y, midX, midY);
        }
        ctx.lineTo(pts[pts.length - 1].x, pts[pts.length - 1].y);
      }
      ctx.stroke();
    } else if (o.type === 'line') {
      ctx.strokeStyle = o.color || '#ef4444';
      ctx.lineWidth = o.size || 5;
      ctx.beginPath();
      ctx.moveTo(o.x1, o.y1);
      ctx.lineTo(o.x2, o.y2);
      ctx.stroke();
    } else if (o.type === 'rect') {
      ctx.strokeStyle = o.color || '#ef4444';
      ctx.lineWidth = o.size || 5;
      ctx.strokeRect(o.x, o.y, o.w, o.h);
    } else if (o.type === 'circle') {
      ctx.strokeStyle = o.color || '#ef4444';
      ctx.lineWidth = o.size || 5;
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
  // 레이저 포인터
  // ===================================================================
  function renderLaserLoop() {
    var now = Date.now();
    state.laserPoints = state.laserPoints.filter(function (p) { return now - p.t < 520; });
    var pg = curPage();
    liveCtx.clearRect(0, 0, pg.width, pg.height);
    if (state.laserPoints.length > 1) {
      for (var i = 1; i < state.laserPoints.length; i++) {
        var p0 = state.laserPoints[i - 1], p1 = state.laserPoints[i];
        var alpha = Math.max(0, 1 - (now - p1.t) / 520);
        liveCtx.save();
        liveCtx.strokeStyle = 'rgba(239, 68, 68, ' + alpha + ')';
        liveCtx.shadowColor = '#ef4444';
        liveCtx.shadowBlur = 12;
        liveCtx.lineWidth = 5 * alpha + 2;
        liveCtx.lineCap = 'round';
        liveCtx.beginPath();
        liveCtx.moveTo(p0.x, p0.y);
        liveCtx.lineTo(p1.x, p1.y);
        liveCtx.stroke();
        liveCtx.restore();
      }
      var head = state.laserPoints[state.laserPoints.length - 1];
      liveCtx.save();
      liveCtx.fillStyle = '#ffffff';
      liveCtx.shadowColor = '#ef4444';
      liveCtx.shadowBlur = 16;
      liveCtx.beginPath();
      liveCtx.arc(head.x, head.y, 6, 0, Math.PI * 2);
      liveCtx.fill();
      liveCtx.restore();
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
    var comp = createCompositeCanvas();
    var sx = Math.max(0, Math.round(rx));
    var sy = Math.max(0, Math.round(ry));
    var sw = Math.min(comp.width - sx, Math.round(rw));
    var sh = Math.min(comp.height - sy, Math.round(rh));
    if (sw < 16 || sh < 16) return;

    var cut = document.createElement('canvas');
    cut.width = sw;
    cut.height = sh;
    cut.getContext('2d').drawImage(comp, sx, sy, sw, sh, 0, 0, sw, sh);
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

    showToast('영역을 게시 조각으로 고정했습니다. (페이지를 넘겨도 유지됩니다)');
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

  function openWebBrowser(urlStr) {
    var raw = String(urlStr || (el('sbWebUrlInput') && el('sbWebUrlInput').value) || '').trim();
    if (!raw) raw = 'https://ko.m.wikipedia.org/wiki/%EC%A1%B0%EA%B1%B4%EB%B6%80_%ED%99%95%EB%A5%A0';
    var ytEmbed = convertYoutubeToEmbedUrl(raw);
    var targetUrl = ytEmbed || raw;
    if (!ytEmbed && !/^https?:\/\//i.test(raw)) {
      if (/^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}(\/.*)?$/.test(raw)) {
        targetUrl = 'https://' + raw;
      } else {
        targetUrl = 'https://ko.m.wikipedia.org/w/index.php?search=' + encodeURIComponent(raw);
      }
    }
    if (el('sbWebUrlInput')) el('sbWebUrlInput').value = targetUrl;
    state.webActive = true;
    if (webBar) webBar.style.display = 'flex';
    if (webFrame) {
      webFrame.style.display = 'block';
      webFrame.src = targetUrl;
    }
    // 유튜브는 바로 재생 버튼을 누를 수 있도록 웹 조작 모드로, 일반 웹은 현재 도구 유지
    if (ytEmbed) {
      setWebInteractMode(true);
      showToast('유튜브 영상을 보드 안에서 열었습니다. 재생 후 하단 [① 원형스탬프] 또는 [펜 필기]로 영상 위에 필기하세요.');
    } else {
      setWebInteractMode(false);
      showToast('내장 웹 브라우저를 열었습니다. 하단 [① 원형스탬프]·[▣ 사각스탬프]·[펜 필기]·[웹 조작]을 바로 전환하세요.');
    }
  }

  function closeWebBrowser() {
    state.webActive = false;
    setWebInteractMode(false);
    if (webBar) webBar.style.display = 'none';
    if (webFrame) {
      webFrame.style.display = 'none';
      webFrame.src = 'about:blank';
    }
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
    showToast('웹 화면을 칠판 이미지로 고정했습니다. 이제 게시(Alt+`)와 PDF 저장에도 모두 포함됩니다.');
  }

  // ===================================================================
  // 모드 & 도구 전환
  // ===================================================================
  function setMode(nextMode) {
    commitInlineText();
    if (nextMode !== 'edit') clearSelection();
    state.mode = nextMode;
    if (state.webActive && state.webInteract) setWebInteractMode(false);
    document.querySelectorAll('.sb-mode-btn').forEach(function (b) {
      b.classList.toggle('active', b.getAttribute('data-sb-mode') === nextMode);
    });
    if (nextMode === 'zoom') {
      if (state.zoomScale < 1.3) state.zoomScale = 2.0;
      applyTransform();
      liveCanvas.style.cursor = 'zoom-in';
      showToast('확대 모드: 이동·휠로 배율 조절 후 클릭하면 그대로 판서 모드로 전환됩니다.');
    } else if (nextMode === 'pin') {
      liveCanvas.style.cursor = 'crosshair';
      showToast('게시 모드(스닙페이스트): 화면에 고정할 영역을 드래그하세요.');
    } else if (nextMode === 'edit') {
      liveCanvas.style.cursor = 'crosshair';
      showToast('편집 모드: 이동하거나 글씨로 변환(OCR)할 판서 영역을 드래그하세요.');
    } else {
      liveCanvas.style.cursor = state.tool === 'text' ? 'text' : 'crosshair';
    }
  }

  function setTool(nextTool) {
    commitInlineText();
    if (nextTool !== 'eraser') {
      state.lastDrawTool = nextTool;
    }
    if (nextTool !== 'select') clearSelection();
    state.tool = nextTool;
    if (state.webActive && state.webInteract) setWebInteractMode(false);
    if (state.mode !== 'draw') setMode('draw');
    document.querySelectorAll('.sb-tool-btn[data-sb-tool]').forEach(function (b) {
      b.classList.toggle('active', b.getAttribute('data-sb-tool') === nextTool);
    });
    syncWebQuickToolsUi();
    liveCanvas.style.cursor = nextTool === 'select' ? 'default' : (nextTool === 'text' ? 'text' : 'crosshair');
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

  function setSize(px) {
    state.size = Math.max(2, Math.min(36, Math.round(px)));
    if (el('sbSizeSlider')) el('sbSizeSlider').value = state.size;
    if (el('sbSizeLabel')) el('sbSizeLabel').textContent = state.size + 'px';
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
  function eraseAt(bx, by) {
    var pg = curPage();
    var er = Math.max(12, state.size * 2.6);
    var before = pg.objects.length;
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
    if (pg.objects.length !== before) renderInk();
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

  liveCanvas.addEventListener('pointerdown', function (ev) {
    if (ev.pointerType === 'touch') {
      pointer.pinchMap[ev.pointerId] = { x: ev.clientX, y: ev.clientY, type: ev.pointerType };
      var pIds = Object.keys(pointer.pinchMap);
      if (pIds.length === 2) {
        pointer.down = false;
        var p1 = pointer.pinchMap[pIds[0]], p2 = pointer.pinchMap[pIds[1]];
        pointer.pinchStartDist = Math.hypot(p2.x - p1.x, p2.y - p1.y);
        pointer.pinchStartZoom = state.zoomScale;
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
      showToast('확대 화면을 고정하고 판서 모드로 전환했습니다.');
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
    else if (ev.shiftKey && eff !== 'circle') eff = 'line';
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
        num: state.stampCounter++,
        color: state.color,
        radius: Math.max(18, Math.round(state.size * 2.6))
      });
      pg.redoStack = [];
      if (el('sbStampNum')) el('sbStampNum').textContent = state.stampCounter;
      renderInk();
      return;
    }
    if (eff === 'eraser') {
      eraseAt(pt.x, pt.y);
      return;
    }
    if (eff === 'laser') {
      state.laserPoints.push({ x: pt.x, y: pt.y, t: Date.now() });
      if (!state.laserRaf) state.laserRaf = requestAnimationFrame(renderLaserLoop);
      return;
    }
    if (eff === 'pen' || eff === 'highlighter') {
      pointer.currentStroke = {
        id: uid(),
        type: eff === 'highlighter' ? 'highlighter' : 'stroke',
        color: state.color,
        size: state.size,
        points: [{ x: pt.x, y: pt.y }]
      };
    }
  });

  liveCanvas.addEventListener('pointermove', function (ev) {
    if (pointer.pinchMap[ev.pointerId]) {
      pointer.pinchMap[ev.pointerId].x = ev.clientX;
      pointer.pinchMap[ev.pointerId].y = ev.clientY;
    }
    var pIds = Object.keys(pointer.pinchMap);
    if (pIds.length === 2 && pointer.pinchStartDist > 10) {
      var p1 = pointer.pinchMap[pIds[0]], p2 = pointer.pinchMap[pIds[1]];
      var dist = Math.hypot(p2.x - p1.x, p2.y - p1.y);
      state.zoomScale = Math.max(1, Math.min(4, pointer.pinchStartZoom * (dist / pointer.pinchStartDist)));
      if (state.zoomScale <= 1.02) { state.panX = 0; state.panY = 0; }
      applyTransform();
      return;
    }

    if (state.mode === 'zoom' && !pointer.down) {
      var sr = stageWrap.getBoundingClientRect();
      var nx = ((ev.clientX - sr.left) / sr.width) - 0.5;
      var ny = ((ev.clientY - sr.top) / sr.height) - 0.5;
      var pg = curPage();
      var maxPanX = (pg.width * state.baseScale * (state.zoomScale - 1)) / 2;
      var maxPanY = (pg.height * state.baseScale * (state.zoomScale - 1)) / 2;
      state.panX = Math.round(-nx * 2 * maxPanX);
      state.panY = Math.round(-ny * 2 * maxPanY);
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
      state.laserPoints.push({ x: pt.x, y: pt.y, t: Date.now() });
      if (!state.laserRaf) state.laserRaf = requestAnimationFrame(renderLaserLoop);
    } else if ((eff === 'pen' || eff === 'highlighter') && pointer.currentStroke) {
      pointer.currentStroke.points.push({ x: pt.x, y: pt.y });
      liveCtx.clearRect(0, 0, curPg.width, curPg.height);
      drawInkObject(liveCtx, pointer.currentStroke);
    } else if (eff === 'line' || eff === 'rect' || eff === 'fillRect' || eff === 'cover' || eff === 'circle') {
      liveCtx.clearRect(0, 0, curPg.width, curPg.height);
      if (eff === 'line') {
        drawInkObject(liveCtx, { type: 'line', x1: pointer.startX, y1: pointer.startY, x2: pt.x, y2: pt.y, color: state.color, size: state.size });
      } else {
        drawInkObject(liveCtx, { type: eff, x: pointer.startX, y: pointer.startY, w: pt.x - pointer.startX, h: pt.y - pointer.startY, color: state.color, size: state.size, revealed: false, perfect: !!ev.shiftKey });
      }
    }
    pointer.lastX = pt.x;
    pointer.lastY = pt.y;
  });

  function finishPointer(ev) {
    delete pointer.pinchMap[ev.pointerId];
    if (!pointer.down || ev.pointerId !== pointer.activeId) return;
    pointer.down = false;

    if (pointer.effectiveTool === 'pan') return;

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
    if ((eff === 'pen' || eff === 'highlighter') && pointer.currentStroke) {
      curPg.objects.push(pointer.currentStroke);
      curPg.redoStack = [];
      pointer.currentStroke = null;
      liveCtx.clearRect(0, 0, curPg.width, curPg.height);
      renderInk();
    } else if (eff === 'line' && pointer.moved) {
      curPg.objects.push({ id: uid(), type: 'line', x1: pointer.startX, y1: pointer.startY, x2: pt.x, y2: pt.y, color: state.color, size: state.size });
      curPg.redoStack = [];
      liveCtx.clearRect(0, 0, curPg.width, curPg.height);
      renderInk();
    } else if ((eff === 'rect' || eff === 'fillRect' || eff === 'cover' || eff === 'circle') && pointer.moved) {
      curPg.objects.push({ id: uid(), type: eff, x: pointer.startX, y: pointer.startY, w: pt.x - pointer.startX, h: pt.y - pointer.startY, color: state.color, size: state.size, revealed: false, perfect: !!ev.shiftKey });
      curPg.redoStack = [];
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
      var nextZ = state.zoomScale + (e.deltaY < 0 ? 0.2 : -0.2);
      state.zoomScale = Math.max(1, Math.min(4, Math.round(nextZ * 10) / 10));
      if (state.zoomScale <= 1.01) { state.panX = 0; state.panY = 0; }
      applyTransform();
    } else if (state.mode === 'draw') {
      setSize(state.size + (e.deltaY < 0 ? 1 : -1));
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
    if (!timing) return [];

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

        var isClickStart = (nt === 'clickEffect') || !steps.length;
        if (isClickStart) {
          steps.push({ targets: [] });
          seenInStep = {};
        }
        var curStep = steps[steps.length - 1];
        tgts.forEach(function (t) {
          var key = t.spid + ':' + t.pStart + ':' + t.pEnd + ':' + t.presetClass;
          if (!seenInStep[key]) {
            seenInStep[key] = true;
            curStep.targets.push(t);
          }
        });
      });
    });

    return steps.filter(function (s) { return s.targets.length > 0; });
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
              await parseShapeTree(ch, nextTf);
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
                  var fontColor = isDarkBox ? '#f8fafc' : (activeThemeMap.tx1 || '#0f172a');
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

        var steps = parseSlideTiming(sDoc);
        var slideObj = { bgColor: bgColor, bgDataUrl: bgDataUrl, elements: elements, steps: steps };
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
      state.currentNoteTitle = file.name.replace(/\.pptx$/i, '');
      updateNoteBadge();
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

  function computePptxVisibility(slide, stepIdx) {
    var shapeEntr = {}, shapeExit = {}, emphSet = {};
    var paraEntr = {}, paraExit = {};
    (slide.steps || []).forEach(function (st, idx) {
      var stepNum = idx + 1;
      (st.targets || []).forEach(function (tg) {
        if (tg.pStart >= 0) {
          for (var p = tg.pStart; p <= Math.max(tg.pStart, tg.pEnd); p++) {
            var pk = tg.spid + ':' + p;
            if (tg.presetClass === 'entr' && !paraEntr[pk]) paraEntr[pk] = stepNum;
            if (tg.presetClass === 'exit' && !paraExit[pk]) paraExit[pk] = stepNum;
          }
        } else {
          if (tg.presetClass === 'entr' && !shapeEntr[tg.spid]) shapeEntr[tg.spid] = stepNum;
          if (tg.presetClass === 'exit' && !shapeExit[tg.spid]) shapeExit[tg.spid] = stepNum;
        }
        if (tg.presetClass === 'emph' && stepNum === stepIdx) emphSet[tg.spid] = true;
      });
    });
    return { shapeEntr: shapeEntr, shapeExit: shapeExit, paraEntr: paraEntr, paraExit: paraExit, emphSet: emphSet };
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
      var reqStep = vis.shapeEntr[elm.spid] || 0;
      if (reqStep > 0 && stepIdx < reqStep) return;
      var hideStep = vis.shapeExit[elm.spid] || 0;
      if (hideStep > 0 && stepIdx >= hideStep) return;

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
            var pReq = vis.paraEntr[pk] || 0;
            if (pReq > 0 && stepIdx < pReq) return;
            var pHide = vis.paraExit[pk] || 0;
            if (pHide > 0 && stepIdx >= pHide) return;

            ctx.fillStyle = p.color || '#0f172a';
            ctx.font = (p.bold ? 'bold ' : '') + p.fontSize + 'px Pretendard, "Noto Sans KR", "Malgun Gothic", sans-serif';
            ctx.textBaseline = 'top';
            ctx.textAlign = p.align === 'ctr' ? 'center' : (p.align === 'r' ? 'right' : 'left');
            var tx = p.align === 'ctr' ? (elm.x + elm.w / 2) : (p.align === 'r' ? (elm.x + elm.w - padX) : (elm.x + padX));
            wrapCanvasText(ctx, p.text, maxW).forEach(function (ln) {
              ctx.fillText(ln, tx, py);
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
      await new Promise(function (resolve) {
        var tx = db.transaction(IDB_STORE, 'readwrite');
        tx.objectStore(IDB_STORE).put(noteRecord);
        tx.oncomplete = function () { resolve(true); };
        tx.onerror = function () { resolve(false); };
      });
      return;
    }
    var list = await dbGetAllNotes();
    var idx = list.findIndex(function (n) { return n.id === noteRecord.id; });
    if (idx >= 0) list[idx] = noteRecord;
    else list.unshift(noteRecord);
    try { localStorage.setItem(FALLBACK_NOTES_KEY, JSON.stringify(list)); } catch (e) {}
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
  }

  async function serializeCurrentNote(asNew) {
    commitInlineText();
    var noteId = (!asNew && state.currentNoteId) ? state.currentNoteId : ('note_' + Date.now() + '_' + Math.floor(Math.random() * 1000));
    var titleInput = el('sbVaultNoteTitleInput');
    var folderSel = el('sbVaultFolderSelect');
    var title = (titleInput && titleInput.value.trim()) || state.currentNoteTitle || '새 판서 노트';
    var folder = (folderSel && folderSel.value) || state.currentFolder || loadFolders()[0];

    // 1쪽 썸네일 생성 (320x180 JPEG)
    var thumbCv = document.createElement('canvas');
    thumbCv.width = 320; thumbCv.height = 180;
    thumbCv.getContext('2d').drawImage(createCompositeCanvas(state.pages[0]), 0, 0, 320, 180);
    var thumbUrl = thumbCv.toDataURL('image/jpeg', 0.72);

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
          objects: JSON.parse(JSON.stringify(pg.objects || []))
        });
      } else if (pg.kind === 'pptx' && pg.pptxSlide) {
        var cleanSlide = {
          bgColor: pg.pptxSlide.bgColor,
          bgDataUrl: pg.pptxSlide.bgDataUrl || null,
          steps: pg.pptxSlide.steps || [],
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
          objects: JSON.parse(JSON.stringify(pg.objects || []))
        });
      } else {
        serializedPages.push({
          kind: 'blank',
          label: pg.label || '칠판',
          width: pg.width || 1280,
          height: pg.height || 720,
          bgMode: pg.bgMode || 'dark',
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
    await dbPutNote(record);
    state.currentNoteId = noteId;
    state.currentNoteTitle = title;
    state.currentFolder = folder;
    updateNoteBadge();
    await renderVaultUi();
    showToast('[' + folder + '] 폴더에 "' + title + '" 노트를 저장했습니다.');
  }

  async function loadNoteFromVault(noteId) {
    var notes = await dbGetAllNotes();
    var target = notes.find(function (n) { return n.id === noteId; });
    if (!target || !Array.isArray(target.pages) || !target.pages.length) return;
    closeWebBrowser();
    state.currentNoteId = target.id;
    state.currentNoteTitle = target.title;
    state.currentFolder = target.folder;
    state.pages = target.pages.map(function (p) {
      return Object.assign({}, p, { redoStack: [] });
    });
    state.pageIndex = 0;
    updateNoteBadge();
    resetZoom();
    resizeBoardToPage();
    await renderBackground();
    renderInk();
    el('sbVaultModal').classList.remove('show');
    showToast('"' + target.title + '" (' + state.pages.length + '쪽) 노트를 불러왔습니다.');
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

    var filtered = state.activeVaultFolder === 'ALL'
      ? notes
      : notes.filter(function (n) { return n.folder === state.activeVaultFolder; });
    if (el('sbVaultActiveFolderTitle')) el('sbVaultActiveFolderTitle').textContent = state.activeVaultFolder === 'ALL' ? '전체 노트' : state.activeVaultFolder;
    if (el('sbVaultNoteCount')) el('sbVaultNoteCount').textContent = filtered.length + '개 노트';

    var grid = el('sbVaultNotesGrid');
    if (!grid) return;
    if (!filtered.length) {
      grid.innerHTML = '<div style="grid-column:1/-1;padding:36px 14px;text-align:center;color:#64748b;font-size:0.86rem;">이 폴더에 저장된 판서 노트가 없습니다. 상단의 <b>[현재 노트 덮어쓰기 저장]</b> 또는 <b>[+ 새 노트로 저장]</b>을 눌러 저장해 보세요.</div>';
      return;
    }
    grid.innerHTML = filtered.map(function (n) {
      return '<div class="sb-note-card">' +
        '<div class="sb-note-thumb">' + (n.thumbUrl ? '<img src="' + n.thumbUrl + '" alt="">' : '') + '</div>' +
        '<div class="sb-note-meta">' +
          '<strong>' + (n.title || '판서 노트') + '</strong>' +
          '<small>' + (n.folder || '기본') + ' · ' + (n.pageCount || 1) + '쪽</small>' +
          '<small>' + (n.updatedStr || '') + '</small>' +
          '<div class="sb-note-actions">' +
            '<button type="button" class="sb-btn sb-btn-primary sb-btn-xs" style="flex:1;" data-load-note="' + n.id + '">열기</button>' +
            '<button type="button" class="sb-btn sb-btn-danger sb-btn-xs" data-del-note="' + n.id + '">삭제</button>' +
          '</div>' +
        '</div>' +
      '</div>';
    }).join('');

    grid.querySelectorAll('[data-load-note]').forEach(function (lb) {
      lb.addEventListener('click', function () { loadNoteFromVault(lb.getAttribute('data-load-note')); });
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
    commitInlineText();
    clearSelection();
    state.pageIndex = idx;
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
  // PNG 저장 (Ctrl + S) & 전체 페이지 합본 PDF 저장
  // ===================================================================
  function saveCurrentPng() {
    var comp = createCompositeCanvas();
    var a = document.createElement('a');
    a.href = comp.toDataURL('image/png');
    a.download = '스마트판서_' + (state.pageIndex + 1) + '쪽_' + new Date().toISOString().slice(0, 10) + '.png';
    a.click();
    showToast('현재 화면을 PNG 이미지로 저장했습니다.');
  }

  async function saveMergedPdf() {
    if (typeof PDFLib === 'undefined') { showToast('pdf-lib 라이브러리를 불러오지 못했습니다.'); return; }
    try {
      showToast('전체 ' + state.pages.length + '페이지 판서 합본 PDF를 생성하는 중...');
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
      var blob = new Blob([bytes], { type: 'application/pdf' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = '스마트판서_합본_' + new Date().toISOString().slice(0, 10) + '.pdf';
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 5000);
      showToast('전체 ' + state.pages.length + '페이지 판서 합본 PDF를 저장했습니다.');
    } catch (err) {
      showToast('합본 PDF 저장 오류: ' + (err.message || err));
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
        this.textContent = state.autoHide ? '자동숨김: 켜짐' : '자동숨김: 꺼짐';
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
        if (el('sbHelpModal')) el('sbHelpModal').classList.add('show');
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
    if (el('sbVaultSaveNewBtn')) el('sbVaultSaveNewBtn').addEventListener('click', function () { serializeCurrentNote(true); });
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
    if (el('sbVaultExportJsonBtn')) el('sbVaultExportJsonBtn').addEventListener('click', async function () {
      var payload = { folders: loadFolders(), notes: await dbGetAllNotes() };
      var blob = new Blob([JSON.stringify(payload)], { type: 'application/json' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = '스마트판서_노트보관함_백업_' + new Date().toISOString().slice(0, 10) + '.json';
      a.click();
      showToast('노트 보관함 전체 백업(.json)을 다운로드했습니다.');
    });
    if (el('sbVaultImportJsonInput')) el('sbVaultImportJsonInput').addEventListener('change', function () {
      var f = this.files && this.files[0];
      if (!f) return;
      var fr = new FileReader();
      fr.onload = async function () {
        try {
          var data = JSON.parse(fr.result);
          if (Array.isArray(data.folders) && data.folders.length) saveFolders(data.folders);
          if (Array.isArray(data.notes)) {
            for (var i = 0; i < data.notes.length; i++) await dbPutNote(data.notes[i]);
          }
          await renderVaultUi();
          showToast('백업 파일에서 노트를 복원했습니다.');
        } catch (err) { showToast('백업 파일 형식이 올바르지 않습니다.'); }
      };
      fr.readAsText(f);
      this.value = '';
    });

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
        showToast('웹 조작 모드: 웹페이지 스크롤 및 유튜브 재생 버튼 클릭이 가능합니다.');
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
      b.addEventListener('click', function () { setMode(b.getAttribute('data-sb-mode')); });
    });

    // 도구 선택
    document.querySelectorAll('.sb-tool-btn[data-sb-tool]').forEach(function (b) {
      b.addEventListener('click', function () { setTool(b.getAttribute('data-sb-tool')); });
    });

    // 색상 & 굵기
    document.querySelectorAll('#sbColorGroup .sb-color-dot').forEach(function (d) {
      d.addEventListener('click', function () { setColor(d.getAttribute('data-color')); });
    });
    if (el('sbSizeSlider')) {
      el('sbSizeSlider').addEventListener('input', function () { setSize(parseInt(this.value, 10)); });
    }

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
        state.palmReject = !state.palmReject;
        this.classList.toggle('active', state.palmReject);
        this.textContent = state.palmReject ? 'S펜전용: 켜짐' : 'S펜전용: 꺼짐';
        showToast(state.palmReject ? 'S펜 전용 모드 ON (손바닥 무시 · 손가락으로 확대 이동)' : 'S펜 전용 모드 OFF (손가락·펜 모두 필기)');
      });
    }

    // 통합 파일·링크 열기 모달 (#sbFileHubModal)
    function closeFileHubModal() {
      if (el('sbFileHubModal')) el('sbFileHubModal').classList.remove('show');
    }
    function handleUnifiedFile(f) {
      if (!f) return;
      closeFileHubModal();
      var name = String(f.name || '').toLowerCase();
      if (name.endsWith('.pdf') || f.type === 'application/pdf') {
        loadPdfFile(f);
      } else if (name.endsWith('.pptx')) {
        loadPptxFile(f);
      } else if (f.type && f.type.indexOf('image/') === 0 || /\.(png|jpe?g|gif|webp|bmp)$/i.test(name)) {
        var fr = new FileReader();
        fr.onload = function () { loadImageFromDataUrl(fr.result, f.name); };
        fr.readAsDataURL(f);
      } else {
        showToast('지원 형식은 PDF, PPTX, 이미지(PNG/JPG)입니다.');
      }
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
    if (el('sbUndoBtn')) el('sbUndoBtn').addEventListener('click', function () {
      var pg = curPage();
      if (!pg.objects.length) return;
      pg.redoStack.push(pg.objects.pop());
      renderInk();
    });
    if (el('sbRedoBtn')) el('sbRedoBtn').addEventListener('click', function () {
      var pg = curPage();
      if (!pg.redoStack.length) return;
      pg.objects.push(pg.redoStack.pop());
      renderInk();
    });
    if (el('sbClearInkBtn')) el('sbClearInkBtn').addEventListener('click', function () {
      var pg = curPage();
      if (!pg.objects.length) return;
      pg.objects = [];
      pg.redoStack = [];
      clearSelection();
      renderInk();
      showToast('현재 페이지 판서를 모두 삭제했습니다.');
    });

    if (el('sbZoomResetBtn')) el('sbZoomResetBtn').addEventListener('click', function () { resetZoom(); if (state.mode === 'zoom') setMode('draw'); });
    inlineTextEl.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); commitInlineText(); }
      else if (e.key === 'Escape') { inlineTextEl.style.display = 'none'; pendingTextPos = null; }
    });

    // 편집 모드 액션 바
    if (el('sbEditOcrBtn')) el('sbEditOcrBtn').addEventListener('click', runOcrOnSelection);
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
      state.stampCounter = 1;
      if (el('sbStampNum')) el('sbStampNum').textContent = '1';
      showToast('번호 스탬프를 1번부터 다시 시작합니다.');
    });
    if (el('sbClearPinsBtn')) el('sbClearPinsBtn').addEventListener('click', function () {
      pinLayer.innerHTML = '';
      showToast('모든 게시 조각을 닫았습니다.');
    });

    // 저장 / 단축키 전체창(좌우 슬라이드) / 사용설명서 2단계 뷰(큰 글씨 요약 <-> 전체화면 GIF 좌우 슬라이드 과정)
    if (el('sbSavePngBtn')) el('sbSavePngBtn').addEventListener('click', saveCurrentPng);
    if (el('sbSavePdfBtn')) el('sbSavePdfBtn').addEventListener('click', saveMergedPdf);
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

    if (el('sbHelpBtn')) el('sbHelpBtn').addEventListener('click', function () {
      goVisSlide(0);
      el('sbHelpModal').classList.add('show');
    });
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

    // 전역 단축키
    document.addEventListener('keydown', function (e) {
      var inInput = document.activeElement && (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'TEXTAREA' || document.activeElement.tagName === 'SELECT');
      if (inInput) return;

      if (e.key === 'Tab') { e.preventDefault(); togglePalette(); return; }
      if (e.key === '?' || e.key === 'F1') { e.preventDefault(); el('sbShortcutModal').classList.toggle('show'); return; }
      // 백스페이스(Backspace) 또는 Delete: 편집 모드 선택 영역 삭제 또는 마지막 판서 취소(되돌리기)
      if (e.key === 'Backspace' || e.key === 'Delete') {
        e.preventDefault();
        if ((state.mode === 'edit' || state.tool === 'select') && state.selection && el('sbEditDelBtn')) {
          el('sbEditDelBtn').click();
        } else if (el('sbUndoBtn')) {
          el('sbUndoBtn').click();
          showToast('마지막 판서 동작을 취소했습니다. (Backspace)');
        }
        return;
      }
      if (e.altKey && (e.code === 'Backquote' || e.key === '`')) { e.preventDefault(); setMode(state.mode === 'pin' ? 'draw' : 'pin'); return; }
      if (e.altKey && e.key === '1') { e.preventDefault(); setMode(state.mode === 'zoom' ? 'draw' : 'zoom'); return; }
      if (e.altKey && e.key === '2') { e.preventDefault(); setMode('draw'); return; }
      if (e.altKey && e.key === '3') { e.preventDefault(); setMode(state.mode === 'edit' ? 'draw' : 'edit'); return; }
      if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 's') { e.preventDefault(); serializeCurrentNote(false); return; }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); saveCurrentPng(); return; }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); el('sbUndoBtn').click(); return; }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') { e.preventDefault(); el('sbRedoBtn').click(); return; }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'e') { e.preventDefault(); el('sbClearInkBtn').click(); return; }

      if (e.key === 'Escape') {
        document.querySelectorAll('.sb-modal.show').forEach(function (m) { m.classList.remove('show'); });
        clearSelection();
        if (state.mode !== 'draw') setMode('draw');
        if (state.zoomScale > 1) resetZoom();
        return;
      }

      if (!e.ctrlKey && !e.altKey && !e.metaKey) {
        var k = e.key.toLowerCase();
        // 도형 및 도구 1글자 단축키: 선택·이동(V), 지우개(E), 원형스탬프(S), 사각스탬프(Q), 텍스트(T), 사각형(R), 원(C), 직선(L), 펜(P), 형광펜(H)
        if (k === 'v') setTool('select');
        else if (k === 'e') setTool(state.tool === 'eraser' ? (state.lastDrawTool || 'pen') : 'eraser');
        else if (k === 's') setTool(state.tool === 'stamp' ? (state.lastDrawTool || 'pen') : 'stamp');
        else if (k === 'q') setTool(state.tool === 'sqStamp' ? (state.lastDrawTool || 'pen') : 'sqStamp');
        else if (k === 't') setTool(state.tool === 'text' ? (state.lastDrawTool || 'pen') : 'text');
        else if (k === 'r') setTool(e.shiftKey ? 'fillRect' : 'rect');
        else if (k === 'c') setTool('circle');
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
        else if (e.key === ' ' || e.key === 'ArrowRight') {
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
          else if (e.key === 'ArrowRight') goPage(state.pageIndex + 1);
        } else if (e.key === 'ArrowLeft') {
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
          if (curPage().kind === 'pptx') { e.preventDefault(); stepPptxAction(-1); }
          else goPage(state.pageIndex - 1);
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
})();
