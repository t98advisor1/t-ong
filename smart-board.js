/* ===================================================================
 * T-ong 스마트 판서 보드 (PDF · PPTX 클릭 액션 · SaniZ 4대 모드 · 무료 OCR)
 *  - 독립 전체화면 판서 스튜디오 (갤럭시탭 S펜 + PC 단축키 동시 지원)
 * =================================================================== */
(function () {
  'use strict';

  var OCR_CFG_KEY = 'tong_smart_board_ocr_cfg_v1';
  var MEMO_NOTES_KEY = 'tong_memo_board_notes_v1';

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

  var state = {
    mode: 'draw',          // 'pin' | 'zoom' | 'draw' | 'edit'
    tool: 'pen',           // 'pen' | 'highlighter' | 'laser' | 'line' | 'rect' | 'fillRect' | 'cover' | 'text' | 'stamp' | 'eraser'
    color: '#ef4444',
    size: 5,
    bgMode: 'dark',        // 'dark' | 'light' | 'grid'
    palmReject: false,
    stampCounter: 1,
    zoomScale: 1,
    panX: 0,
    panY: 0,
    baseScale: 1,
    pages: [],
    pageIndex: 0,
    pdfDoc: null,
    pdfCache: {},
    selection: null,       // { x, y, w, h, ids: [] }
    laserPoints: [],
    laserRaf: null,
    ocrPendingBox: null
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
    toastTimer = setTimeout(function () { t.style.display = 'none'; }, 2400);
  }

  function uid() { return 'o_' + Date.now() + '_' + Math.floor(Math.random() * 10000); }

  function curPage() {
    if (!state.pages.length) {
      state.pages.push(createBlankPage('dark'));
      state.pageIndex = 0;
    }
    return state.pages[state.pageIndex];
  }

  function createBlankPage(bgMode) {
    return {
      kind: 'blank',
      label: bgMode === 'light' ? '백지' : (bgMode === 'grid' ? '모눈' : '칠판'),
      width: 1280,
      height: 720,
      bgMode: bgMode || state.bgMode,
      objects: [],
      redoStack: []
    };
  }

  // ===================================================================
  // 뷰포트 크기 맞춤 및 좌표 변환
  // ===================================================================
  function resizeBoardToPage() {
    var pg = curPage();
    var w = pg.width || 1280;
    var h = pg.height || 720;
    [bgCanvas, inkCanvas, liveCanvas].forEach(function (cv) {
      if (cv.width !== w || cv.height !== h) {
        cv.width = w;
        cv.height = h;
      }
    });
    boardFrame.style.width = w + 'px';
    boardFrame.style.height = h + 'px';
    fitViewport();
  }

  function fitViewport() {
    var pg = curPage();
    var rect = stageWrap.getBoundingClientRect();
    var pad = 20;
    var sx = Math.max(0.2, (rect.width - pad) / (pg.width || 1280));
    var sy = Math.max(0.2, (rect.height - pad) / (pg.height || 720));
    state.baseScale = Math.min(sx, sy);
    applyTransform();
  }

  function applyTransform() {
    var totalScale = state.baseScale * state.zoomScale;
    boardFrame.style.transform = 'translate(' + state.panX + 'px, ' + state.panY + 'px) scale(' + totalScale + ')';
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
    var x = ((clientX - r.left) / r.width) * pg.width;
    var y = ((clientY - r.top) / r.height) * pg.height;
    return { x: x, y: y };
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
    } else if (pg.kind === 'image' && pg.imageObj) {
      bgCtx.fillStyle = '#0f172a';
      bgCtx.fillRect(0, 0, w, h);
      bgCtx.drawImage(pg.imageObj, 0, 0, w, h);
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
      renderPptxSlideToCtx(pg.pptxSlide, pg.pptxStep || 0, bgCtx, w, h);
    }
    updateFooterUi();
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
      for (var x = 0; x <= w; x += 32) {
        ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, h); ctx.stroke();
      }
      for (var y = 0; y <= h; y += 32) {
        ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(w, y); ctx.stroke();
      }
      // 좌표축 보조선
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
    if (o.type === 'rect' || o.type === 'fillRect' || o.type === 'cover') {
      return { x: Math.min(o.x, o.x + o.w), y: Math.min(o.y, o.y + o.h), w: Math.abs(o.w), h: Math.abs(o.h) };
    }
    if (o.type === 'stamp') {
      var r = o.radius || 20;
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
    } else if (o.type === 'rect' || o.type === 'fillRect' || o.type === 'cover' || o.type === 'stamp' || o.type === 'text') {
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
    } else if (o.type === 'fillRect') {
      ctx.fillStyle = o.color || '#ef4444';
      ctx.globalAlpha = 0.88;
      ctx.fillRect(o.x, o.y, o.w, o.h);
    } else if (o.type === 'cover') {
      var rx = Math.min(o.x, o.x + o.w), ry = Math.min(o.y, o.y + o.h);
      var rw = Math.abs(o.w), rh = Math.abs(o.h);
      if (o.revealed) {
        ctx.strokeStyle = '#38bdf8';
        ctx.setLineDash([6, 4]);
        ctx.lineWidth = 2;
        ctx.strokeRect(rx, ry, rw, rh);
      } else {
        ctx.fillStyle = '#1e293b';
        ctx.fillRect(rx, ry, rw, rh);
        ctx.strokeStyle = '#38bdf8';
        ctx.lineWidth = 2.5;
        ctx.strokeRect(rx, ry, rw, rh);
        ctx.fillStyle = '#bae6fd';
        ctx.font = 'bold 16px Pretendard, "Noto Sans KR", sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        if (rw > 70 && rh > 26) ctx.fillText('탭하여 공개', rx + rw / 2, ry + rh / 2);
      }
    } else if (o.type === 'stamp') {
      var rad = o.radius || 20;
      ctx.fillStyle = o.color || '#ef4444';
      ctx.beginPath();
      ctx.arc(o.x, o.y, rad, 0, Math.PI * 2);
      ctx.fill();
      ctx.strokeStyle = '#ffffff';
      ctx.lineWidth = 2.5;
      ctx.stroke();
      ctx.fillStyle = (o.color === '#facc15' || o.color === '#f8fafc') ? '#0f172a' : '#ffffff';
      ctx.font = '900 ' + Math.round(rad * 1.1) + 'px Pretendard, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(String(o.num || 1), o.x, o.y + 1);
    } else if (o.type === 'text') {
      var fs = o.fontSize || 26;
      ctx.fillStyle = o.color || '#f8fafc';
      ctx.font = 'bold ' + fs + 'px Pretendard, "Noto Sans KR", "Malgun Gothic", sans-serif';
      ctx.textBaseline = 'top';
      var lines = String(o.text || '').split('\n');
      lines.forEach(function (ln, li) {
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
    if (state.mode === 'edit' && state.selection) {
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
  // 레이저 포인터 (잔상 자동 소멸)
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
  // 게시 모드 조각 생성 (항상 최상단 · 드래그 이동 · 휠/핸들 리사이즈 · 우클릭/X 삭제)
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
      else if (pg.kind === 'pptx' && pg.pptxSlide) renderPptxSlideToCtx(pg.pptxSlide, pg.pptxStep || 0, c, pg.width, pg.height);
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
    var left = Math.max(84, Math.min(window.innerWidth - dispW - 20, fr.left + (sx / comp.width) * fr.width));
    var top = Math.max(58, Math.min(window.innerHeight - dispH - 60, fr.top + (sy / comp.height) * fr.height));
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
    closeBtn.title = '조각 삭제 (우클릭 가능)';
    closeBtn.addEventListener('click', function (e) { e.stopPropagation(); pin.remove(); });

    var resHandle = document.createElement('div');
    resHandle.className = 'sb-pin-resize';
    resHandle.title = '드래그 또는 마우스 휠로 크기 조절';

    pin.appendChild(img);
    pin.appendChild(closeBtn);
    pin.appendChild(resHandle);
    pinLayer.appendChild(pin);

    // 우클릭으로 삭제 (SaniZ 호환)
    pin.addEventListener('contextmenu', function (e) {
      e.preventDefault();
      pin.remove();
      showToast('게시 조각을 삭제했습니다.');
    });

    // 마우스 휠로 크기 조절 (SaniZ 호환)
    pin.addEventListener('wheel', function (e) {
      e.preventDefault();
      var ratio = e.deltaY < 0 ? 1.09 : 0.91;
      var curW = parseFloat(pin.style.width);
      var newW = Math.max(90, Math.min(window.innerWidth * 0.85, curW * ratio));
      pin.style.width = Math.round(newW) + 'px';
      pin.style.height = Math.round(newW * (sh / sw)) + 'px';
    }, { passive: false });

    // 드래그 이동
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

    // 모서리 드래그 리사이즈 (태블릿용)
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
  // 모드 & 도구 전환
  // ===================================================================
  function setMode(nextMode) {
    commitInlineText();
    if (nextMode !== 'edit') clearSelection();
    state.mode = nextMode;
    document.querySelectorAll('.sb-mode-btn').forEach(function (b) {
      b.classList.toggle('active', b.getAttribute('data-sb-mode') === nextMode);
    });
    if (nextMode === 'zoom') {
      if (state.zoomScale < 1.3) state.zoomScale = 2.0;
      applyTransform();
      liveCanvas.style.cursor = 'zoom-in';
      showToast('확대 모드: 마우스 이동·휠로 조절 후 클릭하면 그대로 판서 모드로 전환됩니다.');
    } else if (nextMode === 'pin') {
      liveCanvas.style.cursor = 'crosshair';
      showToast('게시 모드: 화면에 고정할 영역을 드래그하세요.');
    } else if (nextMode === 'edit') {
      liveCanvas.style.cursor = 'crosshair';
      showToast('편집 모드: 이동하거나 글씨로 변환(OCR)할 판서 영역을 드래그하세요.');
    } else {
      liveCanvas.style.cursor = state.tool === 'text' ? 'text' : 'crosshair';
    }
  }

  function setTool(nextTool) {
    commitInlineText();
    state.tool = nextTool;
    if (state.mode !== 'draw') setMode('draw');
    document.querySelectorAll('.sb-tool-btn[data-sb-tool]').forEach(function (b) {
      b.classList.toggle('active', b.getAttribute('data-sb-tool') === nextTool);
    });
    liveCanvas.style.cursor = nextTool === 'text' ? 'text' : 'crosshair';
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
      // PDF/PPTX 중 배경 버튼 클릭 시 새 칠판/백지/모눈 페이지 추가
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
    pendingTextPos = { x: bx, y: by };
    var fs = Math.max(18, state.size * 3.8);
    inlineTextEl.style.left = Math.max(8, bx) + 'px';
    inlineTextEl.style.top = Math.max(8, by - fs * 0.5) + 'px';
    inlineTextEl.style.fontSize = fs + 'px';
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
    // 두 손가락 핀치 확대 추적
    pointer.pinchMap[ev.pointerId] = { x: ev.clientX, y: ev.clientY, type: ev.pointerType };
    var pIds = Object.keys(pointer.pinchMap);
    if (pIds.length === 2) {
      pointer.down = false;
      var p1 = pointer.pinchMap[pIds[0]], p2 = pointer.pinchMap[pIds[1]];
      pointer.pinchStartDist = Math.hypot(p2.x - p1.x, p2.y - p1.y);
      pointer.pinchStartZoom = state.zoomScale;
      return;
    }

    // 갤럭시탭 S펜 전용 모드: 손가락 터치는 판서 무시하고 확대 시 화면 이동만 수행
    if (state.palmReject && ev.pointerType === 'touch' && state.mode === 'draw') {
      pointer.down = true;
      pointer.activeId = ev.pointerId;
      pointer.effectiveTool = 'pan';
      pointer.lastX = ev.clientX;
      pointer.lastY = ev.clientY;
      liveCanvas.setPointerCapture(ev.pointerId);
      return;
    }

    // 확대 모드: 좌클릭 시 그 확대 화면 그대로 멈추고 판서 모드 전환 (SaniZ 호환)
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
    liveCanvas.setPointerCapture(ev.pointerId);

    if (state.mode === 'pin') {
      pointer.effectiveTool = 'pin_select';
      return;
    }

    if (state.mode === 'edit') {
      if (state.selection &&
          pt.x >= state.selection.x && pt.x <= state.selection.x + state.selection.w &&
          pt.y >= state.selection.y && pt.y <= state.selection.y + state.selection.h) {
        pointer.draggingSelection = true;
      } else {
        pointer.draggingSelection = false;
        clearSelection();
      }
      return;
    }

    // 판서 모드 (S펜 배럴 버튼 buttons===32 또는 eraser 펜촉 -> 임시 지우개!)
    var isSpenEraser = (ev.pointerType === 'pen' && (ev.buttons === 32 || ev.button === 5));
    var eff = state.tool;
    if (isSpenEraser) eff = 'eraser';
    else if (ev.ctrlKey && ev.shiftKey) eff = 'fillRect';
    else if (ev.ctrlKey) eff = 'rect';
    else if (ev.shiftKey) eff = 'line';
    pointer.effectiveTool = eff;

    if (eff === 'text') {
      pointer.down = false;
      openInlineText(pt.x, pt.y);
      return;
    }
    if (eff === 'stamp') {
      pointer.down = false;
      var pg = curPage();
      pg.objects.push({
        id: uid(),
        type: 'stamp',
        x: pt.x,
        y: pt.y,
        num: state.stampCounter++,
        color: state.color,
        radius: Math.max(16, state.size * 2.5)
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
    // 핀치 줌 처리
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

    // 확대 모드: 마우스 커서를 따라 부드럽게 이동 (SaniZ 호환)
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

    if (state.mode === 'edit') {
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
    } else if (eff === 'line' || eff === 'rect' || eff === 'fillRect' || eff === 'cover') {
      liveCtx.clearRect(0, 0, curPg.width, curPg.height);
      if (eff === 'line') {
        drawInkObject(liveCtx, { type: 'line', x1: pointer.startX, y1: pointer.startY, x2: pt.x, y2: pt.y, color: state.color, size: state.size });
      } else {
        drawInkObject(liveCtx, { type: eff, x: pointer.startX, y: pointer.startY, w: pt.x - pointer.startX, h: pt.y - pointer.startY, color: state.color, size: state.size, revealed: false });
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
      setMode('draw'); // 영역 지정 완료 시 자동 해제 (SaniZ 호환)
      return;
    }

    if (state.mode === 'edit') {
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

    // 탭하여 정답 가림막 열기/닫기
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
    } else if ((eff === 'rect' || eff === 'fillRect' || eff === 'cover') && pointer.moved) {
      curPg.objects.push({ id: uid(), type: eff, x: pointer.startX, y: pointer.startY, w: pt.x - pointer.startX, h: pt.y - pointer.startY, color: state.color, size: state.size, revealed: false });
      curPg.redoStack = [];
      liveCtx.clearRect(0, 0, curPg.width, curPg.height);
      renderInk();
    }
  }

  liveCanvas.addEventListener('pointerup', finishPointer);
  liveCanvas.addEventListener('pointercancel', finishPointer);

  // 마우스 휠: 확대 모드에서는 배율 조절, 판서 모드에서는 굵기 조절 (SaniZ 호환)
  stageWrap.addEventListener('wheel', function (e) {
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
  // PPTX 직접 파서 + 클릭 액션(<p:timing>) 재생 엔진 (JSZip + OOXML)
  // ===================================================================
  function xmlAttr(node, name) { return node ? (node.getAttribute(name) || '') : ''; }
  function qAll(root, localName) {
    return Array.from(root.getElementsByTagName('*')).filter(function (n) { return n.localName === localName; });
  }
  function qOne(root, localName) {
    var list = qAll(root, localName);
    return list.length ? list[0] : null;
  }

  function extractFillColor(node) {
    if (!node) return null;
    var srgb = qOne(node, 'srgbClr');
    if (srgb && xmlAttr(srgb, 'val')) return '#' + xmlAttr(srgb, 'val');
    var scheme = qOne(node, 'schemeClr');
    if (scheme) {
      var v = xmlAttr(scheme, 'val');
      var map = { bg1: '#ffffff', lt1: '#ffffff', tx1: '#0f172a', dk1: '#0f172a', accent1: '#2563eb', accent2: '#ea580c', accent3: '#16a34a', accent4: '#9333ea', accent5: '#0891b2', accent6: '#d97706' };
      return map[v] || null;
    }
    return null;
  }

  function parseSlideTiming(doc) {
    // OOXML <p:timing>에서 클릭 시퀀스(nodeType="clickEffect" 또는 clickPar) 추출
    var timing = qOne(doc, 'timing');
    if (!timing) return [];
    var steps = [];
    var cTns = qAll(timing, 'cTn');
    cTns.forEach(function (ctn) {
      var nt = xmlAttr(ctn, 'nodeType');
      var presetClass = xmlAttr(ctn, 'presetClass'); // entr | exit | emph
      if (presetClass === 'entr' || presetClass === 'exit' || presetClass === 'emph') {
        var spTgt = qOne(ctn, 'spTgt');
        var spid = spTgt ? xmlAttr(spTgt, 'spid') : '';
        if (spid) {
          var isWithPrev = (nt === 'withEffect' || nt === 'afterEffect');
          if (isWithPrev && steps.length > 0) {
            steps[steps.length - 1].targets.push({ spid: spid, presetClass: presetClass });
          } else {
            steps.push({ targets: [{ spid: spid, presetClass: presetClass }] });
          }
        }
      }
    });
    return steps;
  }

  async function loadPptxFile(file) {
    if (typeof JSZip === 'undefined') { showToast('JSZip 라이브러리를 불러오지 못했습니다.'); return; }
    try {
      showToast('PPTX 슬라이드와 클릭 액션을 분석하는 중...');
      var zip = await JSZip.loadAsync(await file.arrayBuffer());
      var parser = new DOMParser();

      // 1. presentation.xml에서 슬라이드 크기와 순서 읽기
      var presXmlStr = await zip.file('ppt/presentation.xml').async('string');
      var presDoc = parser.parseFromString(presXmlStr, 'application/xml');
      var sldSz = qOne(presDoc, 'sldSz');
      var slideCx = parseInt(xmlAttr(sldSz, 'cx'), 10) || 9144000;
      var slideCy = parseInt(xmlAttr(sldSz, 'cy'), 10) || 5143500;
      var bw = 1280;
      var bh = Math.round(bw * (slideCy / slideCx));

      var presRelsFile = zip.file('ppt/_rels/presentation.xml.rels');
      var relMap = {};
      if (presRelsFile) {
        var rDoc = parser.parseFromString(await presRelsFile.async('string'), 'application/xml');
        qAll(rDoc, 'Relationship').forEach(function (rel) {
          relMap[xmlAttr(rel, 'Id')] = xmlAttr(rel, 'Target');
        });
      }

      var slidePaths = [];
      qAll(presDoc, 'sldId').forEach(function (sid) {
        var rid = sid.getAttribute('r:id') || sid.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
        if (rid && relMap[rid]) {
          var t = relMap[rid].replace(/^\/+/, '');
          slidePaths.push(t.indexOf('ppt/') === 0 ? t : ('ppt/' + t));
        }
      });
      if (!slidePaths.length) {
        slidePaths = Object.keys(zip.files).filter(function (k) { return /^ppt\/slides\/slide\d+\.xml$/i.test(k); })
          .sort(function (a, b) { return parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10); });
      }

      var mediaImgCache = {};
      async function getMediaImg(relTarget) {
        var norm = relTarget.replace(/^\.\.\//, 'ppt/').replace(/^\/+/, '');
        if (norm.indexOf('ppt/') !== 0) norm = 'ppt/slides/' + norm;
        if (mediaImgCache[norm]) return mediaImgCache[norm];
        var mf = zip.file(norm);
        if (!mf) return null;
        var ext = norm.split('.').pop().toLowerCase();
        var mime = ext === 'png' ? 'image/png' : (ext === 'gif' ? 'image/gif' : (ext === 'svg' ? 'image/svg+xml' : 'image/jpeg'));
        var b64 = await mf.async('base64');
        var im = await new Promise(function (resolve) {
          var img = new Image();
          img.onload = function () { resolve(img); };
          img.onerror = function () { resolve(null); };
          img.src = 'data:' + mime + ';base64,' + b64;
        });
        mediaImgCache[norm] = im;
        return im;
      }

      var newPages = [];
      for (var si = 0; si < slidePaths.length; si++) {
        var spath = slidePaths[si];
        var sf = zip.file(spath);
        if (!sf) continue;
        var sDoc = parser.parseFromString(await sf.async('string'), 'application/xml');

        // 슬라이드별 rels (이미지 매핑)
        var sRelsPath = spath.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels';
        var sRelMap = {};
        if (zip.file(sRelsPath)) {
          var srDoc = parser.parseFromString(await zip.file(sRelsPath).async('string'), 'application/xml');
          qAll(srDoc, 'Relationship').forEach(function (rel) { sRelMap[xmlAttr(rel, 'Id')] = xmlAttr(rel, 'Target'); });
        }

        // 배경색
        var bgNode = qOne(sDoc, 'bg');
        var bgColor = extractFillColor(bgNode) || '#ffffff';
        var bgImg = null;
        var bgBlip = bgNode ? qOne(bgNode, 'blip') : null;
        if (bgBlip) {
          var brId = bgBlip.getAttribute('r:embed') || bgBlip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');
          if (brId && sRelMap[brId]) bgImg = await getMediaImg(sRelMap[brId]);
        }

        var elements = [];
        var spTree = qOne(sDoc, 'spTree');
        if (spTree) {
          var children = Array.from(spTree.children || []);
          for (var ci = 0; ci < children.length; ci++) {
            var ch = children[ci];
            var ln = ch.localName;
            if (ln === 'sp' || ln === 'pic' || ln === 'cxnSp' || ln === 'graphicFrame') {
              var cNvPr = qOne(ch, 'cNvPr');
              var spid = cNvPr ? xmlAttr(cNvPr, 'id') : ('e_' + ci);
              var xfrm = qOne(ch, 'xfrm');
              var off = xfrm ? qOne(xfrm, 'off') : null;
              var ext = xfrm ? qOne(xfrm, 'ext') : null;
              var x = off ? (parseInt(xmlAttr(off, 'x'), 10) / slideCx) * bw : 60;
              var y = off ? (parseInt(xmlAttr(off, 'y'), 10) / slideCy) * bh : 60;
              var w = ext ? (parseInt(xmlAttr(ext, 'cx'), 10) / slideCx) * bw : 240;
              var h = ext ? (parseInt(xmlAttr(ext, 'cy'), 10) / slideCy) * bh : 80;

              if (ln === 'pic') {
                var blip = qOne(ch, 'blip');
                var rid = blip ? (blip.getAttribute('r:embed') || blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed')) : '';
                var imgObj = (rid && sRelMap[rid]) ? await getMediaImg(sRelMap[rid]) : null;
                if (imgObj) elements.push({ kind: 'pic', spid: spid, x: x, y: y, w: w, h: h, img: imgObj });
              } else if (ln === 'graphicFrame') {
                var tbl = qOne(ch, 'tbl');
                if (tbl) {
                  var rows = qAll(tbl, 'tr').map(function (tr) {
                    return qAll(tr, 'tc').map(function (tc) {
                      return qAll(tc, 't').map(function (t) { return t.textContent; }).join('');
                    });
                  });
                  elements.push({ kind: 'table', spid: spid, x: x, y: y, w: w, h: h, rows: rows });
                }
              } else {
                var spPr = qOne(ch, 'spPr');
                var solidFill = spPr ? qOne(spPr, 'solidFill') : null;
                var noFill = spPr ? qOne(spPr, 'noFill') : null;
                var fill = noFill ? null : extractFillColor(solidFill);
                var lnNode = spPr ? qOne(spPr, 'ln') : null;
                var stroke = lnNode ? extractFillColor(lnNode) : null;
                var prstGeom = spPr ? qOne(spPr, 'prstGeom') : null;
                var geom = prstGeom ? xmlAttr(prstGeom, 'prst') : 'rect';

                var paras = [];
                var txBody = qOne(ch, 'txBody');
                if (txBody) {
                  qAll(txBody, 'p').forEach(function (pNode) {
                    var pPr = qOne(pNode, 'pPr');
                    var align = pPr ? (xmlAttr(pPr, 'algn') || 'l') : 'l';
                    var runs = qAll(pNode, 'r');
                    var lineTxt = '';
                    var fontSize = 22;
                    var fontColor = '#0f172a';
                    var bold = false;
                    runs.forEach(function (rNode) {
                      var rPr = qOne(rNode, 'rPr');
                      if (rPr) {
                        var sz = parseInt(xmlAttr(rPr, 'sz'), 10);
                        if (sz > 0) fontSize = Math.max(12, Math.min(72, Math.round((sz / 100) * (bw / 720))));
                        if (xmlAttr(rPr, 'b') === '1') bold = true;
                        var rc = extractFillColor(qOne(rPr, 'solidFill'));
                        if (rc) fontColor = rc;
                      }
                      var tNode = qOne(rNode, 't');
                      if (tNode) lineTxt += tNode.textContent;
                    });
                    if (lineTxt) paras.push({ text: lineTxt, fontSize: fontSize, color: fontColor, bold: bold, align: align });
                  });
                }
                if (fill || stroke || paras.length) {
                  elements.push({ kind: 'shape', spid: spid, geom: geom, x: x, y: y, w: w, h: h, fill: fill, stroke: stroke, paras: paras });
                }
              }
            }
          }
        }

        var steps = parseSlideTiming(sDoc);
        newPages.push({
          kind: 'pptx',
          label: file.name,
          width: bw,
          height: bh,
          pptxSlide: { bgColor: bgColor, bgImg: bgImg, elements: elements, steps: steps },
          pptxStep: 0,
          objects: [],
          redoStack: []
        });
      }

      if (!newPages.length) { showToast('PPTX에서 슬라이드를 찾지 못했습니다.'); return; }
      state.pages = newPages;
      state.pageIndex = 0;
      resetZoom();
      resizeBoardToPage();
      await renderBackground();
      renderInk();
      var totalActs = newPages.reduce(function (a, p) { return a + (p.pptxSlide.steps.length || 0); }, 0);
      showToast(file.name + ' (' + newPages.length + '슬라이드 · 클릭 액션 ' + totalActs + '개)을 불러왔습니다.');
    } catch (err) {
      showToast('PPTX 열기 오류: ' + (err.message || err));
    }
  }

  function computePptxVisibility(slide, stepIdx) {
    // 어떤 spid가 entr(나타나기) 대상이면 초기 숨김 -> 해당 step 도달 시 표시!
    var entrMap = {};
    var exitMap = {};
    var emphSet = {};
    (slide.steps || []).forEach(function (st, idx) {
      var stepNum = idx + 1;
      (st.targets || []).forEach(function (tg) {
        if (tg.presetClass === 'entr' && !entrMap[tg.spid]) entrMap[tg.spid] = stepNum;
        if (tg.presetClass === 'exit' && !exitMap[tg.spid]) exitMap[tg.spid] = stepNum;
        if (tg.presetClass === 'emph' && stepNum === stepIdx) emphSet[tg.spid] = true;
      });
    });
    return { entrMap: entrMap, exitMap: exitMap, emphSet: emphSet };
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
      var reqStep = vis.entrMap[elm.spid] || 0;
      if (reqStep > 0 && stepIdx < reqStep) return; // 아직 클릭 전이므로 숨김!
      var hideStep = vis.exitMap[elm.spid] || 0;
      if (hideStep > 0 && stepIdx >= hideStep) return; // 사라지기 완료!

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
            ctx.fillText(String(elm.rows[r][c] || '').slice(0, 30), cx + 6, cy + cellH / 2);
          }
        }
      } else if (elm.kind === 'shape') {
        if (elm.fill) {
          ctx.fillStyle = elm.fill;
          if (elm.geom === 'ellipse') {
            ctx.beginPath();
            ctx.ellipse(elm.x + elm.w / 2, elm.y + elm.h / 2, elm.w / 2, elm.h / 2, 0, 0, Math.PI * 2);
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
            ctx.ellipse(elm.x + elm.w / 2, elm.y + elm.h / 2, elm.w / 2, elm.h / 2, 0, 0, Math.PI * 2);
            ctx.stroke();
          } else {
            ctx.strokeRect(elm.x, elm.y, elm.w, elm.h);
          }
        }
        if (elm.paras && elm.paras.length) {
          var py = elm.y + 12;
          var padX = 12;
          var maxW = Math.max(40, elm.w - padX * 2);
          elm.paras.forEach(function (p) {
            ctx.fillStyle = p.color || '#0f172a';
            ctx.font = (p.bold ? 'bold ' : '') + p.fontSize + 'px Pretendard, "Noto Sans KR", "Malgun Gothic", sans-serif';
            ctx.textBaseline = 'top';
            ctx.textAlign = p.align === 'ctr' ? 'center' : (p.align === 'r' ? 'right' : 'left');
            var tx = p.align === 'ctr' ? (elm.x + elm.w / 2) : (p.align === 'r' ? (elm.x + elm.w - padX) : (elm.x + padX));
            var wrapped = wrapCanvasText(ctx, p.text, maxW);
            wrapped.forEach(function (ln) {
              ctx.fillText(ln, tx, py);
              py += p.fontSize * 1.35;
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
  // 이미지 불러오기 & 붙여넣기(Ctrl+V)
  // ===================================================================
  function loadImageFromDataUrl(dataUrl, label) {
    var img = new Image();
    img.onload = function () {
      var bw = 1280;
      var bh = Math.round(bw * (img.naturalHeight / Math.max(1, img.naturalWidth)));
      bh = Math.max(480, Math.min(1800, bh));
      state.pages.push({
        kind: 'image',
        label: label || '이미지',
        width: bw,
        height: bh,
        imageObj: img,
        objects: [],
        redoStack: []
      });
      state.pageIndex = state.pages.length - 1;
      resetZoom();
      resizeBoardToPage();
      renderBackground();
      renderInk();
      showToast((label || '이미지') + '를 새 페이지로 불러왔습니다.');
    };
    img.src = dataUrl;
  }

  // ===================================================================
  // 페이지 이동 & 하단 UI 업데이트
  // ===================================================================
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
    if (el('sbPageIndicator')) {
      el('sbPageIndicator').textContent = (state.pageIndex + 1) + ' / ' + state.pages.length + ' (' + (pg.label || '보드').slice(0, 14) + ')';
    }
    var pptxBar = el('sbPptxActionBar');
    if (pptxBar) {
      if (pg.kind === 'pptx' && pg.pptxSlide) {
        pptxBar.style.display = 'flex';
        var tot = (pg.pptxSlide.steps || []).length;
        if (el('sbPptxActIndicator')) el('sbPptxActIndicator').textContent = (pg.pptxStep || 0) + ' / ' + tot;
      } else {
        pptxBar.style.display = 'none';
      }
    }
  }

  // ===================================================================
  // 무료 OCR 손글씨 → 글씨 변환 (Google Gemini 무료 API / OCR.space)
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
    // 손글씨 OCR 정확도 향상을 위해 흰 배경 위에 선택 판서를 진한 검정으로 렌더링
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
    if (engine === 'gemini' && !cfg.geminiKey) {
      el('sbOcrCfgModal').classList.add('show');
      showToast('먼저 무료 OCR 설정에서 Gemini API 키를 입력하거나 OCR.space로 선택해 주세요.');
      return;
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
      el('sbOcrResultText').value = recognized || '(인식된 글자가 없습니다. 직접 입력해 수정할 수 있습니다.)';
    } catch (err) {
      el('sbOcrResultText').value = '변환 실패: ' + (err.message || err);
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
  // 이벤트 바인딩 & SaniZ 호환 단축키
  // ===================================================================
  function init() {
    state.pages = [createBlankPage('dark')];
    resizeBoardToPage();
    renderBackground();
    renderInk();

    window.addEventListener('resize', fitViewport);

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

    // S펜 전용 토글
    if (el('sbPalmRejectBtn')) {
      el('sbPalmRejectBtn').addEventListener('click', function () {
        state.palmReject = !state.palmReject;
        this.classList.toggle('active', state.palmReject);
        this.textContent = state.palmReject ? 'S펜 전용: 켜짐' : 'S펜 전용: 꺼짐';
        showToast(state.palmReject ? 'S펜 전용 모드 ON (손바닥 무시 · 손가락으로 확대 이동)' : 'S펜 전용 모드 OFF (손가락·펜 모두 필기)');
      });
    }

    // 파일 입력
    if (el('sbPdfInput')) el('sbPdfInput').addEventListener('change', function () { if (this.files && this.files[0]) loadPdfFile(this.files[0]); this.value = ''; });
    if (el('sbPptxInput')) el('sbPptxInput').addEventListener('change', function () { if (this.files && this.files[0]) loadPptxFile(this.files[0]); this.value = ''; });
    if (el('sbImgInput')) el('sbImgInput').addEventListener('change', function () {
      var f = this.files && this.files[0];
      if (!f) return;
      var fr = new FileReader();
      fr.onload = function () { loadImageFromDataUrl(fr.result, f.name); };
      fr.readAsDataURL(f);
      this.value = '';
    });

    // Ctrl+V 클립보드 이미지 붙여넣기
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

    // 확대 복귀 & 인라인 텍스트
    if (el('sbZoomResetBtn')) el('sbZoomResetBtn').addEventListener('click', function () { resetZoom(); if (state.mode === 'zoom') setMode('draw'); });
    inlineTextEl.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); commitInlineText(); }
      else if (e.key === 'Escape') { inlineTextEl.style.display = 'none'; pendingTextPos = null; }
    });

    // 편집 모드 액션 바 버튼들
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

    // 페이지 & PPTX 액션 버튼
    if (el('sbPrevPageBtn')) el('sbPrevPageBtn').addEventListener('click', function () { goPage(state.pageIndex - 1); });
    if (el('sbNextPageBtn')) el('sbNextPageBtn').addEventListener('click', function () { goPage(state.pageIndex + 1); });
    if (el('sbAddPageBtn')) el('sbAddPageBtn').addEventListener('click', function () {
      state.pages.push(createBlankPage(state.bgMode));
      goPage(state.pages.length - 1);
    });
    if (el('sbPptxPrevActBtn')) el('sbPptxPrevActBtn').addEventListener('click', function () { stepPptxAction(-1); });
    if (el('sbPptxNextActBtn')) el('sbPptxNextActBtn').addEventListener('click', function () { stepPptxAction(1); });

    if (el('sbStampResetBtn')) el('sbStampResetBtn').addEventListener('click', function () {
      state.stampCounter = 1;
      if (el('sbStampNum')) el('sbStampNum').textContent = '1';
      showToast('숫자 스탬프를 1번부터 다시 시작합니다.');
    });
    if (el('sbClearPinsBtn')) el('sbClearPinsBtn').addEventListener('click', function () {
      pinLayer.innerHTML = '';
      showToast('모든 게시 조각을 닫았습니다.');
    });

    // 저장 / 도움말 / 전체화면
    if (el('sbSavePngBtn')) el('sbSavePngBtn').addEventListener('click', saveCurrentPng);
    if (el('sbSavePdfBtn')) el('sbSavePdfBtn').addEventListener('click', saveMergedPdf);
    if (el('sbHelpBtn')) el('sbHelpBtn').addEventListener('click', function () { el('sbHelpModal').classList.add('show'); });
    if (el('sbFullscreenBtn')) el('sbFullscreenBtn').addEventListener('click', function () {
      if (!document.fullscreenElement) {
        if (document.documentElement.requestFullscreen) document.documentElement.requestFullscreen().catch(function () {});
      } else if (document.exitFullscreen) {
        document.exitFullscreen().catch(function () {});
      }
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

    // SaniZ 호환 전역 단축키
    document.addEventListener('keydown', function (e) {
      var inInput = document.activeElement && (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'TEXTAREA' || document.activeElement.tagName === 'SELECT');
      if (inInput) return;

      if (e.altKey && (e.code === 'Backquote' || e.key === '`')) { e.preventDefault(); setMode(state.mode === 'pin' ? 'draw' : 'pin'); return; }
      if (e.altKey && e.key === '1') { e.preventDefault(); setMode(state.mode === 'zoom' ? 'draw' : 'zoom'); return; }
      if (e.altKey && e.key === '2') { e.preventDefault(); setMode('draw'); return; }
      if (e.altKey && e.key === '3') { e.preventDefault(); setMode(state.mode === 'edit' ? 'draw' : 'edit'); return; }
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
        if (k === 'r') setColor('#ef4444');
        else if (k === 'g') setColor('#22c55e');
        else if (k === 'b') setColor('#3b82f6');
        else if (k === 'y') setColor('#facc15');
        else if (k === 'i') setColor('#f8fafc');
        else if (k === 't') setTool(state.tool === 'text' ? 'pen' : 'text');
        else if (k === 's') setTool(state.tool === 'stamp' ? 'pen' : 'stamp');
        else if (k === 'e') setTool(state.tool === 'eraser' ? 'pen' : 'eraser');
        else if (k === 'k') setBlankBgMode('dark');
        else if (k === 'w') setBlankBgMode('light');
        else if (e.key === ' ' || e.key === 'ArrowRight') {
          if (curPage().kind === 'pptx') { e.preventDefault(); stepPptxAction(1); }
          else if (e.key === 'ArrowRight') goPage(state.pageIndex + 1);
        } else if (e.key === 'ArrowLeft') {
          if (curPage().kind === 'pptx') { e.preventDefault(); stepPptxAction(-1); }
          else goPage(state.pageIndex - 1);
        }
      }
    });
  }

  init();
})();
