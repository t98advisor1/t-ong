/* 스마트 판서 매뉴얼 영상 자동 녹화 도우미
 * 화면에 가짜 커서·클릭 표시·자막·단축키 배지를 띄워, 헤드리스 녹화에서도 동작이 보이게 합니다.
 */
const OVERLAY = () => {
  const css = `
    #demoCursor{position:fixed;left:0;top:0;width:22px;height:22px;margin:-3px 0 0 -3px;z-index:2147483647;pointer-events:none;transition:transform .05s linear;
      background:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='22' height='22' viewBox='0 0 24 24'%3E%3Cpath d='M3 2l16 9-7 2-3 7z' fill='white' stroke='black' stroke-width='1.5' stroke-linejoin='round'/%3E%3C/svg%3E") no-repeat}
    .demoRipple{position:fixed;width:44px;height:44px;margin:-22px 0 0 -22px;border-radius:50%;border:3px solid #facc15;z-index:2147483646;pointer-events:none;animation:demoRip .55s ease-out forwards}
    @keyframes demoRip{from{transform:scale(.3);opacity:1}to{transform:scale(1.4);opacity:0}}
    #demoCaption{position:fixed;left:50%;bottom:34px;transform:translateX(-50%);max-width:88%;z-index:2147483645;pointer-events:none;
      background:rgba(2,6,23,.86);color:#fff;font:700 24px/1.45 'Pretendard','Noto Sans KR',sans-serif;padding:12px 26px;border-radius:14px;
      text-align:center;box-shadow:0 6px 24px rgba(0,0,0,.4);opacity:0;transition:opacity .25s}
    #demoCaption.show{opacity:1}
    #demoCaption b{color:#facc15}
    #demoKey{position:fixed;right:28px;bottom:110px;z-index:2147483645;pointer-events:none;background:#facc15;color:#111827;
      font:800 26px/1 ui-monospace,Consolas,monospace;padding:12px 18px;border-radius:12px;box-shadow:0 6px 20px rgba(0,0,0,.35);opacity:0;transition:opacity .2s}
    #demoKey.show{opacity:1}
    .demoHi{position:fixed;z-index:2147483644;pointer-events:none;border:3px solid #facc15;border-radius:10px;box-shadow:0 0 0 9999px rgba(2,6,23,.45);transition:opacity .25s}
    #demoChapter{position:fixed;left:24px;top:50%;transform:translateY(-50%);z-index:2147483645;pointer-events:none;opacity:0;transition:opacity .3s}
    #demoChapter.show{opacity:1}`;
  const add = () => {
    if (document.getElementById('demoCursor')) return;
    const st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);
    for (const id of ['demoCursor', 'demoCaption', 'demoKey']) { const d = document.createElement('div'); d.id = id; document.body.appendChild(d); }
    const cur = document.getElementById('demoCursor');
    const move = e => { cur.style.transform = `translate(${e.clientX}px,${e.clientY}px)`; };
    window.addEventListener('pointermove', move, true); window.addEventListener('mousemove', move, true);
    window.addEventListener('pointerdown', e => { const r = document.createElement('div'); r.className = 'demoRipple'; r.style.left = e.clientX + 'px'; r.style.top = e.clientY + 'px'; document.body.appendChild(r); setTimeout(() => r.remove(), 600); }, true);
  };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', add); else add();
};

function demo(page) {
  let mx = 640, my = 400;
  const api = {
    async caption(html, ms = 2200) {
      await page.evaluate(h => { const c = document.getElementById('demoCaption'); c.innerHTML = h; c.classList.toggle('show', !!h); }, html);
      if (ms) await page.waitForTimeout(ms);
    },
    // 자막 위치: 'bottom'(기본) 또는 'top' — 화면 아래쪽 버튼을 보여 줄 때 씁니다.
    async captionAt(pos) { await page.evaluate(p => { const c = document.getElementById('demoCaption'); c.style.bottom = p === 'top' ? 'auto' : ''; c.style.top = p === 'top' ? '96px' : ''; }, pos); },
    async hideCaption() { await page.evaluate(() => document.getElementById('demoCaption').classList.remove('show')); },
    // 화면의 한 부분을 노란 테두리로 강조하고 나머지는 살짝 어둡게 합니다.
    async highlight(sel, ms = 1800, pad = 4) {
      const h = await page.waitForSelector(sel, { state: 'visible', timeout: 8000 }); const b = await h.boundingBox();
      await page.evaluate(([x, y, w, hh]) => { const d = document.createElement('div'); d.className = 'demoHi'; Object.assign(d.style, { left: x + 'px', top: y + 'px', width: w + 'px', height: hh + 'px' }); document.body.appendChild(d); }, [b.x - pad, b.y - pad, b.width + pad * 2, b.height + pad * 2]);
      await page.waitForTimeout(ms);
      await page.evaluate(() => document.querySelectorAll('.demoHi').forEach(e => e.remove()));
    },
    async highlightBox(x, y, w, hh, ms = 1800) {
      await page.evaluate(([x, y, w, hh]) => { const d = document.createElement('div'); d.className = 'demoHi'; Object.assign(d.style, { left: x + 'px', top: y + 'px', width: w + 'px', height: hh + 'px' }); document.body.appendChild(d); }, [x, y, w, hh]);
      await page.waitForTimeout(ms);
      await page.evaluate(() => document.querySelectorAll('.demoHi').forEach(e => e.remove()));
    },
    async keyBadge(label, ms = 900) {
      await page.evaluate(l => { const k = document.getElementById('demoKey'); k.textContent = l; k.classList.add('show'); }, label);
      await page.waitForTimeout(ms);
      await page.evaluate(() => document.getElementById('demoKey').classList.remove('show'));
    },
    async press(key, label) { await api.keyBadgeAsync(label || key); await page.keyboard.press(key); await page.waitForTimeout(700); },
    async keyBadgeAsync(label) { await page.evaluate(l => { const k = document.getElementById('demoKey'); k.textContent = l; k.classList.add('show'); setTimeout(() => k.classList.remove('show'), 1100); }, label); },
    async moveTo(x, y, steps = 22) { await page.mouse.move(x, y, { steps }); mx = x; my = y; },
    async center(sel) {
      const h = await page.waitForSelector(sel, { state: 'visible', timeout: 8000 });
      const b = await h.boundingBox(); return [b.x + b.width / 2, b.y + b.height / 2];
    },
    async click(sel, pause = 650) {
      const [x, y] = await api.center(sel);
      await api.moveTo(x, y); await page.waitForTimeout(180);
      await page.mouse.down(); await page.waitForTimeout(70); await page.mouse.up();
      await page.waitForTimeout(pause);
    },
    async tapAt(x, y, pause = 450) { await api.moveTo(x, y, 14); await page.mouse.down(); await page.waitForTimeout(60); await page.mouse.up(); await page.waitForTimeout(pause); },
    // 점 목록을 따라 손으로 그리듯 천천히 긋습니다.
    async stroke(points, stepPx = 6) {
      await api.moveTo(points[0][0], points[0][1], 12);
      await page.mouse.down();
      for (let i = 1; i < points.length; i++) {
        const [x0, y0] = points[i - 1], [x1, y1] = points[i];
        const n = Math.max(2, Math.round(Math.hypot(x1 - x0, y1 - y0) / stepPx));
        for (let s = 1; s <= n; s++) { await page.mouse.move(x0 + (x1 - x0) * s / n, y0 + (y1 - y0) * s / n); await page.waitForTimeout(8); }
      }
      await page.mouse.up(); mx = points[points.length - 1][0]; my = points[points.length - 1][1];
      await page.waitForTimeout(250);
    },
    async drag(x0, y0, x1, y1) { await api.stroke([[x0, y0], [x1, y1]], 8); },
    // 펜 굵기(텍스트 크기도 함께 바뀜)를 화면 조작 없이 바꿉니다.
    async setSize(pt) { await page.evaluate(v => { const b = document.querySelector('[data-size-preset="' + v + '"]'); if (b) b.click(); document.querySelectorAll('.sb-tool-popover.show,.sb-tool-popover.open').forEach(e => e.classList.remove('show', 'open')); }, String(pt)); await page.waitForTimeout(100); },
    // 텍스트 도구로 글자를 넣고 펜으로 돌아옵니다.
    async typeText(x, y, text, pt = 8) {
      await api.setSize(pt);
      await page.keyboard.press('t');
      await api.tapAt(x, y, 350);
      await page.keyboard.type(text, { delay: 80 });
      await page.keyboard.press('Enter');
      await page.waitForTimeout(300);
      await page.keyboard.press('p');
      await api.setSize(5);
    },
    wait: ms => page.waitForTimeout(ms),
  };
  return api;
}

// 글자·도형을 그리는 점 목록 도우미
const shapes = {
  circle(cx, cy, r, n = 48) { const p = []; for (let i = 0; i <= n; i++) { const a = -Math.PI / 2 + i / n * Math.PI * 2; p.push([cx + r * Math.cos(a), cy + r * Math.sin(a)]); } return p; },
  wave(x, y, w, h, n = 40) { const p = []; for (let i = 0; i <= n; i++) p.push([x + w * i / n, y - h * Math.sin(i / n * Math.PI * 2)]); return p; },
  parabola(cx, cy, w, h, n = 40) { const p = []; for (let i = 0; i <= n; i++) { const t = -1 + 2 * i / n; p.push([cx + t * w, cy - h * (1 - t * t)]); } return p; },
};

module.exports = { OVERLAY, demo, shapes };
