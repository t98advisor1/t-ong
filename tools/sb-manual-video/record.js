/* 스마트 판서 매뉴얼 영상 녹화
 * 사용: node tools/sb-manual-video/record.js <서버주소> <출력폴더> [장 이름...]
 * 장마다 1280×720 영상(webm)을 녹화한 뒤 ffmpeg로 mp4와 포스터(jpg)를 만듭니다.
 */
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const { OVERLAY, demo } = require('./demo-lib');
const CHAPTERS = require('./chapters');

const BASE = process.argv[2] || 'http://localhost:4321';
const OUT = process.argv[3] || 'assets/smartboard_guide/video';
const ONLY = process.argv.slice(4);
const W = 1280, H = 720;

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const browser = await chromium.launch();
  for (const ch of CHAPTERS) {
    if (ONLY.length && ONLY.indexOf(ch.id) < 0) continue;
    const tmp = fs.mkdtempSync(path.join(require('os').tmpdir(), 'sbrec-'));
    const ctx = await browser.newContext({ viewport: { width: W, height: H }, recordVideo: { dir: tmp, size: { width: W, height: H } } });
    const t0 = Date.now();
    await ctx.addInitScript(OVERLAY);
    if (ch.routes) await ch.routes(ctx);
    const page = await ctx.newPage();
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));
    await page.goto(BASE + '/smart-board.html', { waitUntil: 'load' });
    await page.waitForTimeout(1200);
    // 영상에서 잘 보이도록 기본 펜을 조금 굵게 시작합니다(녹화 전, 잘라내는 구간).
    await page.evaluate(() => { const b = document.querySelector('[data-size-preset="5"]'); if (b) b.click(); document.querySelectorAll('.sb-tool-popover').forEach(e => e.classList.remove('show', 'open')); });
    if (ch.setup) await ch.setup(page);
    const startSec = (Date.now() - t0) / 1000;      // 로딩 화면은 잘라냅니다
    const d = demo(page);
    await d.moveTo(W / 2, H / 2, 1);
    await ch.run(page, d);
    await page.waitForTimeout(900);
    const video = page.video();
    await ctx.close();
    const webm = await video.path();
    const mp4 = path.join(OUT, ch.id + '.mp4'), poster = path.join(OUT, ch.id + '.jpg');
    execFileSync('ffmpeg', ['-y', '-loglevel', 'error', '-ss', startSec.toFixed(2), '-i', webm,
      '-c:v', 'libx264', '-preset', 'slow', '-crf', '30', '-pix_fmt', 'yuv420p', '-movflags', '+faststart', '-an', mp4]);
    const dur = execFileSync('ffprobe', ['-v', 'error', '-show_entries', 'format=duration', '-of', 'csv=p=0', mp4]).toString().trim();
    // 포스터는 판서가 어느 정도 쌓인 뒤 장면(기본: 영상의 3/4 지점)에서 뽑습니다.
    const posterSec = ch.posterAt != null ? ch.posterAt : (+dur) * 0.75;
    execFileSync('ffmpeg', ['-y', '-loglevel', 'error', '-ss', posterSec.toFixed(2), '-i', mp4, '-frames:v', '1', '-q:v', '4', poster]);
    const kb = Math.round(fs.statSync(mp4).size / 1024);
    console.log(`${ch.id}: ${(+dur).toFixed(1)}s, ${kb}KB${errors.length ? ' errors: ' + errors.join(' | ') : ''}`);
    fs.rmSync(tmp, { recursive: true, force: true });
  }
  await browser.close();
})();
