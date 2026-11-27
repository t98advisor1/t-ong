/* 스마트 판서 매뉴얼 영상 녹화 (음성 해설 포함)
 * 사용: node tools/sb-manual-video/record.js <서버주소> <출력폴더> [장 이름...]
 *   환경변수 TTS_PY(파이썬), TTS_MODEL(음성 모델 폴더), TTS_CACHE(음성 캐시 폴더)를 주면 음성을 넣습니다. (tts.py 참고)
 * 장마다 1280×720 영상을 녹화한 뒤 ffmpeg로 mp4(H.264 + AAC 음성)와 포스터(jpg)를 만듭니다.
 */
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { OVERLAY, demo } = require('./demo-lib');
const CHAPTERS = require('./chapters');

const BASE = process.argv[2] || 'http://localhost:4321';
const OUT = process.argv[3] || 'assets/smartboard_guide/video';
const ONLY = process.argv.slice(4);
const W = 1280, H = 720;
const TTS_PY = process.env.TTS_PY, TTS_MODEL = process.env.TTS_MODEL, TTS_CACHE = process.env.TTS_CACHE;
const md5 = t => require('crypto').createHash('md5').update(t).digest('hex');

// chapters.js 에서 say('자막', '음성') · titleCard(…, '음성') 의 음성 문장을 모아 미리 만들어 둡니다.
function speechTexts() {
  const src = fs.readFileSync(path.join(__dirname, 'chapters.js'), 'utf8');
  const out = new Set();
  const str = "'((?:[^'\\\\]|\\\\.)*)'";
  const reSay = new RegExp('\\.say\\(\\s*' + str + '\\s*,\\s*' + str, 'g');
  const reTitle = new RegExp('\\.titleCard\\(\\s*' + str + '\\s*,\\s*' + str + '\\s*,\\s*' + str + '\\s*,\\s*' + str, 'g');
  let m;
  while ((m = reSay.exec(src))) if (m[2]) out.add(m[2].replace(/\\'/g, "'"));
  while ((m = reTitle.exec(src))) if (m[4]) out.add(m[4].replace(/\\'/g, "'"));
  return [...out];
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  let durations = null;
  if (TTS_PY && TTS_MODEL && TTS_CACHE) {
    const list = path.join(os.tmpdir(), 'sb-tts-list.json');
    fs.writeFileSync(list, JSON.stringify(speechTexts()));
    execFileSync(TTS_PY, [path.join(__dirname, 'tts.py'), TTS_MODEL, TTS_CACHE, list], { stdio: 'inherit' });
    durations = JSON.parse(fs.readFileSync(path.join(TTS_CACHE, 'durations.json'), 'utf8'));
  }
  const browser = await chromium.launch();
  for (const ch of CHAPTERS) {
    if (ONLY.length && ONLY.indexOf(ch.id) < 0) continue;
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'sbrec-'));
    const ctx = await browser.newContext(Object.assign({ viewport: { width: W, height: H }, recordVideo: { dir: tmp, size: { width: W, height: H } } }, ch.contextOptions || {}));
    const t0 = Date.now();
    await ctx.addInitScript(OVERLAY);
    // 영상마다 같은 출발점: 자동 번호 스탬프는 꺼 두고(해당 장에서 직접 켭니다), S펜 자동은 기본값.
    await ctx.addInitScript(() => { try { if (!sessionStorage.getItem('sbRecInit')) { localStorage.setItem('sbAutoStampShape_v1', 'off'); sessionStorage.setItem('sbRecInit', '1'); } } catch (e) {} });
    if (ch.routes) await ch.routes(ctx);
    const page = await ctx.newPage();
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));
    await page.goto(BASE + '/smart-board.html', { waitUntil: 'load' });
    await page.waitForTimeout(1200);
    // 영상에서 잘 보이도록 펜을 조금 굵게 시작합니다(녹화 전, 잘라내는 구간).
    await page.evaluate(() => { const b = document.querySelector('[data-size-preset="5"]'); if (b) b.click(); document.querySelectorAll('.sb-tool-popover').forEach(e => e.classList.remove('show', 'open')); });
    if (ch.setup) await ch.setup(page);
    const startSec = (Date.now() - t0) / 1000;      // 로딩 화면은 잘라냅니다
    const narr = durations ? { t0, durations, list: [] } : null;
    const d = demo(page, narr);
    await d.moveTo(W / 2, H / 2, 1);
    await ch.run(page, d);
    await d.done(900);
    const video = page.video();
    await ctx.close();
    const webm = await video.path();
    const mp4 = path.join(OUT, ch.id + '.mp4'), poster = path.join(OUT, ch.id + '.jpg');
    const silent = path.join(tmp, 'v.mp4');
    execFileSync('ffmpeg', ['-y', '-loglevel', 'error', '-ss', startSec.toFixed(2), '-i', webm,
      '-c:v', 'libx264', '-preset', 'slow', '-crf', '30', '-pix_fmt', 'yuv420p', '-an', silent]);
    if (narr && narr.list.length) {
      // 말한 시각에 맞춰 음성 조각을 겹쳐 붙입니다.
      const args = ['-y', '-loglevel', 'error', '-i', silent];
      const parts = [];
      narr.list.forEach((n, i) => {
        args.push('-i', path.join(TTS_CACHE, n.h + '.wav'));
        const ms = Math.max(0, Math.round(n.t - startSec * 1000));
        parts.push(`[${i + 1}:a]aresample=44100,adelay=${ms}:all=1[a${i}]`);
      });
      const mix = parts.join(';') + ';' + narr.list.map((_, i) => `[a${i}]`).join('') + `amix=inputs=${narr.list.length}:normalize=0,volume=1.6[aout]`;
      args.push('-filter_complex', mix, '-map', '0:v', '-map', '[aout]', '-c:v', 'copy', '-c:a', 'aac', '-b:a', '96k', '-ac', '1', '-movflags', '+faststart', mp4);
      execFileSync('ffmpeg', args);
    } else {
      execFileSync('ffmpeg', ['-y', '-loglevel', 'error', '-i', silent, '-c', 'copy', '-movflags', '+faststart', mp4]);
    }
    const dur = execFileSync('ffprobe', ['-v', 'error', '-show_entries', 'format=duration', '-of', 'csv=p=0', mp4]).toString().trim();
    // 포스터는 판서가 어느 정도 쌓인 뒤 장면(기본: 영상의 3/4 지점)에서 뽑습니다.
    const posterSec = ch.posterAt != null ? ch.posterAt : (+dur) * 0.75;
    execFileSync('ffmpeg', ['-y', '-loglevel', 'error', '-ss', posterSec.toFixed(2), '-i', mp4, '-frames:v', '1', '-q:v', '4', poster]);
    const kb = Math.round(fs.statSync(mp4).size / 1024);
    console.log(`${ch.id}: ${(+dur).toFixed(1)}s, ${kb}KB, 음성 ${narr ? narr.list.length : 0}개${errors.length ? ' errors: ' + errors.join(' | ') : ''}`);
    fs.rmSync(tmp, { recursive: true, force: true });
  }
  await browser.close();
})();
