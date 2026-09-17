/* 배포용 빌드
 *  - app.js 를 난독화해서 dist/app.js 로
 *  - index.html / style.css 는 주석 제거 + 공백 압축
 *  - 나머지 정적 파일은 그대로 복사
 *  - sw.js 의 캐시 버전을 내용 해시로 자동 교체 (캐시가 옛 파일을 물고 있는 문제 방지)
 *
 * 실행: npm run build
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const JavaScriptObfuscator = require('javascript-obfuscator');

const SRC = __dirname;
const OUT = path.join(SRC, 'dist');

const read = (p) => fs.readFileSync(path.join(SRC, p), 'utf8');
const write = (p, s) => {
  const full = path.join(OUT, p);
  fs.mkdirSync(path.dirname(full), { recursive: true });
  fs.writeFileSync(full, s, 'utf8');
};

function copyDir(rel) {
  const from = path.join(SRC, rel);
  if (!fs.existsSync(from)) return 0;
  let n = 0;
  for (const entry of fs.readdirSync(from, { withFileTypes: true })) {
    const r = path.join(rel, entry.name);
    if (entry.isDirectory()) { n += copyDir(r); continue; }
    const dest = path.join(OUT, r);
    fs.mkdirSync(path.dirname(dest), { recursive: true });
    fs.copyFileSync(path.join(SRC, r), dest);
    n++;
  }
  return n;
}

// HTML: 주석 제거 + 줄머리 들여쓰기 제거 (태그 내부 텍스트는 건드리지 않음)
function minifyHtml(html) {
  return html
    .replace(/<!--(?!\[if)[\s\S]*?-->/g, '')
    .split('\n')
    .map((l) => l.replace(/^\s+/, ''))
    .filter((l) => l.length)
    .join('\n');
}

// CSS: 주석 제거 + 공백 압축
function minifyCss(css) {
  return css
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/\s*([{}:;,>~+])\s*/g, '$1')
    .replace(/;}/g, '}')
    .replace(/\s+/g, ' ')
    .trim();
}

console.log('▶ 빌드 시작');
fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(OUT, { recursive: true });

// 1) app.js 난독화
const appSrc = read('app.js');
const obfuscated = JavaScriptObfuscator.obfuscate(appSrc, {
  compact: true,
  controlFlowFlattening: true,
  controlFlowFlatteningThreshold: 0.4,
  deadCodeInjection: true,
  deadCodeInjectionThreshold: 0.2,
  identifierNamesGenerator: 'hexadecimal',
  numbersToExpressions: true,
  simplify: true,
  splitStrings: true,
  splitStringsChunkLength: 8,
  stringArray: true,
  stringArrayEncoding: ['base64'],
  stringArrayRotate: true,
  stringArrayShuffle: true,
  stringArrayThreshold: 0.9,
  transformObjectKeys: true,
  unicodeEscapeSequence: false,
  // 개발자 도구를 방해하는 옵션(debugProtection 등)은 쓰지 않습니다.
  // 정상 사용자에게 피해가 가고, 보호 효과도 크지 않습니다.
  debugProtection: false,
  selfDefending: false,
  disableConsoleOutput: false
}).getObfuscatedCode();
write('app.js', obfuscated);

// 2) CSS / HTML
const css = minifyCss(read('style.css'));
write('style.css', css);

const html = minifyHtml(read('index.html'));
write('index.html', html);

// 3) 캐시 버전을 내용 해시로 교체
const hash = crypto.createHash('sha1')
  .update(obfuscated).update(css).update(html)
  .digest('hex').slice(0, 10);
const sw = read('sw.js').replace(
  /var CACHE_VERSION = '[^']*';/,
  "var CACHE_VERSION = '" + hash + "';"
);
write('sw.js', sw);

// 4) 정적 파일
['manifest.webmanifest', 'icon.svg', 'icon-maskable.svg'].forEach((f) => {
  if (fs.existsSync(path.join(SRC, f))) write(f, read(f));
});
const nGuide = copyDir('guide');
const nAssets = copyDir('assets');

const kb = (s) => Math.round(Buffer.byteLength(s, 'utf8') / 1024);
console.log('  app.js    ' + kb(appSrc) + 'KB → ' + kb(obfuscated) + 'KB (난독화)');
console.log('  style.css ' + kb(read('style.css')) + 'KB → ' + kb(css) + 'KB');
console.log('  index.html ' + kb(read('index.html')) + 'KB → ' + kb(html) + 'KB');
console.log('  guide/ ' + nGuide + '개, assets/ ' + nAssets + '개');
console.log('  캐시 버전 ' + hash);
console.log('✔ dist/ 생성 완료');
