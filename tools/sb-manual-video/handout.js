/* 단축키 핸드아웃 그림(PNG)과 A4 PDF 만들기
 * 사용: node tools/sb-manual-video/handout.js [출력폴더=assets/smartboard_guide]
 */
const { chromium } = require('playwright');
const path = require('path');
const OUT = process.argv[2] || 'assets/smartboard_guide';
(async () => {
  const b = await chromium.launch();
  // A4 가로(297×210mm)를 96dpi 화면 크기로 맞추고 2배로 찍어 선명하게 만듭니다.
  const p = await b.newPage({ viewport: { width: 1123, height: 794 }, deviceScaleFactor: 2 });
  await p.goto('file://' + path.join(__dirname, 'handout.html'));
  await p.waitForTimeout(500);
  await p.screenshot({ path: path.join(OUT, 'shortcut-handout.png'), clip: { x: 0, y: 0, width: 1123, height: 794 } });
  await p.pdf({ path: path.join(OUT, 'shortcut-handout.pdf'), width: '297mm', height: '210mm', printBackground: true, pageRanges: '1' });
  await b.close();
  console.log('핸드아웃 PNG · PDF 완료');
})();
