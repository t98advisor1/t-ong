const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const source = path.resolve(__dirname, '..');
const root = process.argv.includes('--built') ? path.join(source, 'dist') : source;
const read = name => fs.readFileSync(path.join(root, name), 'utf8');
const html = read('index.html');
const version = JSON.parse(fs.readFileSync(path.join(source, 'package.json'), 'utf8')).version;
const url = 'https://t98advisor1.github.io/pose-mission-training/';
const menu = html.match(/<div class="more-menu game-menu" id="gameMenu">([\s\S]*?)<\/div>/);
assert.ok(menu, 'Game menu exists');
const anchors = Array.from(menu[1].matchAll(/<a\b([^>]*)>([\s\S]*?)<\/a>/g));
const pose = anchors.filter(([, attrs]) => attrs.includes('href="' + url + '"'));
assert.equal(pose.length, 1, 'Pose game appears exactly once in the game menu');
assert.match(pose[0][1], /target="_blank"/, 'Game opens separately from T-ong');
assert.match(pose[0][1], /rel="noopener noreferrer"/, 'External tab is isolated');
assert.match(pose[0][2], /온몸으로 찰칵! 포즈 챌린지/, 'Actual game title is shown');
assert.ok(anchors.some(([, attrs]) => attrs.includes('https://handcar26.netlify.app/')), 'Existing game is preserved');
assert.equal(anchors.length, 12, 'All twelve game links remain available');
assert.ok(html.includes('/app.js?v=' + version), 'HTML JS version matches package');
assert.ok(html.includes('/style.css?v=' + version), 'HTML CSS version matches package');
const sw = read('sw.js');
assert.ok(sw.includes('/app.js?v=' + version), 'Service worker JS version matches');
assert.ok(sw.includes('/style.css?v=' + version), 'Service worker CSS version matches');
if (root === source) {
  const app = read('app.js');
  assert.ok(app.includes("var APP_VERSION = '" + version + "'"), 'App version matches');
  assert.match(app, /<strong>온몸으로 찰칵! 포즈 챌린지<\/strong>/, 'Usage guide includes game');
  assert.match(app, /게임 메뉴에 온몸으로 찰칵! 포즈 챌린지 추가/, 'Release notes include addition');
}
console.log(JSON.stringify({ status: 'PASS', target: root === source ? 'source' : 'dist', version, gameLinks: anchors.length, gameUrl: url }, null, 2));
