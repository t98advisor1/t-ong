// Build a fresh, allowlisted deployment input folder. Never copy the whole workspace.
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const root = path.resolve(__dirname, '..');
const expectedProject = 'prj_UipXxJ6Fns0yLaOQS39WIIdbdHuO';
const project = JSON.parse(fs.readFileSync(path.join(root, '.vercel/project.json'), 'utf8'));
if (project.projectId !== expectedProject || project.projectName !== 't-ong') throw Error('Not the T-ong project; refusing deployment preparation');
const version = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8')).version;
if (!/^\d+\.\d+\.\d+-preview\.\d+$/.test(version)) throw Error('Unexpected app version');
const stage = path.join(root, 'tmp', 'tong-release-' + version + '-' + crypto.randomUUID());
if (!stage.startsWith(path.join(root, 'tmp') + path.sep) || fs.existsSync(stage)) throw Error('Unsafe or reused staging path');
fs.mkdirSync(stage, { recursive:true });
const files = [];
function copyPublic(relative) {
  const source = path.join(root, relative);
  const destination = path.join(stage, relative);
  const stat = fs.lstatSync(source);
  if (stat.isSymbolicLink()) throw Error('Symlinks are not allowed in public release: ' + relative);
  if (stat.isDirectory()) {
    fs.mkdirSync(destination, { recursive:true });
    for (const name of fs.readdirSync(source)) copyPublic(path.join(relative, name));
  } else if (stat.isFile()) {
    if (/(?:^|[\\/])(?:\.env[^\\/]*|[^\\/]+\.(?:keystore|jks|apk|aab|zip))$/i.test(relative)) throw Error('Private file in public source');
    fs.mkdirSync(path.dirname(destination), { recursive:true });
    fs.copyFileSync(source, destination);
    files.push(relative.replace(/\\/g, '/'));
  } else throw Error('Unsupported file: ' + relative);
}
for (const file of ['index.html','style.css','app.js','sw.js','build.js','package.json','vercel.json',
  'manifest.webmanifest','icon.svg','icon-maskable.svg','.vercelignore','.vercel/project.json']) copyPublic(file);
for (const directory of ['assets','guide','api','.well-known']) copyPublic(directory);
if (!files.includes('.well-known/assetlinks.json') || !files.includes('api/usage.js')) throw Error('Required public endpoint missing');
console.log(JSON.stringify({ stage, version, projectId:project.projectId, publicFiles:files.length,
  sourceHash:crypto.createHash('sha256').update(fs.readFileSync(path.join(root,'app.js'))).digest('hex'), files }, null, 2));
