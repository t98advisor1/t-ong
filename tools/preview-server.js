/* Local/LAN preview of public app files only. Do not expose .env or work documents. */
const http = require('http');
const fs = require('fs');
const path = require('path');
const root = path.resolve(__dirname, '..');
const files = new Set(['index.html','app.js','style.css','sw.js','manifest.webmanifest','icon.svg','icon-maskable.svg']);
const mime = {'.html':'text/html; charset=utf-8','.js':'application/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.json':'application/json','.webmanifest':'application/manifest+json','.svg':'image/svg+xml','.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg','.webp':'image/webp','.woff2':'font/woff2'};
http.createServer((req,res) => {
  let rel;
  try { rel = decodeURIComponent(new URL(req.url,'http://localhost').pathname).replace(/^\/+/, '') || 'index.html'; }
  catch { res.writeHead(400).end(); return; }
  const full = path.resolve(root,rel);
  const safe = full.startsWith(root + path.sep) && !rel.split(/[\\/]/).some(p => p.startsWith('.')) &&
    (files.has(rel) || rel.startsWith('assets/') || rel.startsWith('guide/'));
  if (!safe) { res.writeHead(404).end(); return; }
  fs.stat(full,(err,stat) => {
    if(err || !stat.isFile()){res.writeHead(404).end();return;}
    res.writeHead(200,{'Content-Type':mime[path.extname(full)]||'application/octet-stream','Cache-Control':'no-store'});
    fs.createReadStream(full).pipe(res);
  });
}).listen(4321,'0.0.0.0',()=>console.log('T-ong public-file preview listening on 4321'));
