/* ===================================================================
 * T-ong AI용 마크다운(.md) 올인원 변환기 (md-converter.js)
 *  - 작업 폴더 `md변환기`의 `hwpx2md.py` (한글 수식 → LaTeX $...$ 변환기) 및
 *    `convert.py` (한컴 PUA 수식글꼴 복원 + MarkItDown 다중 포맷 변환기) 100% 브라우저 이식
 *  - 지원 포맷: .hwpx, .pdf, .docx, .pptx, .xlsx, .xls (+ .hwp 안내)
 * =================================================================== */
(function () {
  'use strict';

  if (window.pdfjsLib) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = './libs/pdf.worker.min.js';
  }

  var OK_EXT = ['.hwpx', '.pdf', '.docx', '.pptx', '.xlsx', '.xls'];

  // -------------------------------------------------------------------
  // 1. convert.py 의 한컴 수식 글꼴(PUA 사용자 정의 영역 문자 0xE000~0xF8FF) 복원 표
  // -------------------------------------------------------------------
  var PUA_MAP = {};
  (function initPuaMap() {
    var i;
    for (i = 0; i < 26; i++) {
      PUA_MAP[String.fromCharCode(0xE000 + i)] = String.fromCharCode(65 + i); // A-Z
      PUA_MAP[String.fromCharCode(0xE0E5 + i)] = String.fromCharCode(97 + i); // a-z
    }
    var digits = '1234567890';
    for (i = 0; i < digits.length; i++) {
      PUA_MAP[String.fromCharCode(0xE034 + i)] = digits[i];
    }
    var greekUpper = 'ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ';
    for (i = 0; i < greekUpper.length; i++) {
      PUA_MAP[String.fromCharCode(0xE085 + i)] = greekUpper[i];
    }
    var greekLower = 'αβγδεζηθικλμνξοπρστυφχψω';
    for (i = 0; i < greekLower.length; i++) {
      PUA_MAP[String.fromCharCode(0xE09D + i)] = greekLower[i];
    }
    var extras = {
      '\ue044': '(', '\ue045': ')', '\ue046': '−', '\ue047': '=', '\ue048': '+',
      '\ue04b': '{', '\ue04c': '}', '\ue04d': '|', '\ue04f': ':', '\ue052': ',',
      '\ue053': '.', '\ue054': '⊄', '\ue055': '<', '\ue056': '>', '\ue05c': '√',
      '\ue06d': '/', '\ue101': '|', '\uf000': '▪',
      '\ue03e': '!', '\ue049': '[', '\ue04a': ']', '\ue05b': '∫', '\ue05f': '˙',
      '\ue063': '^', '\ue065': '⌒', '\ue067': 'Σ', '\ue06c': '←', '\ue06e': '→',
      '\ue0c8': '°', '\ue042': '%',
      '\ue078': '⎧', '\ue079': '⎨', '\ue07a': '⎩', '\ue07b': '⎪'
    };
    Object.keys(extras).forEach(function (k) { PUA_MAP[k] = extras[k]; });
  })();

  var PUA_SEQ = [
    ['\ue054\ue054', '∥'],
    ['\ue06d\ue06e', '→']
  ];

  function restorePua(text) {
    var fixed = 0, unknown = 0;
    var s = String(text || '');
    PUA_SEQ.forEach(function (pair) {
      var k = pair[0], v = pair[1];
      var parts = s.split(k);
      if (parts.length > 1) {
        fixed += (parts.length - 1);
        s = parts.join(v);
      }
    });
    var out = [];
    for (var i = 0; i < s.length; i++) {
      var code = s.charCodeAt(i);
      var ch = s[i];
      if (code >= 0xE000 && code <= 0xF8FF) {
        if (PUA_MAP[ch]) {
          out.push(PUA_MAP[ch]);
          fixed++;
        } else {
          out.push('□');
          unknown++;
        }
      } else {
        out.push(ch);
      }
    }
    return { text: out.join(''), fixed: fixed, unknown: unknown };
  }

  // -------------------------------------------------------------------
  // 2. hwpx2md.py 의 한글(HWPX) 수식 스크립트 → LaTeX($...$) 변환 엔진
  // -------------------------------------------------------------------
  var EQ_KEY = {
    inf: '\\infty', times: '\\times', TIMES: '\\times', cdot: '\\cdot', div: '\\div',
    le: '\\le', leq: '\\le', LEQ: '\\le', ge: '\\ge', geq: '\\ge', GEQ: '\\ge',
    pm: '\\pm', cdots: '\\cdots', CDOTS: '\\cdots', ldots: '\\ldots', LDOTS: '\\ldots',
    prime: "'", sum: '\\sum', SUM: '\\sum', int: '\\int', INT: '\\int', lim: '\\lim', Lim: '\\lim',
    sin: '\\sin', cos: '\\cos', tan: '\\tan', sec: '\\sec', csc: '\\csc', cot: '\\cot',
    log: '\\log', ln: '\\ln', exp: '\\exp', angle: '\\angle', ANGLE: '\\angle',
    triangle: '\\triangle', TRIANGLE: '\\triangle', deg: '^\\circ', DEG: '^\\circ',
    bot: '\\perp', BOT: '\\perp', equiv: '\\equiv', sim: '\\sim', approx: '\\approx',
    rarrow: '\\rightarrow', larrow: '\\leftarrow', iff: '\\iff', therefore: '\\therefore',
    subset: '\\subset', supset: '\\supset', cup: '\\cup', cap: '\\cap', emptyset: '\\emptyset',
    partial: '\\partial', nabla: '\\nabla', neq: '\\ne'
  };
  var GREEK_LIST = 'alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau upsilon phi chi psi omega'.split(' ');
  GREEK_LIST.forEach(function (g) {
    var cap = g.charAt(0).toUpperCase() + g.slice(1);
    EQ_KEY[g] = '\\' + g;
    EQ_KEY[g.toUpperCase()] = '\\' + cap;
    EQ_KEY[cap] = '\\' + cap;
  });
  var EQ_UNARY = {
    sqrt: '\\sqrt', SQRT: '\\sqrt', bar: '\\overline', BAR: '\\overline',
    vec: '\\vec', VEC: '\\vec', hat: '\\hat', dot: '\\dot', tilde: '\\tilde',
    under: '\\underline', dyad: '\\overleftrightarrow'
  };
  var EQ_SKIP = {
    rm: 1, it: 1, bold: 1, RM: 1, IT: 1, BOLD: 1,
    left: 1, right: 1, LEFT: 1, RIGHT: 1,
    big: 1, BIG: 1, bigg: 1, BIGG: 1,
    eqalign: 1, pile: 1, lpile: 1, rpile: 1
  };
  var EQ_SPECIAL = { over: 1, cases: 1, matrix: 1, pmatrix: 1, bmatrix: 1, dmatrix: 1 };
  var EQ_WORDS = Object.keys(EQ_KEY)
    .concat(Object.keys(EQ_UNARY))
    .concat(Object.keys(EQ_SKIP))
    .concat(Object.keys(EQ_SPECIAL))
    .sort(function (a, b) { return b.length - a.length; });

  var EQ_SYM = [
    ['<->', '\\leftrightarrow'],
    ['->', '\\to'],
    ['<=', '\\le'],
    ['>=', '\\ge'],
    ['!=', '\\ne'],
    ['+-', '\\pm'],
    ['-+', '\\mp'],
    ['==', '\\equiv'],
    ['=>', '\\Rightarrow']
  ];

  function splitIdent(w) {
    var out = [];
    while (w.length > 0) {
      var matched = false;
      for (var i = 0; i < EQ_WORDS.length; i++) {
        var k = EQ_WORDS[i];
        if (w.indexOf(k) === 0) {
          out.push(k);
          w = w.slice(k.length);
          matched = true;
          break;
        }
      }
      if (!matched) {
        var j = 1;
        while (j < w.length) {
          var sub = w.slice(j);
          var hit = EQ_WORDS.some(function (k) { return k.length >= 2 && sub.indexOf(k) === 0; });
          if (hit) break;
          j++;
        }
        out.push(w.slice(0, j));
        w = w.slice(j);
      }
    }
    return out;
  }

  function tokenizeEq(s) {
    s = String(s || '').replace(/[`~\r\n]/g, ' ');
    var toks = [];
    var i = 0;
    while (i < s.length) {
      var c = s[i];
      if (/\s/.test(c)) {
        if (toks.length && toks[toks.length - 1] !== ' ') toks.push(' ');
        i++;
        continue;
      }
      if ('{}^_&#'.indexOf(c) >= 0) {
        toks.push(c);
        i++;
        continue;
      }
      var mAlpha = s.slice(i).match(/^[A-Za-z]+/);
      if (mAlpha) {
        toks = toks.concat(splitIdent(mAlpha[0]));
        i += mAlpha[0].length;
        continue;
      }
      var symHit = false;
      for (var si = 0; si < EQ_SYM.length; si++) {
        var sy = EQ_SYM[si][0], rep = EQ_SYM[si][1];
        if (s.slice(i, i + sy.length) === sy) {
          toks.push(rep);
          i += sy.length;
          symHit = true;
          break;
        }
      }
      if (symHit) continue;
      var mNum = s.slice(i).match(/^[0-9.]+/);
      if (mNum) {
        toks.push(mNum[0]);
        i += mNum[0].length;
      } else {
        toks.push(c);
        i++;
      }
    }
    return toks;
  }

  function parseEqTokens(toks, i, depth) {
    i = i || 0;
    depth = depth || 0;
    var seq = [];
    while (i < toks.length) {
      var t = toks[i];
      if (t === '{') {
        var res = parseEqTokens(toks, i + 1, depth + 1);
        seq.push({ group: res.seq });
        i = res.nextIdx;
        continue;
      }
      if (t === '}') {
        if (depth === 0) { i++; continue; }
        return { seq: seq, nextIdx: i + 1 };
      }
      seq.push(t);
      i++;
    }
    return { seq: seq, nextIdx: i };
  }

  var EQ_STOP = { ' ': 1, '&': 1, '#': 1, 'over': 1, ',': 1 };

  function getOperand(seq, i, scripts) {
    if (typeof scripts === 'undefined') scripts = true;
    while (i < seq.length && seq[i] === ' ') i++;
    if (i < seq.length && typeof seq[i] === 'object' && seq[i].group && !scripts) {
      return { op: [seq[i]], nextIdx: i + 1 };
    }
    var j = i, depth = 0;
    while (j < seq.length) {
      var t = seq[j];
      if (typeof t === 'string') {
        if (EQ_STOP[t]) break;
        if (!scripts && (t === '^' || t === '_')) break;
        if (t === '(') depth++;
        else if (t === ')') {
          if (depth === 0) break;
          depth--;
        }
      }
      j++;
    }
    return { op: seq.slice(i, j), nextIdx: j };
  }

  function renderEqSeq(seq) {
    var out = [];
    var i = 0;
    while (i < seq.length) {
      var t = seq[i];
      if (typeof t === 'object' && t.group) {
        out.push(renderEqSeq(t.group));
        i++;
        continue;
      }
      if (t === '^' || t === '_') {
        var r1 = getOperand(seq, i + 1, false);
        out.push(t + '{' + renderEqSeq(r1.op) + '}');
        i = r1.nextIdx;
        continue;
      }
      if (t === 'over') {
        while (out.length && out[out.length - 1] === ' ') out.pop();
        var num = [];
        while (out.length && out[out.length - 1] !== ' ') {
          num.unshift(out.pop());
        }
        var r2 = getOperand(seq, i + 1, true);
        out.push('\\frac{' + num.join('') + '}{' + renderEqSeq(r2.op) + '}');
        i = r2.nextIdx;
        continue;
      }
      if (EQ_UNARY[t]) {
        var r3 = getOperand(seq, i + 1, true);
        var op3 = r3.op;
        if (op3.length === 1 && typeof op3[0] === 'object' && op3[0].group) op3 = op3[0].group;
        out.push(EQ_UNARY[t] + '{' + renderEqSeq(op3) + '}');
        i = r3.nextIdx;
        continue;
      }
      if (t === 'cases' && i + 1 < seq.length && typeof seq[i + 1] === 'object') {
        out.push('\\begin{cases}' + renderEqSeq(seq[i + 1].group) + '\\end{cases}');
        i += 2;
        continue;
      }
      if ((t === 'matrix' || t === 'pmatrix' || t === 'bmatrix' || t === 'dmatrix') && i + 1 < seq.length && typeof seq[i + 1] === 'object') {
        var envMap = { matrix: 'matrix', pmatrix: 'pmatrix', bmatrix: 'bmatrix', dmatrix: 'vmatrix' };
        var env = envMap[t];
        out.push('\\begin{' + env + '}' + renderEqSeq(seq[i + 1].group) + '\\end{' + env + '}');
        i += 2;
        continue;
      }
      if ((t === 'eqalign' || t === 'pile' || t === 'lpile' || t === 'rpile') && i + 1 < seq.length && typeof seq[i + 1] === 'object') {
        var g = seq[i + 1].group.slice();
        while (g.length && (g[0] === '#' || g[0] === ' ')) g.shift();
        out.push(renderEqSeq(g));
        i += 2;
        continue;
      }
      if (EQ_SKIP[t]) { i++; continue; }
      if (t === '#') { out.push(' \\\\ '); i++; continue; }
      if (EQ_KEY[t]) {
        var v = EQ_KEY[t];
        if (v.charAt(0) === '\\' && /^[A-Za-z]+$/.test(v.slice(1)) && i + 1 < seq.length && typeof seq[i + 1] === 'string' && !EQ_STOP[seq[i + 1]] && /^[A-Za-z0-9]/.test(seq[i + 1])) {
          v += ' ';
        }
        out.push(v);
        i++;
        continue;
      }
      out.push(t);
      i++;
    }
    var s = '';
    out.forEach(function (piece) {
      if (piece && /\\[A-Za-z]+$/.test(s) && /^[A-Za-z]/.test(piece)) s += ' ';
      s += piece;
    });
    return s.replace(/ {2,}/g, ' ').trim();
  }

  function eqToLatex(script) {
    try {
      var parsed = parseEqTokens(tokenizeEq(script));
      return renderEqSeq(parsed.seq);
    } catch (e) {
      return String(script || '').trim();
    }
  }

  // -------------------------------------------------------------------
  // 3. HWPX(한글) → 마크다운(.md) 변환기 (hwpx2md.py 이식)
  // -------------------------------------------------------------------
  function localTag(node) {
    if (!node || !node.tagName) return '';
    var parts = node.tagName.split(':');
    return parts[parts.length - 1];
  }

  function childElements(node, tagName) {
    if (!node || !node.children) return [];
    var out = [];
    for (var i = 0; i < node.children.length; i++) {
      var ch = node.children[i];
      if (!tagName || localTag(ch) === tagName) out.push(ch);
    }
    return out;
  }

  function descElements(node, tagName) {
    if (!node) return [];
    var out = [];
    function walk(n) {
      if (!n || n.nodeType !== 1) return;
      if (localTag(n) === tagName) out.push(n);
      var kids = n.children || [];
      for (var i = 0; i < kids.length; i++) walk(kids[i]);
    }
    walk(node);
    return out;
  }

  function parseHwpxParaInline(pNode) {
    var buf = [];
    childElements(pNode, 'run').forEach(function (run) {
      childElements(run).forEach(function (el) {
        var tag = localTag(el);
        if (tag === 't') {
          for (var i = 0; i < el.childNodes.length; i++) {
            var cn = el.childNodes[i];
            if (cn.nodeType === 3) {
              buf.push(cn.nodeValue || '');
            } else if (cn.nodeType === 1) {
              var st = localTag(cn);
              if (st === 'tab') buf.push(' ');
              else if (st === 'lineBreak') buf.push('\n');
            }
          }
        } else if (tag === 'equation') {
          var sc = childElements(el, 'script')[0] || descElements(el, 'script')[0];
          if (sc && (sc.textContent || '').trim()) {
            buf.push(' $' + eqToLatex(sc.textContent) + '$ ');
          }
        }
      });
    });
    return buf.join('').replace(/[ \t]{2,}/g, ' ').trim();
  }

  function parseHwpxTable(tblNode, lines) {
    var rows = [];
    childElements(tblNode, 'tr').forEach(function (tr) {
      var row = [];
      childElements(tr, 'tc').forEach(function (tc) {
        var cellParts = descElements(tc, 'p').map(parseHwpxParaInline).filter(Boolean);
        row.push(cellParts.join(' ').replace(/\n/g, ' '));
      });
      rows.push(row);
    });
    if (!rows.length) return;
    var nCols = Math.max.apply(null, rows.map(function (r) { return r.length; }));
    if (nCols === 1) {
      lines.push('');
      rows.forEach(function (r) {
        r.forEach(function (c) { if (c) lines.push('> ' + c); });
      });
      lines.push('');
      return;
    }
    lines.push('');
    var normRows = rows.map(function (r) {
      var copy = r.slice();
      while (copy.length < nCols) copy.push('');
      return copy.map(function (c) { return c.replace(/\|/g, '\\|'); });
    });
    lines.push('| ' + normRows[0].join(' | ') + ' |');
    lines.push('|' + Array(nCols + 1).join('---|'));
    for (var i = 1; i < normRows.length; i++) {
      lines.push('| ' + normRows[i].join(' | ') + ' |');
    }
    lines.push('');
  }

  function walkHwpxPara(pNode, lines) {
    var text = parseHwpxParaInline(pNode);
    if (text) lines.push(text);
    childElements(pNode, 'run').forEach(function (run) {
      childElements(run).forEach(function (el) {
        var tag = localTag(el);
        if (tag === 'tbl') {
          parseHwpxTable(el, lines);
        } else if (tag === 'rect' || tag === 'ellipse' || tag === 'polygon' || tag === 'container' || tag === 'curve') {
          descElements(el, 'drawText').forEach(function (dt) {
            var sub = descElements(dt, 'p').map(parseHwpxParaInline).filter(Boolean);
            if (sub.length) {
              lines.push('');
              sub.forEach(function (x) { lines.push('> ' + x); });
              lines.push('');
            }
          });
        } else if (tag === 'pic') {
          lines.push('[그림]');
        }
      });
    });
  }

  async function convertHwpxToMd(arrayBuffer, fileName) {
    var zip = await JSZip.loadAsync(arrayBuffer);
    var secNames = Object.keys(zip.files)
      .filter(function (n) { return /^Contents\/section\d+\.xml$/i.test(n); })
      .sort(function (a, b) {
        var na = parseInt((a.match(/\d+/) || [0])[0], 10);
        var nb = parseInt((b.match(/\d+/) || [0])[0], 10);
        return na - nb;
      });
    var parser = new DOMParser();
    var lines = [];
    var nEq = 0, nPic = 0;

    for (var i = 0; i < secNames.length; i++) {
      var xmlStr = await zip.file(secNames[i]).async('string');
      var doc = parser.parseFromString(xmlStr, 'application/xml');
      var root = doc.documentElement;
      nEq += descElements(root, 'equation').length;
      nPic += descElements(root, 'pic').length;
      childElements(root, 'p').forEach(function (p) { walkHwpxPara(p, lines); });
      lines.push('');
    }
    var body = lines.join('\n').replace(/\n{3,}/g, '\n\n').replace(/(\[그림\]\n)+/g, '[그림]\n').trim();
    var head = '<!-- 원본: ' + fileName + ' / md 변환기 한글(HWPX) 변환본. 수식 ' + nEq.toLocaleString() + '개는 LaTeX($...$)로 옮겼고, 그림 ' + nPic.toLocaleString() + '개는 [그림]으로 표시함(그림 속 글자는 없음). 글상자는 > 로 표시함. -->';
    return {
      md: head + '\n\n' + body + '\n',
      body: body,
      nEq: nEq,
      nPic: nPic,
      fixed: 0,
      unknown: 0,
      warnMsg: ''
    };
  }

  // -------------------------------------------------------------------
  // 4. PDF / DOCX / PPTX / XLSX → 마크다운(.md) 변환기 (convert.py 이식)
  // -------------------------------------------------------------------
  async function convertPdfToMd(arrayBuffer, fileName) {
    if (!window.pdfjsLib) throw new Error('PDF 엔진(pdf.js)이 로드되지 않았습니다.');
    var pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    var pages = pdf.numPages;
    var pageTexts = [];
    for (var p = 1; p <= pages; p++) {
      var page = await pdf.getPage(p);
      var tc = await page.getTextContent();
      var lines = [];
      var lastY = null;
      tc.items.forEach(function (item) {
        var y = item.transform ? Math.round(item.transform[5]) : 0;
        if (lastY !== null && Math.abs(y - lastY) > 6) lines.push('\n');
        else if (lines.length && lines[lines.length - 1] !== '\n') lines.push(' ');
        lines.push(item.str || '');
        lastY = y;
      });
      pageTexts.push('## ' + p + '쪽\n\n' + lines.join('').replace(/[ \t]{2,}/g, ' ').trim());
    }
    var rawText = pageTexts.join('\n\n').trim();
    var pua = restorePua(rawText);
    var chars = pua.text.replace(/\s+/g, '').length;
    var warnMsg = '';
    if (pages > 0 && chars < pages * 80) {
      warnMsg = '주의: 글자가 거의 없습니다. 스캔한 PDF(그림 파일)일 수 있습니다. 글자를 선택할 수 있는 PDF로 다시 받아 주세요.';
    }
    var head = '<!-- 원본: ' + fileName + ' / md 변환기(MarkItDown) 변환본.';
    if (pua.fixed || pua.unknown) {
      head += ' 수식 글꼴 기호 ' + pua.fixed.toLocaleString() + '개를 문자로 복원했으며 분수, 행렬 등 2차원 수식은 한 줄로 풀려 있을 수 있음. □는 복원되지 않은 수식 기호(' + pua.unknown.toLocaleString() + '개)';
    }
    head += ' -->';
    return {
      md: head + '\n\n' + pua.text + '\n',
      body: pua.text,
      nEq: 0,
      nPic: 0,
      fixed: pua.fixed,
      unknown: pua.unknown,
      warnMsg: warnMsg
    };
  }

  async function convertDocxToMd(arrayBuffer, fileName) {
    var zip = await JSZip.loadAsync(arrayBuffer);
    var docFile = zip.file('word/document.xml');
    if (!docFile) throw new Error('유효한 워드(.docx) 파일이 아닙니다.');
    var xmlStr = await docFile.async('string');
    var doc = new DOMParser().parseFromString(xmlStr, 'application/xml');
    var bodyNode = descElements(doc.documentElement, 'body')[0] || doc.documentElement;
    var lines = [];

    childElements(bodyNode).forEach(function (ch) {
      var tag = localTag(ch);
      if (tag === 'p') {
        var txt = descElements(ch, 't').map(function (t) { return t.textContent || ''; }).join('').trim();
        if (!txt) return;
        var pStyle = descElements(ch, 'pStyle')[0];
        var styleVal = pStyle ? (pStyle.getAttribute('w:val') || pStyle.getAttribute('val') || '') : '';
        if (/Heading1|제목1/i.test(styleVal)) lines.push('# ' + txt);
        else if (/Heading2|제목2/i.test(styleVal)) lines.push('## ' + txt);
        else if (/Heading3|제목3/i.test(styleVal)) lines.push('### ' + txt);
        else lines.push(txt);
      } else if (tag === 'tbl') {
        var rows = childElements(ch, 'tr').map(function (tr) {
          return childElements(tr, 'tc').map(function (tc) {
            return descElements(tc, 't').map(function (t) { return t.textContent || ''; }).join(' ').replace(/\s+/g, ' ').trim();
          });
        });
        if (rows.length) {
          var nCols = Math.max.apply(null, rows.map(function (r) { return r.length; }));
          lines.push('');
          lines.push('| ' + rows[0].join(' | ') + ' |');
          lines.push('|' + Array(nCols + 1).join('---|'));
          for (var i = 1; i < rows.length; i++) lines.push('| ' + rows[i].join(' | ') + ' |');
          lines.push('');
        }
      }
    });
    var pua = restorePua(lines.join('\n\n').trim());
    var head = '<!-- 원본: ' + fileName + ' / md 변환기(MarkItDown) 워드(.docx) 변환본. -->';
    return { md: head + '\n\n' + pua.text + '\n', body: pua.text, nEq: 0, nPic: 0, fixed: pua.fixed, unknown: pua.unknown, warnMsg: '' };
  }

  async function convertPptxToMd(arrayBuffer, fileName) {
    var zip = await JSZip.loadAsync(arrayBuffer);
    var slideFiles = Object.keys(zip.files)
      .filter(function (k) { return /^ppt\/slides\/slide\d+\.xml$/i.test(k); })
      .sort(function (a, b) {
        return parseInt((a.match(/\d+/) || [0])[0], 10) - parseInt((b.match(/\d+/) || [0])[0], 10);
      });
    var parser = new DOMParser();
    var lines = [];
    for (var i = 0; i < slideFiles.length; i++) {
      var xmlStr = await zip.file(slideFiles[i]).async('string');
      var doc = parser.parseFromString(xmlStr, 'application/xml');
      lines.push('## 슬라이드 ' + (i + 1));
      descElements(doc.documentElement, 'sp').forEach(function (sp) {
        descElements(sp, 'p').forEach(function (p) {
          var t = descElements(p, 't').map(function (n) { return n.textContent || ''; }).join('').trim();
          if (t) lines.push('- ' + t);
        });
      });
      descElements(doc.documentElement, 'tbl').forEach(function (tbl) {
        var rows = childElements(tbl, 'tr').map(function (tr) {
          return childElements(tr, 'tc').map(function (tc) {
            return descElements(tc, 't').map(function (n) { return n.textContent || ''; }).join(' ').trim();
          });
        });
        if (rows.length) {
          var nCols = Math.max.apply(null, rows.map(function (r) { return r.length; }));
          lines.push('');
          lines.push('| ' + rows[0].join(' | ') + ' |');
          lines.push('|' + Array(nCols + 1).join('---|'));
          for (var r = 1; r < rows.length; r++) lines.push('| ' + rows[r].join(' | ') + ' |');
          lines.push('');
        }
      });
      lines.push('');
    }
    var pua = restorePua(lines.join('\n').trim());
    var head = '<!-- 원본: ' + fileName + ' / md 변환기(MarkItDown) 파워포인트(.pptx) 변환본. -->';
    return { md: head + '\n\n' + pua.text + '\n', body: pua.text, nEq: 0, nPic: 0, fixed: pua.fixed, unknown: pua.unknown, warnMsg: '' };
  }

  async function convertXlsxToMd(arrayBuffer, fileName) {
    var zip = await JSZip.loadAsync(arrayBuffer);
    var parser = new DOMParser();
    var sharedStrings = [];
    if (zip.file('xl/sharedStrings.xml')) {
      var ssDoc = parser.parseFromString(await zip.file('xl/sharedStrings.xml').async('string'), 'application/xml');
      descElements(ssDoc.documentElement, 'si').forEach(function (si) {
        sharedStrings.push(descElements(si, 't').map(function (t) { return t.textContent || ''; }).join(''));
      });
    }
    var sheetNames = Object.keys(zip.files)
      .filter(function (k) { return /^xl\/worksheets\/sheet\d+\.xml$/i.test(k); })
      .sort();
    var lines = [];
    for (var i = 0; i < sheetNames.length; i++) {
      var sDoc = parser.parseFromString(await zip.file(sheetNames[i]).async('string'), 'application/xml');
      var rows = [];
      descElements(sDoc.documentElement, 'row').forEach(function (rowNode) {
        var rowVals = [];
        childElements(rowNode, 'c').forEach(function (cNode) {
          var tAttr = cNode.getAttribute('t') || '';
          var vNode = childElements(cNode, 'v')[0];
          var isNode = childElements(cNode, 'is')[0];
          var val = '';
          if (tAttr === 's' && vNode) val = sharedStrings[parseInt(vNode.textContent, 10)] || '';
          else if (tAttr === 'inlineStr' && isNode) val = descElements(isNode, 't').map(function (t) { return t.textContent || ''; }).join('');
          else if (vNode) val = vNode.textContent || '';
          rowVals.push(String(val).replace(/\|/g, '\\|').trim());
        });
        if (rowVals.some(Boolean)) rows.push(rowVals);
      });
      if (rows.length) {
        lines.push('## 시트 ' + (i + 1));
        var nCols = Math.max.apply(null, rows.map(function (r) { return r.length; }));
        var norm = rows.map(function (r) {
          var c = r.slice();
          while (c.length < nCols) c.push('');
          return c;
        });
        lines.push('| ' + norm[0].join(' | ') + ' |');
        lines.push('|' + Array(nCols + 1).join('---|'));
        for (var r = 1; r < norm.length; r++) lines.push('| ' + norm[r].join(' | ') + ' |');
        lines.push('');
      }
    }
    var body = lines.join('\n').trim();
    var head = '<!-- 원본: ' + fileName + ' / md 변환기(MarkItDown) 엑셀(.xlsx) 변환본. -->';
    return { md: head + '\n\n' + body + '\n', body: body, nEq: 0, nPic: 0, fixed: 0, unknown: 0, warnMsg: '' };
  }

  // -------------------------------------------------------------------
  // 5. UI 상태 관리 & 일괄 변환 실행기
  // -------------------------------------------------------------------
  var convertedItems = [];
  var activeIdx = -1;

  function el(id) { return document.getElementById(id); }

  function appendLog(msg) {
    var box = el('mdConsoleLog');
    if (!box) return;
    box.textContent += '\n' + msg;
    box.scrollTop = box.scrollHeight;
  }

  function clearLog() {
    var box = el('mdConsoleLog');
    if (!box) return;
    box.textContent = '====================================================\n  md 변환기  ·  한글(hwpx), PDF, 워드, PPT, 엑셀 → 마크다운(.md)\n====================================================';
  }

  function escapeHtml(s) {
    return String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  function renderMarkdownPreviewHtml(mdText) {
    var lines = String(mdText || '').split('\n');
    var html = [];
    var inTable = false;

    function formatInline(str) {
      var s = escapeHtml(str);
      // LaTeX $...$ 수식 배지 강조
      s = s.replace(/\$([^$]+)\$/g, '<span class="md-latex-chip">$$$1$$</span>');
      s = s.replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>');
      return s;
    }

    for (var i = 0; i < lines.length; i++) {
      var ln = lines[i];
      if (/^<!--[\s\S]*-->$/.test(ln.trim())) {
        html.push('<div style="font-size:0.77rem;background:#f1f5f9;color:#475569;padding:6px 10px;border-radius:6px;margin-bottom:10px;border:1px solid #cbd5e1;">' + escapeHtml(ln.replace(/^<!--\s*|\s*-->$/g, '')) + '</div>');
        continue;
      }
      if (/^\|.*\|$/.test(ln.trim())) {
        if (/^\|[\s\-:|]+\|$/.test(ln.trim())) continue;
        var cells = ln.trim().slice(1, -1).split('|').map(function (c) { return formatInline(c.trim()); });
        if (!inTable) {
          inTable = true;
          html.push('<table><thead><tr>' + cells.map(function (c) { return '<th>' + c + '</th>'; }).join('') + '</tr></thead><tbody>');
        } else {
          html.push('<tr>' + cells.map(function (c) { return '<td>' + c + '</td>'; }).join('') + '</tr>');
        }
        continue;
      } else if (inTable) {
        inTable = false;
        html.push('</tbody></table>');
      }

      if (/^###\s+/.test(ln)) html.push('<h3>' + formatInline(ln.replace(/^###\s+/, '')) + '</h3>');
      else if (/^##\s+/.test(ln)) html.push('<h2>' + formatInline(ln.replace(/^##\s+/, '')) + '</h2>');
      else if (/^#\s+/.test(ln)) html.push('<h1>' + formatInline(ln.replace(/^#\s+/, '')) + '</h1>');
      else if (/^>\s*/.test(ln)) html.push('<blockquote>' + formatInline(ln.replace(/^>\s*/, '')) + '</blockquote>');
      else if (/^-\s+/.test(ln)) html.push('<div style="margin-left:12px;">• ' + formatInline(ln.replace(/^-\s+/, '')) + '</div>');
      else if (ln.trim()) html.push('<p style="margin:6px 0;">' + formatInline(ln) + '</p>');
    }
    if (inTable) html.push('</tbody></table>');
    return html.join('');
  }

  function selectConvertedItem(idx) {
    if (idx < 0 || idx >= convertedItems.length) return;
    activeIdx = idx;
    var item = convertedItems[idx];

    document.querySelectorAll('.md-file-item').forEach(function (node, i) {
      node.classList.toggle('active', i === idx);
    });

    el('mdActiveFileName').textContent = item.outName + ' (' + item.origName + ')';
    var pills = [];
    pills.push('<span class="md-stat-pill">글자 ' + item.chars.toLocaleString() + '자</span>');
    if (item.nEq > 0) pills.push('<span class="md-stat-pill latex">한글 수식 ' + item.nEq + '개 → LaTeX($...$) 변환 완료</span>');
    if (item.fixed > 0) pills.push('<span class="md-stat-pill latex">PUA 수식기호 ' + item.fixed + '개 복원</span>');
    if (item.nPic > 0) pills.push('<span class="md-stat-pill">그림 ' + item.nPic + '개 ([그림] 표시)</span>');
    el('mdActiveStats').innerHTML = pills.join('');

    el('mdRawTextarea').value = item.md;
    el('mdRenderedPreview').innerHTML = renderMarkdownPreviewHtml(item.md);
    el('mdDownloadCurBtn').disabled = false;
    el('mdCopyCurBtn').disabled = false;
  }

  function renderFileList() {
    var listEl = el('mdFileList');
    var okItems = convertedItems.filter(function (x) { return x.status === 'ok' || x.status === 'warn'; });
    el('mdCountSummary').textContent = '변환 결과 목록 (' + okItems.length + '개 완료)';
    el('mdDownloadAllZipBtn').disabled = okItems.length === 0;

    listEl.innerHTML = convertedItems.map(function (item, idx) {
      var badgeClass = item.status === 'ok' ? 'ok' : (item.status === 'warn' ? 'warn' : 'skip');
      var badgeText = item.status === 'ok' ? '완료' : (item.status === 'warn' ? '주의' : '건너뜀');
      var sub = item.status === 'skip'
        ? item.warnMsg
        : ('글자 ' + item.chars.toLocaleString() + '자' + (item.nEq ? ' · LaTeX 수식 ' + item.nEq + '개' : '') + (item.fixed ? ' · 기호복원 ' + item.fixed + '개' : ''));
      return '<div class="md-file-item ' + (idx === activeIdx ? 'active' : '') + '" data-md-idx="' + idx + '">' +
        '<div class="md-file-meta">' +
          '<strong>' + escapeHtml(item.outName || item.origName) + '</strong>' +
          '<small>' + escapeHtml(sub) + '</small>' +
        '</div>' +
        '<span class="md-status-tag ' + badgeClass + '">' + badgeText + '</span>' +
      '</div>';
    }).join('');

    listEl.querySelectorAll('.md-file-item').forEach(function (node) {
      node.addEventListener('click', function () {
        var i = parseInt(node.getAttribute('data-md-idx'), 10);
        selectConvertedItem(i);
      });
    });
  }

  async function processFiles(fileList) {
    var files = Array.from(fileList || []).filter(function (f) {
      return f.name && !f.name.startsWith('~$') && !f.name.startsWith('.') &&
             !f.name.toLowerCase().endsWith('.md') && f.name !== '여기에 파일을 넣으세요.txt';
    });
    if (!files.length) {
      appendLog('\n바꿀 파일이 없습니다.');
      return;
    }

    appendLog('\n파일 ' + files.length + '개를 확인합니다.\n');
    var doneCnt = 0, skipCnt = 0, failCnt = 0;

    for (var i = 0; i < files.length; i++) {
      var f = files[i];
      var tag = '[' + (i + 1) + '/' + files.length + '] ' + f.name;
      var dotIdx = f.name.lastIndexOf('.');
      var ext = dotIdx >= 0 ? f.name.slice(dotIdx).toLowerCase() : '';
      var stem = dotIdx >= 0 ? f.name.slice(0, dotIdx) : f.name;

      if (ext === '.hwp') {
        var hwpMsg = '건너뜀: 옛 한글 파일(hwp)입니다. 한글에서 [다른 이름으로 저장하기] → 파일 형식을 HWPX로 저장한 뒤 넣어 주세요.';
        appendLog(tag + '\n   ' + hwpMsg);
        convertedItems.push({ origName: f.name, outName: f.name, status: 'skip', md: hwpMsg, chars: 0, nEq: 0, nPic: 0, fixed: 0, warnMsg: hwpMsg });
        skipCnt++;
        continue;
      }
      if (OK_EXT.indexOf(ext) < 0) {
        var extMsg = '건너뜀: 한글(hwpx), PDF, 워드(docx), PPT(pptx), 엑셀(xlsx, xls)만 바꿀 수 있습니다.';
        appendLog(tag + '\n   ' + extMsg);
        skipCnt++;
        continue;
      }

      appendLog(tag + '\n   바꾸는 중...');
      try {
        var buf = await f.arrayBuffer();
        var res;
        if (ext === '.hwpx') res = await convertHwpxToMd(buf, f.name);
        else if (ext === '.pdf') res = await convertPdfToMd(buf, f.name);
        else if (ext === '.docx') res = await convertDocxToMd(buf, f.name);
        else if (ext === '.pptx') res = await convertPptxToMd(buf, f.name);
        else if (ext === '.xlsx' || ext === '.xls') res = await convertXlsxToMd(buf, f.name);

        var outName = stem + '.md';
        var chars = res.body.replace(/\s+/g, '').length;
        appendLog('   완료 → ' + outName + '  (글자 ' + chars.toLocaleString() + '자)');
        if (ext === '.hwpx') {
          appendLog('   수식 ' + res.nEq.toLocaleString() + '개를 수식 문법(LaTeX)으로 옮겼습니다. 그림 ' + res.nPic.toLocaleString() + '개는 [그림]으로만 표시했습니다.');
        }
        if (res.fixed || res.unknown) {
          appendLog('   수식 기호 ' + res.fixed.toLocaleString() + '개를 문자로 되돌렸습니다.' +
            (res.unknown ? (' 되돌리지 못한 기호 ' + res.unknown + '개는 □로 표시했습니다. 원본과 대조해 주세요.') : ''));
        }
        if (res.warnMsg) appendLog('   ' + res.warnMsg);

        convertedItems.push({
          origName: f.name,
          outName: outName,
          status: res.warnMsg ? 'warn' : 'ok',
          md: res.md,
          chars: chars,
          nEq: res.nEq,
          nPic: res.nPic,
          fixed: res.fixed,
          unknown: res.unknown,
          warnMsg: res.warnMsg
        });
        doneCnt++;
      } catch (err) {
        appendLog('   실패: ' + (err.message || err));
        failCnt++;
      }
    }

    appendLog('\n----------------------------------------------------');
    appendLog('완료 ' + doneCnt + '개 · 건너뜀 ' + skipCnt + '개 · 실패 ' + failCnt + '개');
    renderFileList();
    var firstOk = convertedItems.findIndex(function (x) { return x.status === 'ok' || x.status === 'warn'; });
    if (firstOk >= 0) selectConvertedItem(firstOk);
  }

  // -------------------------------------------------------------------
  // 6. 1클릭 샘플 파일 생성기 (실제 HWPX XML 수식·표 포함 ZIP 생성 후 변환)
  // -------------------------------------------------------------------
  async function createSampleHwpxFile() {
    var zip = new JSZip();
    var sectionXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<hs:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section">' +
        '<hp:p><hp:run><hp:t>2026학년도 수학과(미적분·확률과 통계) 교수·학습 지도안 및 평가 계획</hp:t></hp:run></hp:p>' +
        '<hp:p><hp:run>' +
          '<hp:t>1. 함수의 극한과 미분계수 핵심 정리: </hp:t>' +
          '<hp:equation><hp:script>lim _{x -&gt; 1} {{x ^{2} - 1} over {x - 1}} = 2</hp:script></hp:equation>' +
          '<hp:t> 이며, 도함수 정의는 </hp:t>' +
          '<hp:equation><hp:script>f prime (x) = lim _{h -&gt; 0} {{f(x+h) - f(x)} over h}</hp:script></hp:equation>' +
          '<hp:t> 이다.</hp:t>' +
        '</hp:run></hp:p>' +
        '<hp:p><hp:run>' +
          '<hp:t>2. 정적분과 급수의 합 공식: </hp:t>' +
          '<hp:equation><hp:script>int _{0} ^{pi} sin x dx = 2</hp:script></hp:equation>' +
          '<hp:t> 및 근의 공식 </hp:t>' +
          '<hp:equation><hp:script>x = {{-b +- sqrt {b ^{2} - 4ac}} over {2a}}</hp:script></hp:equation>' +
        '</hp:run></hp:p>' +
        '<hp:p><hp:run>' +
          '<hp:tbl>' +
            '<hp:tr>' +
              '<hp:tc><hp:p><hp:run><hp:t>단원명</hp:t></hp:run></hp:p></hp:tc>' +
              '<hp:tc><hp:p><hp:run><hp:t>핵심 성취기준 수식</hp:t></hp:run></hp:p></hp:tc>' +
              '<hp:tc><hp:p><hp:run><hp:t>배점 비중</hp:t></hp:run></hp:p></hp:tc>' +
            '</hp:tr>' +
            '<hp:tr>' +
              '<hp:tc><hp:p><hp:run><hp:t>Ⅰ. 수열의 극한</hp:t></hp:run></hp:p></hp:tc>' +
              '<hp:tc><hp:p><hp:run><hp:equation><hp:script>sum _{n=1} ^{inf} r ^{n-1} = {1 over {1-r}}</hp:script></hp:equation></hp:run></hp:p></hp:tc>' +
              '<hp:tc><hp:p><hp:run><hp:t>35%</hp:t></hp:run></hp:p></hp:tc>' +
            '</hp:tr>' +
            '<hp:tr>' +
              '<hp:tc><hp:p><hp:run><hp:t>Ⅱ. 미분법과 적분법</hp:t></hp:run></hp:p></hp:tc>' +
              '<hp:tc><hp:p><hp:run><hp:equation><hp:script>int _{a} ^{b} f(x) dx = F(b) - F(a)</hp:script></hp:equation></hp:run></hp:p></hp:tc>' +
              '<hp:tc><hp:p><hp:run><hp:t>65%</hp:t></hp:run></hp:p></hp:tc>' +
            '</hp:tr>' +
          '</hp:tbl>' +
        '</hp:run></hp:p>' +
        '<hp:p><hp:run>' +
          '<hp:rect><hp:drawText><hp:p><hp:run><hp:t>[교사용 지도 유의사항] 클로드·제미나이·노트북LM에 이 마크다운(.md) 파일을 올리면 LaTeX($...$) 수식과 표를 100% 정확하게 인식합니다.</hp:t></hp:run></hp:p></hp:drawText></hp:rect>' +
        '</hp:run></hp:p>' +
      '</hs:sec>';
    zip.file('Contents/section0.xml', sectionXml);
    var blob = await zip.generateAsync({ type: 'blob' });
    return new File([blob], '2026_미적분_교수학습지도안_수식포함.hwpx', { type: 'application/octet-stream' });
  }

  async function createSamplePptxFile() {
    var zip = new JSZip();
    var slideXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
        '<p:cSld><p:spTree>' +
          '<p:sp><p:txBody><a:p><a:r><a:t>1차시: 확률과 통계 조건부확률 수업 슬라이드</a:t></a:r></a:p>' +
          '<a:p><a:r><a:t>사건 A가 일어났을 때 사건 B의 조건부확률 P(B|A) = P(A∩B)/P(A)</a:t></a:r></a:p></p:txBody></p:sp>' +
        '</p:spTree></p:cSld>' +
      '</p:sld>';
    zip.file('ppt/slides/slide1.xml', slideXml);
    var blob = await zip.generateAsync({ type: 'blob' });
    return new File([blob], '확률과통계_수업자료.pptx', { type: 'application/octet-stream' });
  }

  async function createSampleXlsxFile() {
    var zip = new JSZip();
    var sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
        '<row r="1"><c t="inlineStr"><is><t>문항번호</t></is></c><c t="inlineStr"><is><t>영역</t></is></c><c t="inlineStr"><is><t>배점</t></is></c></row>' +
        '<row r="2"><c t="inlineStr"><is><t>1번</t></is></c><c t="inlineStr"><is><t>극한과 연속</t></is></c><c t="inlineStr"><is><t>4.2점</t></is></c></row>' +
        '<row r="3"><c t="inlineStr"><is><t>2번</t></is></c><c t="inlineStr"><is><t>정적분 활용</t></is></c><c t="inlineStr"><is><t>4.8점</t></is></c></row>' +
      '</sheetData></worksheet>';
    zip.file('xl/worksheets/sheet1.xml', sheetXml);
    var blob = await zip.generateAsync({ type: 'blob' });
    return new File([blob], '중간고사_이원목적분류표.xlsx', { type: 'application/octet-stream' });
  }

  // -------------------------------------------------------------------
  // 7. 이벤트 바인딩
  // -------------------------------------------------------------------
  function init() {
    var dz = el('mdDropzone');
    if (dz) {
      ['dragenter', 'dragover'].forEach(function (evName) {
        dz.addEventListener(evName, function (e) {
          e.preventDefault();
          e.stopPropagation();
          dz.classList.add('dragover');
        });
      });
      ['dragleave', 'drop'].forEach(function (evName) {
        dz.addEventListener(evName, function (e) {
          e.preventDefault();
          e.stopPropagation();
          dz.classList.remove('dragover');
        });
      });
      dz.addEventListener('drop', function (e) {
        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
          processFiles(e.dataTransfer.files);
        }
      });
    }

    if (el('mdFileInput')) {
      el('mdFileInput').addEventListener('change', function () {
        if (this.files && this.files.length) processFiles(this.files);
        this.value = '';
      });
    }
    if (el('mdFolderInput')) {
      el('mdFolderInput').addEventListener('change', function () {
        if (this.files && this.files.length) processFiles(this.files);
        this.value = '';
      });
    }

    if (el('mdSampleHwpxBtn')) {
      el('mdSampleHwpxBtn').addEventListener('click', async function () {
        var f = await createSampleHwpxFile();
        await processFiles([f]);
      });
    }
    if (el('mdSampleAllBtn')) {
      el('mdSampleAllBtn').addEventListener('click', async function () {
        var f1 = await createSampleHwpxFile();
        var f2 = await createSamplePptxFile();
        var f3 = await createSampleXlsxFile();
        await processFiles([f1, f2, f3]);
      });
    }

    if (el('mdClearAllBtn')) {
      el('mdClearAllBtn').addEventListener('click', function () {
        convertedItems = [];
        activeIdx = -1;
        clearLog();
        renderFileList();
        el('mdRawTextarea').value = '';
        el('mdRenderedPreview').innerHTML = '';
        el('mdDownloadCurBtn').disabled = true;
        el('mdCopyCurBtn').disabled = true;
      });
    }

    if (el('mdRawTextarea')) {
      el('mdRawTextarea').addEventListener('input', function () {
        if (activeIdx >= 0 && convertedItems[activeIdx]) {
          convertedItems[activeIdx].md = this.value;
          el('mdRenderedPreview').innerHTML = renderMarkdownPreviewHtml(this.value);
        }
      });
    }

    if (el('mdDownloadCurBtn')) {
      el('mdDownloadCurBtn').addEventListener('click', function () {
        if (activeIdx < 0 || !convertedItems[activeIdx]) return;
        var item = convertedItems[activeIdx];
        var blob = new Blob([el('mdRawTextarea').value || item.md], { type: 'text/markdown;charset=utf-8' });
        var a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = item.outName || '변환결과.md';
        a.click();
      });
    }

    if (el('mdCopyCurBtn')) {
      el('mdCopyCurBtn').addEventListener('click', function () {
        var txt = el('mdRawTextarea').value;
        if (!txt) return;
        if (navigator.clipboard) {
          navigator.clipboard.writeText(txt);
          this.textContent = '복사 완료! (AI에 붙여넣으세요)';
          var self = this;
          setTimeout(function () { self.textContent = 'AI 업로드용 전체 복사'; }, 2000);
        }
      });
    }

    if (el('mdDownloadAllZipBtn')) {
      el('mdDownloadAllZipBtn').addEventListener('click', async function () {
        var okItems = convertedItems.filter(function (x) { return x.status === 'ok' || x.status === 'warn'; });
        if (!okItems.length) return;
        var zip = new JSZip();
        var folder = zip.folder('md 결과');
        okItems.forEach(function (item) {
          folder.file(item.outName, item.md);
        });
        var blob = await zip.generateAsync({ type: 'blob' });
        var a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'md_결과_모음_' + new Date().toISOString().slice(0, 10) + '.zip';
        a.click();
      });
    }

    if (el('mdOpenGuideModalBtn')) {
      el('mdOpenGuideModalBtn').addEventListener('click', function () {
        el('mdGuideModal').classList.add('show');
      });
    }
    if (el('mdCloseGuideModalBtn')) {
      el('mdCloseGuideModalBtn').addEventListener('click', function () {
        el('mdGuideModal').classList.remove('show');
      });
    }
  }

  init();
})();
