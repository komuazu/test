/* 稼動日報 印刷実績ビューア
   data.json（build.py が生成）を読み、月別・営業部別に表示する。
   記入欄の内容は /api/memo（memo.json）に保存する。元の .xls は読み取りのみ。 */
'use strict';

var YEARS = null, DATA = null, MEMO = {},
    /* PEND = まだサーバーに届いていない記入欄。値が null なら「消した」。
       手元（localStorage）にも同じものを置き、届くまで消さない。 */
    PEND = {},
    DEPTKEY = { '本社営業部': 'hq', '東京営業部': 'tk', '池袋営業部': 'ik',
                '生産管理部（工務）': 'km', 'その他': 'ot' };
var S = { view: 'year', year: null, month: null, dept: '全社', deptC: '全社', monthC: 0,
          monthR: null, machineR: '全機械', q: '', qc: '', qr: '',
          monthO: null, oper: 'rate',
          monthW: 0, waste: 'rate', wasteSort: 'yare',
          monthWD: null, wasteDay: 'rate',
          chart: 'stack', chartDept: 'すべて' };
var MONTHS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];

var $ = function (s) { return document.querySelector(s); };
var $$ = function (s, root) {
  return Array.prototype.slice.call((root || document).querySelectorAll(s));
};
var num = function (n) { return (n || 0).toLocaleString('ja-JP'); };
var esc = function (s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
};
var el = function (tag, cls, html) {
  var e = document.createElement(tag);
  if (cls) e.className = cls;
  if (html != null) e.innerHTML = html;
  return e;
};

/* ── 起動 ──
   years.json（そのフォルダにどの年が入っているか）を読んでから、その年のデータを読む。
   13年〜25年のように複数年あるフォルダでも、画面上部で年を切り替えられる。 */
function boot() {
  PEND = loadPend();                     // 前回送れずに残った分（閉じ方が急でも拾える）
  Promise.all([
    fetch('years.json?t=' + Date.now()).then(function (r) { return r.ok ? r.json() : null; })
      .catch(function () { return null; }),
    fetch('/api/memo?t=' + Date.now()).then(function (r) {
      if (!r.ok) { throw new Error('記入欄を読めませんでした'); }
      return r.json();
    }).then(function (m) { return { memo: m || {}, live: true }; })
      .catch(function () { return { memo: localMemo(), live: false }; })
  ]).then(function (a) {
    YEARS = a[0];
    MEMO = a[1].memo;
    // サーバーから読めなかったときは、手元の控えを丸ごと「まだ送っていない分」に
    // する。つながり次第サーバーへ書き戻すので、記入欄が消えたままにならない。
    if (!a[1].live) {
      Object.keys(MEMO).forEach(function (k) {
        if (!(k in PEND)) { PEND[k] = MEMO[k]; }
      });
    }
    applyPend();                         // 前回閉じたときに送れなかった分
    indexMemo();
    writeLocal();
    if (hasPend()) { save(); }           // それをすぐ送り直す
    if (!YEARS || !YEARS.years || !YEARS.years.length) {
      $('#selYear').innerHTML = '';
      $('#warns').innerHTML = '<div class="warn"><b>データがありません。</b>'
        + '右上の「設定」で稼動日報フォルダを指定し、「保存して読み直す」を押してください。</div>';
      return;
    }
    var want = YEARS.years.some(function (y) { return y.year === S.year; })
      ? S.year : YEARS.current;
    renderYearSelect(want);
    loadYear(want);
  });
}

/* ヘッダーの年プルダウン（年が1つだけのときは出さない） */
function renderYearSelect(cur) {
  var sel = $('#selYear');
  sel.innerHTML = '';
  YEARS.years.forEach(function (y) {
    var o = el('option', null, y.year + '年（' + num(y.records) + '件）');
    o.value = y.year;
    if (y.year === cur) o.selected = true;
    sel.appendChild(o);
  });
  sel.style.display = YEARS.years.length > 1 ? '' : 'none';
  sel.onchange = function () { loadYear(+sel.value); };
}

/* その年のデータを読み込んで描き直す */
function loadYear(year) {
  fetch('data_' + year + '.json?t=' + Date.now())
    .then(function (r) {
      if (!r.ok) throw new Error(year + '年のデータが読み込めませんでした。');
      return r.json();
    })
    .then(function (d) {
      DATA = d;
      indexKeys();
      S.year = year;
      S.month = S.monthR = DATA.months[DATA.months.length - 1];
      S.monthC = 0;
      render();
    })
    .catch(function (e) {
      $('#warns').innerHTML = '<div class="warn"><b>' + esc(e.message) + '</b></div>';
    });
}

/* 画面上部に少しのあいだ出す案内 */
var sayTimer = null;
function say(html, ms) {
  var e = $('#saved');
  if (!e) { return; }
  e.innerHTML = html;
  e.className = 'saved on';
  e.title = '';
  clearTimeout(sayTimer);
  sayTimer = setTimeout(function () { showNote(false); }, ms || 2000);
}

/* 保存の状態。ng（保存できていない）のときは消さずに出しっぱなしにする。
   ヘッダのボタンを押し下げないよう、札は短く。わけは吹き出しに回す。 */
var noteTimer = null, NOTE = { msg: '保存しました', cls: 'ok', tip: '' };
function note(msg, cls, ms, tip) {
  NOTE = { msg: msg, cls: cls, tip: tip || '' };
  showNote(true);
  clearTimeout(noteTimer);
  if (ms) { noteTimer = setTimeout(function () { showNote(false); }, ms); }
}
function showNote(on) {
  var e = $('#saved');
  if (!e) { return; }
  e.textContent = NOTE.msg;
  e.title = NOTE.tip;
  e.className = 'saved ' + NOTE.cls + ((on || NOTE.cls === 'ng') ? ' on' : '');
}

/* ── 記入欄の手元の控え ──
   サーバーへ送るのは 0.4秒まとめてからなので、その間に画面を閉じられても
   消えないよう、打った時点で localStorage にも置く。送れたぶんだけ PEND から
   外すので、送れていない分は次に開いたときに必ず送り直される。 */
var LS_MEMO = 'kadouMemo', LS_PEND = 'kadouMemoPend';

function localMemo() { try { return JSON.parse(localStorage.getItem(LS_MEMO) || '{}') || {}; } catch (e) { return {}; } }
function loadPend() { try { return JSON.parse(localStorage.getItem(LS_PEND) || '{}') || {}; } catch (e) { return {}; } }
function hasPend() { return Object.keys(PEND).length > 0; }
function applyPend() {
  Object.keys(PEND).forEach(function (k) {
    if (PEND[k]) { MEMO[k] = PEND[k]; } else { delete MEMO[k]; }
  });
}
function writePend() { try { localStorage.setItem(LS_PEND, JSON.stringify(PEND)); } catch (e) { /* 容量超過は無視 */ } }
function writeLocal() {
  writePend();
  try { localStorage.setItem(LS_MEMO, JSON.stringify(MEMO)); } catch (e) { /* 容量超過は無視 */ }
}

/* ── 抽出ヘルパ ── */
function recs(month, dept) {
  return DATA.records.filter(function (r) {
    return (month == null || r.m === month) && (!dept || dept === '全社' || r.dept === dept);
  });
}
function key(r) { return DATA.year + '|' + r.m + '|' + r.dept + '|' + r.no; }

/* 部署を抜いた目印。部署の分け方を変えても記入欄を見失わないようにするため
   （営業担当ｺｰﾄﾞ6930を「その他」から「生産管理部（工務）」に移したときのように、
   同じ管理番号でも部署名が変わることがある） */
function key2(year, month, no) { return year + '|' + month + '|' + no; }

var MEMOIDX = {};
function indexMemo() {
  MEMOIDX = {};
  Object.keys(MEMO).forEach(function (k) {
    var p = k.split('|');
    if (p.length === 4) { MEMOIDX[key2(p[0], p[1], p[3])] = k; }
  });
}

/* いまの表にある目印の一覧。ほかの行のものを取り違えないための照合用 */
var KEYSET = {};
function indexKeys() {
  KEYSET = {};
  DATA.records.forEach(function (r) { KEYSET[key(r)] = 1; });
}

/* その行の記入欄が実際に入っているキー。部署名がいまと違っていても拾う。
   直すときにこのキーから作り直さないと、触っていない欄まで空になる。
   ただし、同じ月・同じ管理番号が別の部署にも並ぶことが実データで133件ある。
   その部署が「いまの表にある」なら別の行のものなので、拾わない。 */
function memoKey(r) {
  var k = key(r);
  if (MEMO[k]) { return k; }
  var o = MEMOIDX[key2(DATA.year, r.m, r.no)];
  return (o && MEMO[o] && !KEYSET[o]) ? o : k;
}

function memoOf(r) { return MEMO[memoKey(r)] || { trend: '', plan: '', tsu: '' }; }
function planTsu(r) { var v = parseFloat(String(memoOf(r).tsu).replace(/[^0-9.-]/g, '')); return isFinite(v) ? v : 0; }

/* ── 描画 ── */
function render() {
  $('#ttl').textContent = DATA.year + '年 稼動日報 印刷実績ビューア';
  $('#meta').textContent = '読込 ' + DATA.generated + '　／　'
    + DATA.files.length + 'ファイル　／　' + num(DATA.records.length) + '件'
    + (YEARS && YEARS.years.length > 1
       ? '　／　このフォルダの年: ' + YEARS.years.map(function (y) { return y.year; }).join('、')
       : '');
  var w = DATA.warnings || [];
  $('#warns').innerHTML = w.length
    ? '<div class="warn"><b>注意</b><br>' + w.map(esc).join('<br>') + '</div>' : '';
  renderChartDeptSelect(); renderYear(); renderCompare();
  renderMonthControls(); renderMonth();
  renderClientControls(); renderClient(); renderOper(); renderWaste();
  renderRawControls(); renderRaw(); renderSrc();
}

/* ── 前年との比較 ──
   前年の data_<年>.json は10MB前後あるので読み込まない。years.json に入れて
   ある「月×部署」の小さな集計だけで、年間と月別の前年比を出す。
   今年が年の途中までのときは、前年も同じ月までで比べる。 */
function yearSummary(year) {
  var ys = (YEARS && YEARS.years) || [];
  for (var i = 0; i < ys.length; i++) { if (ys[i].year === year) { return ys[i]; } }
  return null;
}

/* 今表示している年は、読み込み済みのデータから同じ形に集計する */
function curByDept() {
  var m = {};
  (DATA.stats || []).forEach(function (st) {
    (m[st.dept] || (m[st.dept] = {}))[String(st.m)] = [st.tsuGroups, st.groups];
  });
  return m;
}

function pick(byDept, dept, months, idx) {
  var d = byDept[dept] || {}, t = 0;
  months.forEach(function (m) { var v = d[String(m)]; if (v) { t += v[idx]; } });
  return t;
}

function ratio(now, before) {
  return before ? (now / before * 100).toFixed(1) + '%' : '－';
}

/* 前年より減っていれば赤、増えていれば太字（要望による見せ方） */
function mark(text, now, before) {
  var d = now - before;
  return '<span class="' + (d < 0 ? 'dn' : d > 0 ? 'up' : '') + '">' + text + '</span>';
}

function delta(now, before) {
  var d = now - before;
  return (d > 0 ? '+' : d < 0 ? '-' : '±') + num(Math.abs(d));
}

function renderChartDeptSelect() {
  var sel = $('#selChartDept'), keep = S.chartDept;
  sel.innerHTML = '';
  var list = ['すべて', '全社'].concat(DATA.depts);
  list.forEach(function (d) {
    var o = el('option', null, d === 'すべて' ? 'すべて（月ごとに部署を並べる）' : d);
    o.value = d;
    sel.appendChild(o);
  });
  if (list.indexOf(keep) < 0) { S.chartDept = 'すべて'; }
  sel.value = S.chartDept;
}

function renderCompare() {
  var prev = yearSummary(DATA.year - 1);
  var months = DATA.months || [];
  var ok = !!(prev && prev.byDept && months.length);
  $('#cmpYearCard').style.display = ok ? '' : 'none';
  if (!ok) { return; }

  var cur = curByDept(), pb = prev.byDept;
  var span = months.length === 12 ? '通年'
    : months[0] + '月〜' + months[months.length - 1] + '月';
  $('#cmpNote').textContent = DATA.year + '年（' + span + '）と '
    + prev.year + '年の同じ月';

  // ── 年間（部署別） ──
  var rows = DATA.depts.concat(['合計']);
  var h = '<thead><tr><th>部署</th><th class="num">通し数</th>'
    + '<th class="num">前年</th><th class="num">増減</th><th class="num">前年比</th>'
    + '<th class="num">件数</th><th class="num">前年</th><th class="num">前年比</th>'
    + '</tr></thead><tbody>';
  rows.forEach(function (d) {
    var all = d === '合計';
    var ds = all ? DATA.depts : [d];
    var ct = 0, pt = 0, cc = 0, pc = 0;
    ds.forEach(function (x) {
      ct += pick(cur, x, months, 0); pt += pick(pb, x, months, 0);
      cc += pick(cur, x, months, 1); pc += pick(pb, x, months, 1);
    });
    h += '<tr' + (all ? ' class="total"' : '') + '><th class="rh">' + esc(d) + '</th>'
      + '<td class="num">' + num(ct) + '</td>'
      + '<td class="num">' + num(pt) + '</td>'
      + '<td class="num">' + mark(delta(ct, pt), ct, pt) + '</td>'
      + '<td class="num">' + mark(ratio(ct, pt), ct, pt) + '</td>'
      + '<td class="num">' + num(cc) + '</td>'
      + '<td class="num">' + num(pc) + '</td>'
      + '<td class="num">' + mark(ratio(cc, pc), cc, pc) + '</td></tr>';
  });
  $('#cmpSum').innerHTML = h + '</tbody>';
}

/* ── 月別グラフ ──
   Excel の棒グラフと同じ形で描く。左に通し数の目盛り、うすい横線、下に月、
   凡例は下。「積み上げ（部署別）」と「前年と比較」を切り替えられる。
   どちらも部署を積み上げた形で、前年と比較のときは今年と前年の2本を並べる。 */
var CHARTCOL = { hq: '#4472c4', tk: '#ed7d31', ik: '#70ad47',
                 km: '#ffc000', ot: '#a5a5a5' };
var SVGNS = 'http://www.w3.org/2000/svg';
var patSeq = 0;

function deptColor(d) { return CHARTCOL[DEPTKEY[d]] || '#a5a5a5'; }

/* 目盛りの刻みを、1・2・2.5・5 の切りのいい数にそろえる */
function niceStep(max, want) {
  if (!(max > 0)) { return 1; }
  var raw = max / (want || 5);
  var mag = Math.pow(10, Math.floor(Math.log(raw) / Math.LN10));
  var step = mag * 10;
  [1, 2, 2.5, 5, 10].some(function (k) {
    if (raw <= mag * k) { step = mag * k; return true; }
    return false;
  });
  return step;
}

function sv(name, attrs) {
  var e = document.createElementNS(SVGNS, name);
  Object.keys(attrs || {}).forEach(function (k) { e.setAttribute(k, attrs[k]); });
  return e;
}

/* 棒や点にマウスを乗せたときの吹き出し */
function svTip(e, text) {
  var t = sv('title');
  t.textContent = text;
  e.appendChild(t);
  return e;
}

function svText(x, y, s, cls, anchor) {
  var t = sv('text', { x: x, y: y, 'text-anchor': anchor || 'middle' });
  if (cls) { t.setAttribute('class', cls); }
  t.textContent = s;
  return t;
}

/* 前年の棒に使う斜線。色ごとに1つ作る */
function hatchFill(defs, color) {
  var id = 'hatch' + (patSeq++);
  var p = sv('pattern', { id: id, width: 6, height: 6,
                          patternUnits: 'userSpaceOnUse',
                          patternTransform: 'rotate(45)' });
  p.appendChild(sv('rect', { width: 6, height: 6, fill: '#fff' }));
  p.appendChild(sv('rect', { width: 3, height: 6, fill: color }));
  defs.appendChild(p);
  return 'url(#' + id + ')';
}

/* 積み上げ棒グラフを描く
   opt = { months,
           groups: [{ name, hatch, series: [{name, color, values}] }],
           line: [%] または null, yTitle }
   groups が2つなら、月ごとに2本の積み上げを並べる（今年と前年）。 */
function drawColumns(box, opt) {
  box.innerHTML = '';
  var W = Math.max(box.clientWidth || 900, 560), H = 330;
  var L = 76, R = opt.line ? 56 : 16, T = 14, B = 54;
  var w = W - L - R, h = H - T - B;
  var n = opt.months.length || 1, ng = opt.groups.length;
  var svg = sv('svg', { class: 'chart', width: W, height: H,
                        viewBox: '0 0 ' + W + ' ' + H });  // 印刷で縮むように
  var defs = sv('defs');
  svg.appendChild(defs);

  // 縦の目盛り（通し数）
  var top = 0;
  opt.months.forEach(function (m, i) {
    opt.groups.forEach(function (g) {
      var t = 0;
      g.series.forEach(function (s) { t += s.values[i] || 0; });
      top = Math.max(top, t);
    });
  });
  var step = niceStep(top, 5), ymax = Math.max(Math.ceil(top / step) * step, step);
  for (var v = 0; v <= ymax + 0.5; v += step) {
    var y = T + h - (v / ymax) * h;
    svg.appendChild(sv('line', { class: v ? 'grid' : 'axis',
                                 x1: L, x2: L + w, y1: y, y2: y }));
    svg.appendChild(svText(L - 8, y + 4, num(Math.round(v)), 'tick', 'end'));
  }

  // 棒（積み上げ。前年と比較のときは2本並べる）
  var band = w / n, pad = band * 0.2;
  var each = (band - pad * 2) / ng, cw = each * (ng > 1 ? 0.78 : 0.56);
  opt.months.forEach(function (m, i) {
    opt.groups.forEach(function (g, gi) {
      var cx = L + band * i + pad + each * gi + (each - cw) / 2;
      var acc = 0, tot = 0;
      g.series.forEach(function (s) { tot += s.values[i] || 0; });
      g.series.forEach(function (s) {
        var val = s.values[i] || 0;
        if (!val) { return; }
        var bh = (val / ymax) * h;
        acc += bh;
        var r = sv('rect', { class: 'col', x: cx, y: T + h - acc,
                             width: cw, height: Math.max(bh, 0.5),
                             fill: g.hatch ? hatchFill(defs, s.color) : s.color });
        svg.appendChild(svTip(r, m + '月 ' + g.name + ' ' + s.name + '　'
                              + num(val) + ' 通し'));
      });
      if (tot) {
        svg.appendChild(svText(cx + cw / 2, T + h - acc - 4, num(tot), 'val'));
      }
      // 2本並ぶときは、どちらの年の棒かを下に出す（凡例を見に行かなくて済む）
      if (ng > 1 && g.tag) {
        svg.appendChild(svText(cx + cw / 2, T + h + 13, g.tag, 'ylab'));
      }
    });
    svg.appendChild(svText(L + band * i + band / 2, T + h + (ng > 1 ? 29 : 18),
                           m + '月', 'lab'));
  });

  // 区分線（隣り合う月の同じ段どうしをつなぐ細い線）
  var colX = function (i, gi) { return L + band * i + pad + each * gi + (each - cw) / 2; };
  var cum = function (g, i, k) {
    var t = 0;
    for (var q = 0; q <= k; q++) { t += g.series[q].values[i] || 0; }
    return t;
  };
  var yv = function (val) { return T + h - (val / ymax) * h; };
  var conn = function (x1, v1, x2, v2) {
    if (!v1 || !v2) { return; }
    svg.appendChild(sv('line', { class: 'conn',
                                 x1: x1, y1: yv(v1), x2: x2, y2: yv(v2) }));
  };
  if (ng === 1) {
    var g0 = opt.groups[0];
    g0.series.forEach(function (_s, k) {          // 月と月の同じ段をつなぐ
      for (var i = 0; i + 1 < n; i++) {
        conn(colX(i, 0) + cw, cum(g0, i, k), colX(i + 1, 0), cum(g0, i + 1, k));
      }
    });
  } else {
    // 2本並べるときは、同じ月の今年と前年の段をつなぐ。月をまたいでつなぐと
    // 隣の棒の上を線が走って読みにくいため、隣り合う2本の間だけを結ぶ。
    opt.months.forEach(function (_m, i) {
      opt.groups[0].series.forEach(function (_s, k) {
        conn(colX(i, 0) + cw, cum(opt.groups[0], i, k),
             colX(i, 1), cum(opt.groups[1], i, k));
      });
    });
  }

  // 前年比の折れ線（右の目盛り）
  if (opt.line) {
    var vals = opt.line.filter(function (x) { return x != null; });
    var lmax = Math.max(150, Math.ceil(Math.max.apply(null, vals.concat([0])) / 50) * 50);
    var ly = function (p) { return T + h - (p / lmax) * h; };
    for (var p = 0; p <= lmax + 0.5; p += 50) {
      svg.appendChild(svText(L + w + 8, ly(p) + 4, p + '%', 'tick2', 'start'));
    }
    svg.appendChild(sv('line', { class: 'base', x1: L, x2: L + w,
                                 y1: ly(100), y2: ly(100) }));
    var pts = [];
    opt.line.forEach(function (p2, i) {
      if (p2 == null) { return; }
      var cx2 = L + band * i + band / 2, cy = ly(Math.min(p2, lmax));
      pts.push(cx2 + ',' + cy);
      var dot = sv('circle', { class: 'dot', cx: cx2, cy: cy, r: 3 });
      svg.appendChild(svTip(dot, opt.months[i] + '月 前年比 ' + p2.toFixed(1) + '%'));
    });
    if (pts.length > 1) {
      svg.appendChild(sv('polyline', { class: 'rate', points: pts.join(' ') }));
    }
  }

  // 軸の名前
  var yt = svText(16, T + h / 2, opt.yTitle || '通し数', 'atitle');
  yt.setAttribute('transform', 'rotate(-90 16 ' + (T + h / 2) + ')');
  svg.appendChild(yt);
  svg.appendChild(svText(L + w / 2, H - 8, '月', 'atitle'));
  svg.appendChild(sv('line', { class: 'axis', x1: L, x2: L, y1: T, y2: T + h }));
  box.appendChild(svg);
}

/* 棒の下に出す年。稼動日報のシート名（「26年1月」）と同じ下2桁の書き方 */
function yearTag(y) { return ('0' + (y % 100)).slice(-2) + '年'; }

function legendHtml(items, extra) {
  return items.map(function (it) {
    return '<span><i style="background:' + it.color
      + (it.hatch ? ';background-image:repeating-linear-gradient(45deg,'
        + '#fff 0 1.5px,transparent 1.5px 4px)' : '') + '"></i>'
      + esc(it.name) + '</span>';
  }).join('') + (extra || '');
}

function renderChart() {
  var prev = yearSummary(DATA.year - 1);
  var canYoY = !!(prev && prev.byDept);
  $('#pillChart').querySelector('[data-c=yoy]').disabled = !canYoY;
  if (!canYoY && S.chart === 'yoy') { S.chart = 'stack'; }
  $$('#pillChart button').forEach(function (b) {
    b.className = b.dataset.c === S.chart ? 'on' : '';
  });
  var yoy = S.chart === 'yoy';
  $('#selChartDept').style.display = yoy ? '' : 'none';
  $('#chartNote').textContent = yoy
    ? DATA.year + '年 と ' + (DATA.year - 1) + '年　'
      + (S.chartDept === 'すべて' ? '部署別' : S.chartDept)
    : '部署の積み上げ';
  if (yoy) { renderChartYoY(prev); } else { renderChartStack(); }
}

/* 部署を積み上げた棒グラフ（今年だけ） */
function renderChartStack() {
  var series = DATA.depts.map(function (d) {
    return {
      name: d, color: deptColor(d),
      values: MONTHS.map(function (m) {
        return recs(m, d).reduce(function (s, r) { return s + r.tsu; }, 0);
      })
    };
  });
  drawColumns($('#chart'), {
    months: MONTHS, groups: [{ name: DATA.year + '年', series: series }],
    line: null, yTitle: '通し数'
  });
  $('#legend').innerHTML = legendHtml(series);
}

/* 今年と前年の積み上げを月ごとに並べる。前年は years.json の小さな集計から取る */
function renderChartYoY(prev) {
  var cur = curByDept(), pb = prev.byDept;
  var months = (DATA.months && DATA.months.length) ? DATA.months : MONTHS;
  var ds = (S.chartDept === 'すべて' || S.chartDept === '全社')
    ? DATA.depts : [S.chartDept];
  var mk = function (src) {
    return ds.map(function (d) {
      return {
        name: d, color: deptColor(d),
        values: months.map(function (m) { return pick(src, d, [m], 0); })
      };
    });
  };
  drawColumns($('#chart'), {
    months: months,
    groups: [{ name: DATA.year + '年', tag: yearTag(DATA.year), series: mk(cur) },
             { name: prev.year + '年', tag: yearTag(prev.year), hatch: true,
               series: mk(pb) }],
    line: null,
    yTitle: '通し数'
  });
  $('#legend').innerHTML = legendHtml(mk(cur))
    + '<span class="lg">塗り＝' + DATA.year + '年</span>'
    + '<span class="lg"><i class="hatch"></i>斜線＝' + prev.year + '年</span>';
}

/* 年間サマリー */
/* 営業部の見出しに添える営業担当ｺｰﾄﾞ。「その他」は数が多くなるので4つまで出す */
function deptCodeLabel(d) {
  var cs = DATA.deptCodes[d] || [];
  if (!cs.length) { return ''; }
  return cs.length > 4 ? cs.slice(0, 4).join(' / ') + ' ほか' + (cs.length - 4) + '件'
                       : cs.join(' / ');
}

function renderYear() {
  var all = DATA.records, tot = all.reduce(function (s, r) { return s + r.tsu; }, 0);
  // 年間 集計（営業部別）。集計Excelと同じ並び順で、合計を最後に置く
  var s = ['<thead><tr><th>部署</th><th class="num">通し数</th>'
    + '<th class="num">件数</th><th class="num">構成比</th></tr></thead><tbody>'];
  DATA.depts.forEach(function (d) {
    var rs = all.filter(function (r) { return r.dept === d; });
    var t = rs.reduce(function (a, r) { return a + r.tsu; }, 0);
    s.push('<tr><th class="rh">' + esc(d)
      + '<small>' + esc(deptCodeLabel(d)) + '</small></th>'
      + '<td class="num">' + num(t) + '</td>'
      + '<td class="num">' + num(rs.length) + '</td>'
      + '<td class="num">' + (tot ? (t / tot * 100).toFixed(1) : '0.0') + '%</td></tr>');
  });
  s.push('<tr class="total"><th class="rh">合計</th>'
    + '<td class="num">' + num(tot) + '</td>'
    + '<td class="num">' + num(all.length) + '</td>'
    + '<td class="num">' + (tot ? '100.0' : '0.0') + '%</td></tr></tbody>');
  $('#sum').innerHTML = s.join('');

  renderChart();

  renderMatrix();
}

/* ── 月別 × 部署別 ──
   「月別×部署別（通し数・件数）」と「月別 前年比（前年比・増減）」は、どちらも
   行=部署／列=月の同じ形だったので1つの表にまとめた。部署ごとに区分の行を
   4段（通し数・件数・前年比・増減）並べる。前年の稼動日報が無い年では、
   前年比・増減の2段は出さない。 */

/* この表で使う区分。前年の稼動日報が無い年では、通し数と件数だけになる。
   前年・増減・前年比は通し数どうしの比較なので、通し数のすぐ下にまとめて置く */
function matrixKinds() {
  var prev = yearSummary(DATA.year - 1);
  if (!prev || !prev.byDept || !(DATA.months || []).length) {
    return [{ k: 'tsu', label: '本年（通し数）' }, { k: 'cnt', label: '件数' }];
  }
  return [{ k: 'tsu', label: '本年（通し数）' }, { k: 'pre', label: '前年' },
          { k: 'dif', label: '増減' }, { k: 'yoy', label: '前年比' },
          { k: 'cnt', label: '件数' }];
}

/* 部署（'合計' なら全社）の、その月の [通し数, 件数] */
function cell(dept, month) {
  var rs = recs(month, dept === '合計' ? null : dept);
  return [rs.reduce(function (s, r) { return s + r.tsu; }, 0), rs.length];
}

function renderMatrix() {
  var kinds = matrixKinds(), yoy = kinds.length > 2;   // 前年の行が出るか
  var prev = yoy ? yearSummary(DATA.year - 1) : null;
  var cur = yoy ? curByDept() : null, pb = prev ? prev.byDept : null;
  var have = DATA.months || [];   // 今年データのある月。前年比はこの月だけ出す

  var h = '<thead><tr><th>部署</th><th>区分</th>';
  MONTHS.forEach(function (m) { h += '<th class="num">' + m + '月</th>'; });
  h += '<th class="num">年間合計</th></tr></thead>';

  // 部署ごとに tbody を分ける。区切り線を引きやすく、印刷でも4段が離れない
  DATA.depts.concat(['合計']).forEach(function (d) {
    var all = d === '合計';
    var ds = all ? DATA.depts : [d];
    h += '<tbody' + (all ? ' class="total"' : '') + '>';
    kinds.forEach(function (kind, i) {
      h += '<tr class="k-' + kind.k + '">';
      if (i === 0) {
        h += '<th class="rh" rowspan="' + kinds.length + '">' + esc(d)
          + (all ? '' : '<small>' + esc(deptCodeLabel(d)) + '</small>') + '</th>';
      }
      h += '<th class="kh">' + kind.label + '</th>';
      MONTHS.concat([0]).forEach(function (m) {     // 0 = 年間合計の列
        h += matrixCell(kind.k, d, ds, m, cur, pb, have);
      });
      h += '</tr>';
    });
    h += '</tbody>';
  });
  $('#matrix').innerHTML = h;
}

/* 1マス分。m=0 は年間合計の列 */
function matrixCell(k, d, ds, m, cur, pb, have) {
  var year = !m, name = year ? '年間合計' : m + '月';
  var cls = 'num' + (year ? ' yr' : '');
  if (k === 'tsu' || k === 'cnt') {
    var v = cell(d, year ? null : m)[k === 'tsu' ? 0 : 1];
    var txt = v ? (k === 'cnt' ? num(v) + '件' : num(v)) : '－';
    return '<td class="' + cls + '">' + txt + '</td>';
  }
  // 前年・増減・前年比は、今年データのある月だけ出す。年の途中までしか入って
  // いない年で、未到来の月を 0% や前年だけの数字で埋めてしまわないようにする
  var ms = year ? have : (have.indexOf(m) >= 0 ? [m] : null);
  if (!ms || !ms.length) { return '<td class="' + cls + '"></td>'; }
  var c = 0, p = 0;
  ds.forEach(function (x) { c += pick(cur, x, ms, 0); p += pick(pb, x, ms, 0); });
  if (k === 'pre') { return '<td class="' + cls + '">' + (p ? num(p) : '－') + '</td>'; }
  return '<td class="' + cls + '" title="' + esc(d) + ' ' + name
    + '　本年 ' + num(c) + ' ／ 前年 ' + num(p) + '">'
    + mark(k === 'yoy' ? ratio(c, p) : delta(c, p), c, p) + '</td>';
}

/* 月別明細 */
/* 月のプルダウンを作る（all=true なら「年間（全月）」を先頭に足す） */
function fillMonths(sel, cur, all, onchange) {
  sel.innerHTML = '';
  if (all) {
    var a = el('option', null, DATA.year + '年 年間（全月）');
    a.value = 0;
    if (!cur) a.selected = true;
    sel.appendChild(a);
  }
  MONTHS.forEach(function (m) {
    var n = recs(m, null).length;
    var o = el('option', null, DATA.year + '年 ' + m + '月' + (n ? '（' + n + '件）' : '（データなし）'));
    o.value = m; o.disabled = !n;
    if (m === cur) o.selected = true;
    sel.appendChild(o);
  });
  sel.onchange = function () { onchange(+sel.value); };
}

function renderMonthControls() {
  fillMonths($('#selMonth'), S.month, false, function (m) { S.month = m; renderMonth(); });

  var p = $('#pillDept'); p.innerHTML = '';
  ['全社'].concat(DATA.depts).forEach(function (d) {
    var b = el('button', S.dept === d ? 'on' : '', d);
    b.onclick = function () { S.dept = d; renderMonthControls(); renderMonth(); };
    p.appendChild(b);
  });
  $('#q').value = S.q;
  $('#q').oninput = function () { S.q = this.value; renderMonth(); };
}

function filtered() {
  var rs = recs(S.month, S.dept), q = S.q.trim().toLowerCase();
  if (q) {
    rs = rs.filter(function (r) {
      return (r.no + ' ' + r.client + ' ' + r.name).toLowerCase().indexOf(q) >= 0;
    });
  }
  return rs;
}

function renderMonth() {
  var rs = filtered();
  $('#mTitle').innerHTML = DATA.year + '年' + S.month + '月　' + esc(S.dept)
    + ' <small>' + rs.length + '件　通し数 ' + num(rs.reduce(function (s, r) { return s + r.tsu; }, 0)) + '</small>';
  var t = $('#detail');
  if (!rs.length) {
    t.innerHTML = '<tbody><tr><td class="empty">該当するデータがありません。</td></tr></tbody>';
    return;
  }
  var all = S.dept === '全社';
  var h = '<thead><tr>'
    + (all ? '<th>部署</th>' : '')
    + '<th>管理番号</th><th>ｸﾗｲｱﾝﾄ名</th><th>品名</th>'
    + '<th class="inp">今年の動向</th><th class="inp">無しの場合の代替対策</th><th class="inp num">対策通し数</th>'
    + '<th class="c">営業担当ｺｰﾄﾞ</th><th class="c">印刷日</th><th class="c">色数</th><th class="num">通し数</th>'
    + '</tr></thead><tbody>';
  rs.forEach(function (r) {
    var m = memoOf(r), fill = (m.trend || m.plan || m.tsu) ? ' filled' : '';
    var tip = '機械: ' + r.machines.join(' / ') + (r.nos.length ? '\n統合した管理番号: ' + r.nos.join(', ') : '');
    h += '<tr data-i="' + r.i + '">'
      + (all ? '<td>' + esc(shortDept(r.dept)) + '</td>' : '')
      + '<td title="' + esc(tip) + '"><a href="#" class="drill" title="日報の元の行を表示">▶</a> '
        + (r.nos.length ? '<b>' + esc(r.no) + '</b><span class="badge">枝番'
        + r.nos.length + '</span>' : esc(r.no)) + '</td>'
      + '<td>' + esc(r.client) + '</td>'
      + '<td>' + esc(r.name) + '</td>'
      + '<td class="inp' + fill + '"><textarea rows="1" data-f="trend">' + esc(m.trend) + '</textarea></td>'
      + '<td class="inp' + fill + '"><textarea rows="1" data-f="plan">' + esc(m.plan) + '</textarea></td>'
      + '<td class="inp' + fill + '"><input class="num" data-f="tsu" value="' + esc(m.tsu) + '"></td>'
      + '<td class="c">' + r.code + '</td>'
      + '<td class="c' + (r.multi ? ' multi' : '') + '" title="' + (r.multi ? '複数日にまたがります（初日を表示）' : '') + '">'
        + r.date.slice(5).replace('-', '/') + '</td>'
      + '<td class="c">' + esc(r.color) + '</td>'
      + '<td class="num">' + num(r.tsu) + '</td></tr>';
  });
  var tot = rs.reduce(function (s, r) { return s + r.tsu; }, 0);
  var pt = rs.reduce(function (s, r) { return s + planTsu(r); }, 0);
  h += '<tr class="total">' + (all ? '<td></td>' : '') + '<td>合計</td><td></td><td>' + rs.length + '件</td><td></td><td></td>'
    + '<td class="num" id="sumPlan">' + num(pt) + '</td><td></td><td></td><td></td>'
    + '<td class="num">' + num(tot) + '</td></tr>'
    + '<tr class="diff"><td colspan="' + (all ? 6 : 5) + '">差引（通し数 − 対策通し数）</td>'
    + '<td class="num' + (tot - pt < 0 ? ' dn' : '') + '" id="sumDiff">'
    + num(tot - pt) + '</td><td></td><td></td>'
    + '<td class="c">充足率→</td><td class="num" id="sumRate">'
    + (tot ? (pt / tot * 100).toFixed(1) + '%' : '－') + '</td></tr></tbody>';
  t.innerHTML = h;

  var span = t.querySelector('thead tr').children.length;
  $$('#detail a.drill').forEach(function (a) {
    a.onclick = function (e) {
      e.preventDefault();
      var tr = a.closest('tr'), nx = tr.nextElementSibling;
      if (nx && nx.classList.contains('sub')) { nx.remove(); a.textContent = '▶'; return; }
      var r = DATA.records[+tr.dataset.i];
      var sub = document.createElement('tr');
      sub.className = 'sub';
      sub.innerHTML = '<td colspan="' + span + '"><div class="subwrap">'
        + '<div class="subttl">' + esc(r.no) + ' の内訳　日報の明細 ' + r.det.length + '行</div>'
        + detTable(r.det) + '</div></td>';
      tr.after(sub);
      a.textContent = '▼';
    };
  });

  $$('#detail textarea, #detail input').forEach(function (f) {
    autosize(f);
    f.oninput = function () { autosize(f); onEdit(f); };
    f.onchange = function () { onEdit(f, true); };
  });
}

/* 日報の値の表示。ｺｰﾄﾞや管理番号は数量ではないので桁区切りを付けない */
function rawVal(col, v) {
  if (v === undefined || v === null || v === '') return '';
  if (typeof v !== 'number') return esc(v);
  return /ｺｰﾄﾞ|コード|管理番号|No|番号/.test(col) ? esc(String(v)) : num(v);
}

/* 日報の明細行をそのままの列構成で表にする */
function detTable(dets) {
  var cols = ['__機械'].concat(DATA.cols).filter(function (c) {
    return dets.some(function (d) { return d[c] !== undefined && d[c] !== null && d[c] !== ''; });
  });
  var h = '<table class="sub"><thead><tr>' + cols.map(function (c) {
    return '<th>' + esc(c === '__機械' ? '機械' : c) + '</th>';
  }).join('') + '</tr></thead><tbody>';
  dets.forEach(function (d) {
    h += '<tr>' + cols.map(function (c) {
      return '<td class="' + (typeof d[c] === 'number' ? 'num' : '') + '">'
        + rawVal(c, d[c]) + '</td>';
    }).join('') + '</tr>';
  });
  return h + '</tbody></table>';
}

function autosize(f) {
  if (f.tagName !== 'TEXTAREA') return;
  f.style.height = 'auto';
  f.style.height = Math.max(26, f.scrollHeight) + 'px';
}

var saveTimer = null, retryTimer = null, sending = false;

function onEdit(f, now) {
  var tr = f.closest('tr'), r = DATA.records[+tr.dataset.i];
  var k = key(r), old = memoKey(r), b = MEMO[old] || {};
  // いま入っている値から作り直す。MEMO[k] だけを見ていたころは、部署名が
  // 変わった行を1欄だけ直すと、触っていない残り2欄が空になっていた。
  var m = { trend: b.trend || '', plan: b.plan || '', tsu: b.tsu || '' };
  m[f.dataset.f] = f.value;
  if (old !== k && MEMO[old]) {          // 古い部署名のぶんは、いまの目印へ寄せる
    delete MEMO[old];
    PEND[old] = null;
  }
  if (m.trend || m.plan || m.tsu) { MEMO[k] = m; PEND[k] = m; }
  else { delete MEMO[k]; PEND[k] = null; }   // 消したことも伝える（黙って戻らないように）
  indexMemo();
  writePend();                           // 打った時点で手元に残す
  tr.querySelectorAll('td.inp').forEach(function (td) {
    td.classList.toggle('filled', !!(m.trend || m.plan || m.tsu));
  });
  updateTotals();
  clearTimeout(saveTimer);
  saveTimer = setTimeout(save, now ? 0 : 400);
}

function updateTotals() {
  var rs = filtered();
  var tot = rs.reduce(function (s, r) { return s + r.tsu; }, 0);
  var pt = rs.reduce(function (s, r) { return s + planTsu(r); }, 0);
  if ($('#sumPlan')) $('#sumPlan').textContent = num(pt);
  if ($('#sumDiff')) {
    $('#sumDiff').textContent = num(tot - pt);
    $('#sumDiff').className = 'num' + (tot - pt < 0 ? ' dn' : '');
  }
  if ($('#sumRate')) $('#sumRate').textContent = tot ? (pt / tot * 100).toFixed(1) + '%' : '－';
}

/* まだ送っていない分だけをサーバーへ送る。
   ・送れた分だけ PEND から外す（送っている間に直した分は残す）
   ・失敗したら手元に持ったまま、5秒ごとに送り直す
   ・成功も失敗も同じ「保存しました」を出していたので、気づけなかった */
function save() {
  clearTimeout(retryTimer);
  if (!hasPend() || sending) { return; }
  var sent = {};
  Object.keys(PEND).forEach(function (k) { sent[k] = PEND[k]; });
  sending = true;
  note('保存中…', 'busy', 0);
  fetch('/api/memo', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(sent)
  }).then(function (r) {
    if (!r.ok) { throw new Error('HTTP ' + r.status); }
    return r.json();
  }).then(function (j) {
    if (!j || j.ok !== true) { throw new Error('ことわられました'); }
    sending = false;
    Object.keys(sent).forEach(function (k) {
      if (sameMemo(PEND[k], sent[k])) { delete PEND[k]; }
    });
    writeLocal();
    note('保存しました', 'ok', 1200);
    if (hasPend()) { clearTimeout(saveTimer); saveTimer = setTimeout(save, 200); }
  }).catch(function () {
    sending = false;
    note('保存できていません', 'ng', 0,
         'つながり次第もう一度保存します。'
         + 'この画面を閉じても入力は残っているので、開き直せば保存されます。');
    retryTimer = setTimeout(save, 5000);
  });
}

function sameMemo(a, b) {
  if (!a || !b) { return !a && !b; }
  return a.trend === b.trend && a.plan === b.plan && a.tsu === b.tsu;
}

/* 画面を閉じる・再読込するときに、まだ送っていない分を投げておく。
   0.4秒まとめる間にリロードされると、入力が丸ごと消えていた。
   sendBeacon は画面が閉じたあとも送り切ってくれる。 */
function flush() {
  if (!hasPend()) { return; }
  var body = JSON.stringify(PEND);
  try {
    if (navigator.sendBeacon
        && navigator.sendBeacon('/api/memo', new Blob([body], { type: 'application/json' }))) {
      return;
    }
  } catch (e) { /* 下の fetch で送る */ }
  try {
    fetch('/api/memo', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: body, keepalive: true
    });
  } catch (e) { /* 手元に残っているので、次に開いたとき送り直す */ }
}

window.addEventListener('pagehide', flush);
window.addEventListener('beforeunload', flush);
document.addEventListener('visibilitychange', function () {
  if (document.visibilityState === 'hidden') { flush(); }
});

/* 得意先別 */
function renderClientControls() {
  fillMonths($('#selMonthC'), S.monthC, true, function (m) { S.monthC = m; renderClient(); });
  var p = $('#pillDeptC'); p.innerHTML = '';
  ['全社'].concat(DATA.depts).forEach(function (d) {
    var b = el('button', S.deptC === d ? 'on' : '', d);
    b.onclick = function () { S.deptC = d; renderClientControls(); renderClient(); };
    p.appendChild(b);
  });
  $('#qc').value = S.qc;
  $('#qc').oninput = function () { S.qc = this.value; renderClient(); };
}

/* 得意先別の表に入れる短い部署名（本社営業部→本社、生産管理部（工務）→工務） */
function shortDept(d) {
  return d === '生産管理部（工務）' ? '工務' : d.replace('営業部', '');
}

/* ｸﾗｲｱﾝﾄ名の表記ゆれ吸収キー。

   NFKC で ㈱→(株)、半角カナ・全角英数をそろえたうえで、空白と「会社の種類」を
   表す語を落とす。日報では同じ得意先でも（株）が付いたり付かなかったりするため、
   「武陽ｶﾞｽ（株）」「武陽ガス」「武陽ガス株式会社」を同じ得意先として数える。
   落とすと空になる名前（「（株）」だけ など）は、落とす前をそのまま使う。 */
var CORPWORD = /株式会社|有限会社|合同会社|合名会社|合資会社|\(株\)|\(有\)|\(名\)|\(資\)|\(同\)/g;
function ckey(s) {
  var t = String(s || '').normalize('NFKC').replace(/[\s　]/g, '');
  return t.replace(CORPWORD, '') || t;
}

function clientRows() {
  var map = {};
  recs(S.monthC || null, S.deptC).forEach(function (r) {
    var k = r.client ? ckey(r.client) : '';
    var e = map[k] || (map[k] = { name: '（得意先名なし）', tsu: 0, cnt: 0, months: {},
                                  names: {}, items: [], depts: {} });
    if (r.client) e.names[r.client] = (e.names[r.client] || 0) + 1;
    e.depts[r.dept] = (e.depts[r.dept] || 0) + r.tsu;
    e.items.push({ name: r.name, tsu: r.tsu, no: r.no });
    e.tsu += r.tsu; e.cnt++;
    e.months[r.m] = (e.months[r.m] || 0) + r.tsu;
  });
  var rows = Object.keys(map).map(function (k) {
    var e = map[k], ns = Object.keys(e.names);
    if (ns.length) {
      // 代表表記: 出現回数が多いもの → 同数なら長いもの（明細の統合ルールと同じ考え方）
      ns.sort(function (a, b) { return (e.names[b] - e.names[a]) || (b.length - a.length) || (a < b ? -1 : 1); });
      e.name = ns[0];
      e.alias = ns.length > 1 ? ns : null;
    }
    e.items.sort(function (a, b) { return b.tsu - a.tsu; });
    e.dept = Object.keys(e.depts).sort(function (a, b) { return e.depts[b] - e.depts[a]; })
      .map(shortDept).join('・');
    return e;
  });
  var q = S.qc.trim().toLowerCase();
  if (q) rows = rows.filter(function (e) { return e.name.toLowerCase().indexOf(q) >= 0; });
  rows.sort(function (a, b) { return b.tsu - a.tsu; });
  return rows;
}

function renderClient() {
  var rows = clientRows(), tot = rows.reduce(function (s, e) { return s + e.tsu; }, 0);
  var one = !!S.monthC;                       // 月を選んでいるか（0 = 年間）
  $('#cTitle').innerHTML = DATA.year + '年' + (one ? S.monthC + '月' : ' 年間') + '　' + esc(S.deptC)
    + ' <small>得意先 ' + rows.length + '社　' + rows.reduce(function (s, e) { return s + e.cnt; }, 0)
    + '件　通し数 ' + num(tot) + '</small>';
  var t = $('#clients');
  if (!rows.length) { t.innerHTML = '<tbody><tr><td class="empty">該当なし</td></tr></tbody>'; return; }
  var h = '<thead><tr><th class="c">順位</th><th>ｸﾗｲｱﾝﾄ名</th>'
    + (S.deptC === '全社' ? '<th class="c">部署</th>' : '')
    + '<th class="num">件数</th><th class="num">通し数</th><th class="num">構成比</th>';
  if (one) h += '<th>品名（管理番号 / 通し数）</th>';
  else MONTHS.forEach(function (m) { h += '<th class="num">' + m + '月</th>'; });
  h += '</tr></thead><tbody>';
  rows.forEach(function (e, i) {
    h += '<tr><td class="c">' + (i + 1) + '</td><td' + (e.alias ? ' title="同一とみなした表記: ' + esc(e.alias.join(' / ')) + '"' : '') + '>'
      + esc(e.name) + (e.alias ? '<span class="badge">表記' + e.alias.length + '</span>' : '') + '</td>'
      + (S.deptC === '全社' ? '<td class="c">' + esc(e.dept) + '</td>' : '')
      + '<td class="num">' + e.cnt + '</td><td class="num"><b>' + num(e.tsu) + '</b></td>'
      + '<td class="num">' + (tot ? (e.tsu / tot * 100).toFixed(1) : '0.0') + '%</td>';
    if (one) {
      h += '<td class="items">' + e.items.map(function (x) {
        return esc(x.name || '（品名なし）') + ' <small>（' + esc(x.no) + ' / ' + num(x.tsu) + '）</small>';
      }).join('<br>') + '</td>';
    } else {
      MONTHS.forEach(function (m) { h += '<td class="num">' + (e.months[m] ? num(e.months[m]) : '') + '</td>'; });
    }
    h += '</tr>';
  });
  h += '<tr class="total"><td></td><td>合計</td>' + (S.deptC === '全社' ? '<td></td>' : '')
    + '<td class="num">' + rows.reduce(function (s, e) { return s + e.cnt; }, 0) + '</td>'
    + '<td class="num">' + num(tot) + '</td><td class="num">100.0%</td>';
  if (one) h += '<td></td>';
  else MONTHS.forEach(function (m) {
    var v = rows.reduce(function (s, e) { return s + (e.months[m] || 0); }, 0);
    h += '<td class="num">' + (v ? num(v) : '') + '</td>';
  });
  h += '</tr></tbody>';
  t.innerHTML = h;
}


/* ── 稼働率 ──
   分母 = 就業可能時間（1日22時間の2直 × その月の稼働日数）
   分子 = 稼動日報の「有効時間」（機械が実際に動いていた時間）
   稼働日 = 平日 − 祝日（振替休日・国民の休日を含む） − 会社の休業日
   もとになる数字は build.py が data_<年>.json の oper に入れてくれる。 */
var WDAY = ['日', '月', '火', '水', '木', '金', '土'];

function oper() { return DATA && DATA.oper; }
function operMonth(m) {
  var o = oper();
  return (o && o.months && o.months[String(m)]) || null;
}

/* その月・その機械の分母（時間）。機械を選ばなければ全機械ぶん */
function operBase(m, machines) {
  var o = oper(), v = operMonth(m);
  if (!o || !v) { return 0; }
  return o.shift * v.work * (machines == null ? 1 : machines);
}

function rate(hours, base) {
  return base ? (hours / base * 100).toFixed(1) + '%' : '－';
}

/* 稼働率の高さで濃さを変える（表を目で追えるように） */
function heat(pct) {
  if (pct == null) { return ''; }
  var k = Math.max(0, Math.min(1, pct / 100));
  return ' style="background:rgba(58,68,80,' + (0.03 + k * 0.17).toFixed(3) + ')"';
}

function renderOper() {
  var o = oper();
  var box = $('#v-oper');
  if (!o || !o.machines.length) {
    $('#operSum').innerHTML = '<tbody><tr><td class="empty">'
      + 'この年の稼動日報に「有効時間」が入っていないため、稼働率は出せません。'
      + '</td></tr></tbody>';
    $('#operDay').innerHTML = '';
    $('#operNote').textContent = '';
    return;
  }
  var ms = o.machines, months = DATA.months || [];
  $('#operNote').textContent = '1日' + o.shift + '時間（2直）× 稼働日数 が分母';

  // ── 機械 × 月 ──
  var h = '<thead><tr><th>機械</th>';
  months.forEach(function (m) { h += '<th class="num">' + m + '月</th>'; });
  h += '<th class="num">年間</th></tr></thead><tbody>';

  // 稼働日数の行（分母の内訳が分かるように）
  h += '<tr class="sub2"><th class="rh">稼働日数</th>';
  var workAll = 0;
  months.forEach(function (m) {
    var v = operMonth(m);
    var w = v ? v.work : 0;
    workAll += w;
    h += '<td class="num">' + w + '日</td>';
  });
  h += '<td class="num">' + workAll + '日</td></tr>';

  ms.concat(['全機械']).forEach(function (mc) {
    var all = mc === '全機械';
    var list = all ? ms : [mc];
    h += '<tr' + (all ? ' class="total"' : '') + '><th class="rh">' + esc(mc) + '</th>';
    var th = 0, tb = 0;
    months.forEach(function (m) {
      var v = operMonth(m);
      var hh = v ? list.reduce(function (s, k) { return s + (v.tot[k] || 0); }, 0) : 0;
      var b = operBase(m, list.length);
      th += hh; tb += b;
      var pct = b ? hh / b * 100 : null;
      h += '<td class="num"' + heat(pct) + ' title="' + esc(mc) + ' ' + m + '月　有効時間 '
        + hh.toFixed(1) + 'h ／ 就業可能 ' + b.toFixed(0) + 'h（'
        + (v ? v.work : 0) + '日）">' + rate(hh, b) + '</td>';
    });
    h += '<td class="num yr" title="' + esc(mc) + ' 年間　有効時間 ' + th.toFixed(1)
      + 'h ／ 就業可能 ' + tb.toFixed(0) + 'h">' + rate(th, tb) + '</td></tr>';
  });
  $('#operSum').innerHTML = h + '</tbody>';

  renderOperDayControls();
  renderOperDay();
}

function renderOperDayControls() {
  var sel = $('#selMonthO'), months = DATA.months || [];
  if (months.indexOf(S.monthO) < 0) { S.monthO = months[months.length - 1] || 1; }
  sel.innerHTML = '';
  months.forEach(function (m) {
    var v = operMonth(m);
    var o = el('option', null, DATA.year + '年 ' + m + '月'
      + (v ? '（稼働 ' + v.work + '日）' : ''));
    o.value = m;
    if (m === S.monthO) { o.selected = true; }
    sel.appendChild(o);
  });
  sel.onchange = function () { S.monthO = +sel.value; renderOperDay(); };
  $$('#pillOper button').forEach(function (b) {
    b.className = b.dataset.o === S.oper ? 'on' : '';
  });
}

/* 選んだ月の 日 × 機械 */
function renderOperDay() {
  var o = oper(), v = operMonth(S.monthO);
  if (!o || !v) { $('#operDay').innerHTML = ''; return; }
  var ms = o.machines, rates = S.oper === 'rate';
  var last = new Date(DATA.year, S.monthO, 0).getDate();
  $('#operDayNote').textContent = DATA.year + '年' + S.monthO + '月　稼働 ' + v.work
    + '日 × ' + o.shift + '時間 = ' + (v.work * o.shift) + '時間';

  var h = '<thead><tr><th>日</th><th>曜</th><th>区分</th>';
  ms.forEach(function (m) { h += '<th class="num">' + esc(m) + '</th>'; });
  h += '<th class="num">全機械</th></tr></thead><tbody>';

  var sum = {}, shown = {};
  for (var d = 1; d <= last; d++) {
    var w = new Date(DATA.year, S.monthO - 1, d).getDay();
    var why = v.off[String(d)] || (w === 0 ? '日曜' : w === 6 ? '土曜' : '');
    var hours = v.days[String(d)] || {};
    var work = !why;
    h += '<tr class="' + (why ? 'off' : '') + '"><td class="num">' + d + '</td>'
      + '<td class="c' + (w === 0 ? ' sun' : w === 6 ? ' sat' : '') + '">' + WDAY[w] + '</td>'
      + '<td>' + esc(why || '稼働日') + '</td>';
    var day = 0;
    ms.forEach(function (m) {
      var x = hours[m] || 0;
      sum[m] = (sum[m] || 0) + x;
      shown[m] = (shown[m] || 0) + x;
      day += x;
      h += '<td class="num"' + (work ? heat(x ? x / o.shift * 100 : 0) : '') + '>'
        + (x ? (rates ? rate(x, o.shift) : x.toFixed(2) + 'h') : '－') + '</td>';
    });
    h += '<td class="num yr">'
      + (day ? (rates ? rate(day, o.shift * ms.length) : day.toFixed(2) + 'h') : '－')
      + '</td></tr>';
  }

  // 日付が取れなかった行（日報のブロックに日付が無く、直前の明細も無いとき）
  var extra = {}, any = false;
  ms.forEach(function (m) {
    extra[m] = Math.round(((v.tot[m] || 0) - (shown[m] || 0)) * 100) / 100;
    if (Math.abs(extra[m]) > 0.005) { any = true; }
  });
  if (any) {
    h += '<tr class="off"><td class="c" colspan="3">日付不明</td>';
    ms.forEach(function (m) {
      h += '<td class="num">' + (extra[m] ? extra[m].toFixed(2) + 'h' : '－') + '</td>';
    });
    h += '<td class="num yr">'
      + ms.reduce(function (s, m) { return s + extra[m]; }, 0).toFixed(2) + 'h</td></tr>';
  }

  h += '<tr class="total"><td class="c" colspan="3">月合計</td>';
  var base = o.shift * v.work;
  ms.forEach(function (m) {
    var t = v.tot[m] || 0;
    h += '<td class="num" title="有効時間 ' + t.toFixed(1) + 'h ／ 就業可能 '
      + base.toFixed(0) + 'h">' + (rates ? rate(t, base) : t.toFixed(1) + 'h') + '</td>';
  });
  var tall = ms.reduce(function (s, m) { return s + (v.tot[m] || 0); }, 0);
  h += '<td class="num yr">'
    + (rates ? rate(tall, base * ms.length) : tall.toFixed(1) + 'h') + '</td></tr>';
  $('#operDay').innerHTML = h + '</tbody>';
}


/* ── 損紙率・予備率 ──
     出庫枚数 = 実印刷枚数 ＋ 基準印刷予備（倉庫から出した紙）
     損紙率   = やれ枚数 ÷ 出庫枚数
                やれ枚数 = 出庫枚数からはみ出して使ってしまった紙（損紙）
     予備率   = 基準印刷予備 ÷ 実印刷枚数
                基準印刷予備 = 出庫枚数 − 実印刷枚数
   出庫枚数は用紙を出すたびに1行へまとめて入るので、明細行すべてには入って
   いない。その行にも日付と機械があるので、月別・日別・機械別に足せる。
   2015年より前の日報には出庫枚数の列が無く、その年はどちらも出せない。
   1マスは [出庫枚数, やれ枚数, 通し枚数, 実印刷枚数] の形で入っている。 */
var W_OUT = 0, W_YARE = 1, W_TSU = 2, W_PRN = 3;

function waste() { return DATA && DATA.waste; }

function wasteRate(out, yare) {
  return out ? (yare / out * 100).toFixed(2) + '%' : '－';
}

/* 予備率 = 基準印刷予備 ÷ 実印刷枚数 */
function spareRate(v) {
  return v[W_PRN] ? ((v[W_OUT] - v[W_PRN]) / v[W_PRN] * 100).toFixed(2) + '%' : '－';
}

/* 率が高いほど濃くする（max で最も濃い） */
function wheat(pct, max) {
  if (pct == null) { return ''; }
  var k = Math.max(0, Math.min(1, pct / (max || 8)));
  return ' style="background:rgba(153,34,34,' + (0.03 + k * 0.17).toFixed(3) + ')"';
}

/* いま選んでいる表示のマスの濃さ（率のときだけ色を付ける） */
function wasteHeat(v, kind) {
  if (kind === 'rate') {
    return v[W_OUT] ? wheat(v[W_YARE] / v[W_OUT] * 100, 8) : '';
  }
  if (kind === 'spare') {
    return v[W_PRN] ? wheat((v[W_OUT] - v[W_PRN]) / v[W_PRN] * 100, 40) : '';
  }
  return '';
}

function wasteMonth(m) {
  var w = waste();
  return (w && w.months && w.months[String(m)]) || null;
}

/* 表の中のマスを足す。machine を省くと全機械ぶん */
function wasteSumOf(src, machine) {
  var s = [0, 0, 0, 0];
  if (src) {
    (machine ? [machine] : Object.keys(src)).forEach(function (k) {
      if (src[k]) { src[k].forEach(function (v, i) { s[i] += v; }); }
    });
  }
  return s;
}

/* その月・その機械のマス。機械を省くと全機械ぶん */
function wasteOf(m, machine) {
  var v = wasteMonth(m);
  return wasteSumOf(v && v.tot, machine);
}

/* マスの表示（損紙率／予備率／やれ枚数／出庫枚数） */
function wasteCell(v, kind) {
  if (kind === 'spare') { return spareRate(v); }
  if (kind === 'yare') { return v[W_YARE] ? num(v[W_YARE]) : '－'; }
  if (kind === 'out') { return v[W_OUT] ? num(v[W_OUT]) : '－'; }
  return wasteRate(v[W_OUT], v[W_YARE]);
}

/* マスに乗せたときの吹き出し */
function wasteTip(v) {
  return '出庫 ' + num(v[W_OUT]) + ' ／ 実印刷 ' + num(v[W_PRN])
    + ' ／ 基準予備 ' + num(v[W_OUT] - v[W_PRN]) + ' ／ やれ ' + num(v[W_YARE])
    + ' ／ 通し ' + num(v[W_TSU]);
}

function renderWaste() {
  var w = waste();
  if (!w || !w.machines.length) {
    $('#wasteSum').innerHTML = '<tbody><tr><td class="empty">'
      + 'この年の稼動日報に「やれ枚数」が入っていないため、損紙率は出せません。'
      + '</td></tr></tbody>';
    $('#wasteDay').innerHTML = '';
    $('#wasteRank').innerHTML = '';
    $('#wasteNote').textContent = '';
    return;
  }
  var ms = w.machines, months = DATA.months || [], kind = S.waste;
  var yr = [0, 0, 0, 0];
  months.forEach(function (m) { wasteOf(m).forEach(function (v, i) { yr[i] += v; }); });
  // 出庫枚数の無い年（2015年より前）は、その旨を出す
  // 出庫枚数の入っている案件が少なすぎる年（2015年より前）は数字を出さない
  var cover = w.cover == null ? 1 : w.cover;
  var enough = yr[W_OUT] && cover >= 0.5;
  $('#wasteNote').textContent = enough
    ? '年間　損紙率 ' + wasteRate(yr[W_OUT], yr[W_YARE]) + '（やれ '
      + num(yr[W_YARE]) + ' ÷ 出庫 ' + num(yr[W_OUT]) + '）　／　予備率 '
      + spareRate(yr) + '（基準予備 ' + num(yr[W_OUT] - yr[W_PRN]) + ' ÷ 実印刷 '
      + num(yr[W_PRN]) + '）'
      + (cover < 0.95 ? '　※出庫枚数の入っている案件 '
                        + (cover * 100).toFixed(0) + '% ぶんです' : '')
    : 'この年の稼動日報には出庫枚数がほとんど入っていないため、損紙率・予備率は'
      + '出せません（出庫枚数のある案件 ' + (cover * 100).toFixed(1) + '%）';
  if (!enough) {
    var msg = '<tbody><tr><td class="empty">'
      + 'この年の稼動日報には出庫枚数の列がほとんど入っていないため、'
      + '損紙率・予備率は出せません。<br>出庫枚数は2016年以降の日報に入っています。'
      + '</td></tr></tbody>';
    $('#wasteSum').innerHTML = msg;
    $('#wasteDay').innerHTML = '';
    $('#wasteRank').innerHTML = '';
    $('#wasteDayNote').textContent = '';
    $('#wasteRankNote').textContent = '';
    return;
  }

  var h = '<thead><tr><th>機械</th>';
  months.forEach(function (m) { h += '<th class="num">' + m + '月</th>'; });
  h += '<th class="num">年間</th></tr></thead><tbody>';
  ms.concat(['全機械']).forEach(function (mc) {
    var all = mc === '全機械';
    h += '<tr' + (all ? ' class="total"' : '') + '><th class="rh">' + esc(mc) + '</th>';
    var t = [0, 0, 0, 0];
    months.concat([0]).forEach(function (m) {
      var v = m ? wasteOf(m, all ? null : mc) : t;
      if (m) { v.forEach(function (x, i) { t[i] += x; }); }
      h += '<td class="num' + (m ? '' : ' yr') + '"' + (m ? wasteHeat(v, kind) : '')
        + ' title="' + esc(mc) + ' ' + (m ? m + '月' : '年間') + '　' + wasteTip(v) + '">'
        + wasteCell(v, kind) + '</td>';
    });
    h += '</tr>';
  });
  $('#wasteSum').innerHTML = h + '</tbody>';

  renderWasteControls();
  renderWasteDay();
  renderWasteRank();
}

function renderWasteControls() {
  var months = DATA.months || [];
  if (months.indexOf(S.monthWD) < 0) { S.monthWD = months[months.length - 1] || 1; }
  fillMonths($('#selMonthWD'), S.monthWD, false, function (m) {
    S.monthWD = m; renderWasteDay();
  });
  fillMonths($('#selMonthW'), S.monthW, true, function (m) {
    S.monthW = m; renderWasteRank();
  });
  $$('#pillWasteDay button').forEach(function (b) {
    b.className = b.dataset.d === S.wasteDay ? 'on' : '';
  });
  $$('#pillWaste button').forEach(function (b) {
    b.className = b.dataset.w === S.waste ? 'on' : '';
  });
  $$('#pillWasteSort button').forEach(function (b) {
    b.className = b.dataset.s === S.wasteSort ? 'on' : '';
  });
}

/* 日ごとの損紙率（日付順・機械順） */
function renderWasteDay() {
  var w = waste(), v = wasteMonth(S.monthWD);
  if (!w || !v) { $('#wasteDay').innerHTML = ''; return; }
  var ms = w.machines, kind = S.wasteDay;
  var last = new Date(DATA.year, S.monthWD, 0).getDate();
  var tot = wasteSumOf(v.tot);
  $('#wasteDayNote').textContent = DATA.year + '年' + S.monthWD + '月　損紙率 '
    + wasteRate(tot[W_OUT], tot[W_YARE]) + '　予備率 ' + spareRate(tot)
    + '（出庫 ' + num(tot[W_OUT]) + ' ／ 実印刷 ' + num(tot[W_PRN])
    + ' ／ やれ ' + num(tot[W_YARE]) + '）';

  var h = '<thead><tr><th>日</th><th>曜</th>';
  ms.forEach(function (m) { h += '<th class="num">' + esc(m) + '</th>'; });
  h += '<th class="num">全機械</th></tr></thead><tbody>';

  var shown = {};
  ms.forEach(function (m) { shown[m] = [0, 0, 0, 0]; });
  for (var d = 1; d <= last; d++) {
    var wd = new Date(DATA.year, S.monthWD - 1, d).getDay();
    var day = v.days[String(d)] || {};
    var any = Object.keys(day).length > 0;
    h += '<tr class="' + (any ? '' : 'off') + '"><td class="num">' + d + '</td>'
      + '<td class="c' + (wd === 0 ? ' sun' : wd === 6 ? ' sat' : '') + '">'
      + WDAY[wd] + '</td>';
    ms.forEach(function (m) {
      var x = day[m];
      if (x) { x.forEach(function (q, i) { shown[m][i] += q; }); }
      h += '<td class="num"' + (x ? wasteHeat(x, kind) : '')
        + ' title="' + (x ? wasteTip(x) : '') + '">'
        + (x ? wasteCell(x, kind) : '－') + '</td>';
    });
    var all = wasteSumOf(day);
    h += '<td class="num yr">' + (any ? wasteCell(all, kind) : '－') + '</td></tr>';
  }

  // 日報の行に日付が無かった分（月合計と日ごとの和の差）
  var extra = {}, some = false;
  ms.forEach(function (m) {
    var t = v.tot[m] || [0, 0, 0, 0];
    extra[m] = t.map(function (q, i) { return q - shown[m][i]; });
    if (extra[m].some(function (q) { return q; })) { some = true; }
  });
  if (some) {
    h += '<tr class="off"><td class="c" colspan="2">日付不明</td>';
    ms.forEach(function (m) {
      h += '<td class="num">'
        + (extra[m].some(function (q) { return q; }) ? wasteCell(extra[m], kind) : '－')
        + '</td>';
    });
    h += '<td class="num yr">' + wasteCell(wasteSumOf(extra), kind) + '</td></tr>';
  }

  h += '<tr class="total"><td class="c" colspan="2">月合計</td>';
  ms.forEach(function (m) {
    h += '<td class="num">' + wasteCell(v.tot[m] || [0, 0, 0, 0], kind) + '</td>';
  });
  h += '<td class="num yr">' + wasteCell(tot, kind) + '</td></tr>';
  $('#wasteDay').innerHTML = h + '</tbody>';
}

/* 損紙の多い案件（管理番号ごと。統合後の件数で数える） */
var WASTETOP = 100;
function renderWasteRank() {
  // 数える範囲は上の「機械別 月間」の表とそろえる。r.ob は「その管理番号に
  // 出庫枚数があるか」で、部署をまたいで判定してある。r.out > 0 で絞ると、
  // 同じ管理番号が部署で分かれている案件のやれ枚数が落ちて月合計と合わなくなる。
  var rows = recs(S.monthW || null, null).filter(function (r) {
    return r.yare > 0 && r.ob;
  });
  var byRate = S.wasteSort === 'rate';
  rows.sort(function (a, b) {
    if (byRate) {
      // 出庫枚数が別部署の行に入っている案件は率を出せないので後ろへ
      var ra = a.out ? a.yare / a.out : -1, rb = b.out ? b.yare / b.out : -1;
      return rb - ra || b.yare - a.yare;
    }
    return b.yare - a.yare || b.out - a.out;
  });
  var top = rows.slice(0, WASTETOP);
  $('#wasteRankNote').textContent = (S.monthW ? S.monthW + '月' : '年間（全月）')
    + '　損紙のあった案件 ' + num(rows.length) + ' 件'
    + (rows.length > WASTETOP ? ' のうち上位 ' + WASTETOP + ' 件' : '');

  var h = '<thead><tr><th>順位</th><th>管理番号</th><th>ｸﾗｲｱﾝﾄ名</th><th>品名</th>'
    + '<th>営業部</th><th>機械</th>'
    + (S.monthW ? '' : '<th class="c">月</th>')
    + '<th class="num">出庫枚数</th><th class="num">やれ枚数</th>'
    + '<th class="num">損紙率</th></tr></thead><tbody>';
  top.forEach(function (r, i) {
    var pct = r.out ? r.yare / r.out * 100 : null;
    h += '<tr><td class="num">' + (i + 1) + '</td><td>' + esc(r.no) + '</td>'
      + '<td>' + esc(r.client) + '</td><td>' + esc(r.name) + '</td>'
      + '<td>' + esc(shortDept(r.dept)) + '</td>'
      + '<td>' + esc((r.machines || []).map(shortMachine).join(' / ')) + '</td>'
      + (S.monthW ? '' : '<td class="c">' + r.m + '月</td>')
      + '<td class="num">' + (r.out ? num(r.out) : '－') + '</td>'
      + '<td class="num">' + num(r.yare) + '</td>'
      + '<td class="num"' + (pct == null ? '' : wheat(pct)) + '>'
      + wasteRate(r.out, r.yare) + '</td></tr>';
  });
  var to = rows.reduce(function (s, r) { return s + r.out; }, 0);
  var ty = rows.reduce(function (s, r) { return s + r.yare; }, 0);
  h += '<tr class="total"><td colspan="' + (S.monthW ? 6 : 7) + '">合計（損紙のあった案件）</td>'
    + '<td class="num">' + num(to) + '</td><td class="num">' + num(ty) + '</td>'
    + '<td class="num">' + wasteRate(to, ty) + '</td></tr>';
  $('#wasteRank').innerHTML = h + '</tbody>';
}

/* 明細の機械名も、損紙率の表と同じ短い名前にそろえる
   （build.py の short_machine と同じ切り方） */
function shortMachine(s) {
  var t = String(s == null ? '' : s)
    .replace(/稼[働動](?:日報|実績)/g, '')
    .replace(/^\s*新[・･\-_\s]*(?:(?:20)?\d{2}\s*年?[・･\-_\s]*)?/, '')
    .replace(/^\s*(?:20)?\d{2}\s*年?[・･\-_\s]*/, '')
    .replace(/^[\s・･\-_]+|[\s・･\-_]+$/g, '');
  return t || String(s == null ? '' : s).trim();
}

/* 日報明細 ── 日報作成に使った内容をそのまま表示する */
function rawRows() {
  var out = [];
  DATA.records.forEach(function (r) {
    if (r.m !== S.monthR) return;
    r.det.forEach(function (d) {
      if (S.machineR !== '全機械' && d['__機械'] !== S.machineR) return;
      out.push({ d: d, dept: r.dept, no: r.no });
    });
  });
  var q = S.qr.trim().toLowerCase();
  if (q) {
    out = out.filter(function (x) {
      return Object.keys(x.d).some(function (k) {
        return String(x.d[k]).toLowerCase().indexOf(q) >= 0;
      });
    });
  }
  out.sort(function (a, b) {
    return String(a.d['日付']).localeCompare(String(b.d['日付']))
      || String(a.d['管理番号']).localeCompare(String(b.d['管理番号']));
  });
  return out;
}

function machines() {
  var set = {};
  DATA.records.forEach(function (r) { r.machines.forEach(function (m) { set[m] = 1; }); });
  return Object.keys(set).sort();
}

function renderRawControls() {
  fillMonths($('#selMonthR'), S.monthR, false, function (m) { S.monthR = m; renderRaw(); });
  var sm = $('#selMachineR'); sm.innerHTML = '';
  ['全機械'].concat(machines()).forEach(function (m) {
    var o = el('option', null, m);
    o.value = m;
    if (m === S.machineR) o.selected = true;
    sm.appendChild(o);
  });
  sm.onchange = function () { S.machineR = sm.value; renderRaw(); };
  $('#qr').value = S.qr;
  $('#qr').oninput = function () { S.qr = this.value; renderRaw(); };
}

function renderRaw() {
  var rows = rawRows();
  var tsu = rows.reduce(function (s, x) { return s + (+x.d['通し枚数'] || 0); }, 0);
  $('#rTitle').innerHTML = DATA.year + '年' + S.monthR + '月　' + esc(S.machineR)
    + ' <small>明細 ' + rows.length + '行　通し枚数 ' + num(tsu) + '（統合前の生の行）</small>';
  var t = $('#raw');
  if (!rows.length) {
    t.innerHTML = '<tbody><tr><td class="empty">該当するデータがありません。</td></tr></tbody>';
  } else {
    var cols = ['__機械'].concat(DATA.cols).filter(function (c) {
      return rows.some(function (x) { return x.d[c] !== undefined && x.d[c] !== null && x.d[c] !== ''; });
    });
    var h = '<thead><tr><th>部署</th>' + cols.map(function (c) {
      return '<th>' + esc(c === '__機械' ? '機械' : c) + '</th>';
    }).join('') + '</tr></thead><tbody>';
    rows.forEach(function (x) {
      h += '<tr><td class="c">' + esc(shortDept(x.dept)) + '</td>' + cols.map(function (c) {
        return '<td class="' + (typeof x.d[c] === 'number' ? 'num' : '') + '">'
          + rawVal(c, x.d[c]) + '</td>';
      }).join('') + '</tr>';
    });
    t.innerHTML = h + '</tbody>';
  }

  // 日次集計ブロック
  var dd = (DATA.daily || []).filter(function (x) {
    return x.m === S.monthR && (S.machineR === '全機械' || x.machine === S.machineR);
  });
  var dt = $('#dailyT');
  if (!dd.length) {
    dt.innerHTML = '<tbody><tr><td class="empty">この月の日次集計ブロックはありません。</td></tr></tbody>';
    return;
  }
  var keys = [];
  dd.forEach(function (x) {
    Object.keys(x.vals).forEach(function (k) {
      if (k !== '日付' && keys.indexOf(k) < 0) keys.push(k);   // 日付は専用列で出す
    });
  });
  var dh = '<thead><tr><th>機械</th><th>日付</th><th>項目</th>'
    + keys.map(function (k) { return '<th>' + esc(k) + '</th>'; }).join('') + '</tr></thead><tbody>';
  dd.forEach(function (x) {
    dh += '<tr><td>' + esc(x.machine) + '</td><td class="c">'
      + esc((x.date || '').slice(5).replace('-', '/')) + '</td><td>' + esc(x.label) + '</td>'
      + keys.map(function (k) {
        return '<td class="' + (typeof x.vals[k] === 'number' ? 'num' : '') + '">'
          + rawVal(k, x.vals[k]) + '</td>';
      }).join('') + '</tr>';
  });
  dt.innerHTML = dh + '</tbody>';
}

function csvRaw() {
  var rows = rawRows();
  var cols = ['__機械'].concat(DATA.cols).filter(function (c) {
    return rows.some(function (x) { return x.d[c] !== undefined && x.d[c] !== null && x.d[c] !== ''; });
  });
  var out = [['部署'].concat(cols.map(function (c) { return c === '__機械' ? '機械' : c; }))];
  rows.forEach(function (x) {
    out.push([x.dept].concat(cols.map(function (c) {
      return x.d[c] === undefined || x.d[c] === null ? '' : x.d[c];
    })));
  });
  csv(out, DATA.year + '年' + S.monthR + '月_日報明細_' + S.machineR + '.xlsx');
}

/* 元ファイル・検証 */
function renderSrc() {
  var h = '<thead><tr><th>ファイル名</th><th>機械</th><th>収録月</th><th class="num">明細行</th></tr></thead><tbody>';
  DATA.files.forEach(function (f) {
    h += '<tr><td>' + esc(f.name) + '</td><td>' + esc(f.machine) + '</td><td>'
      + (f.months.length ? f.months.map(function (m) { return m + '月'; }).join('、') : '（対象年のシートなし）')
      + '</td><td class="num">' + num(f.rows) + '</td></tr>';
  });
  h += '</tbody>';
  $('#files').innerHTML = h;
  var srcs = DATA.sources || [DATA.source];
  var sk = DATA.skipped || [];
  $('#srcPath').innerHTML = '<b>読込元フォルダ</b><br>　' + srcs.map(esc).join('<br>　')
    + (sk.length ? '<br><b>このPCに無かったフォルダ（読み飛ばし）</b><br>　'
        + sk.map(esc).join('<br>　') : '')
    + '<br>読み込みは一時フォルダへコピーして行うため、元のExcelファイルは変更されません。'
    + '同じファイルが複数のフォルダにある場合は1つだけ数えます。';

  var s = '<thead><tr><th class="c">月</th><th>営業部</th><th class="num">明細行</th><th class="num">統合件数</th>'
    + '<th class="num">明細の通し数</th><th class="num">統合後の通し数</th><th class="c">判定</th></tr></thead><tbody>';
  DATA.stats.forEach(function (x) {
    var ok = x.tsuDetail === x.tsuGroups;
    s += '<tr><td class="c">' + x.m + '月</td><td>' + esc(x.dept) + '</td>'
      + '<td class="num">' + num(x.detail) + '</td><td class="num">' + num(x.groups) + '</td>'
      + '<td class="num">' + num(x.tsuDetail) + '</td><td class="num">' + num(x.tsuGroups) + '</td>'
      + '<td class="c" style="color:' + (ok ? '#1e7a4d' : '#b91c1c') + '">' + (ok ? '一致' : '不一致') + '</td></tr>';
  });
  s += '</tbody>';
  $('#stats').innerHTML = s;
}

/* CSV 出力（Excel で文字化けしないよう BOM 付き） */
/* 表を Excel ファイル(.xlsx)にして保存する。
   見出しに色と罫線を付け、上で固定してオートフィルタを付ける。数値は数値の
   まま（3桁区切り）入るので、Excel でそのまま並べ替えや集計ができる。
   組み立てはサーバー側（server.py + xlsx.py）で行う。 */
function csv(rows, filename) {
  var base = filename.replace(/\.(csv|xlsx)$/, '');
  fetch('/api/xlsx', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ rows: rows, sheet: base })
  }).then(function (r) {
    if (!r.ok) { throw new Error('サーバーが受け付けませんでした'); }
    return r.blob();
  }).then(function (blob) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = base + '.xlsx';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }).catch(function (e) {
    alert('Excelファイルを作れませんでした。起動.bat の黒い画面が'
      + '開いたままか確かめてください。\n' + e);
  });
}

function csvCmp() {
  var prev = yearSummary(DATA.year - 1), months = DATA.months || [];
  if (!prev || !prev.byDept) { return; }
  var cur = curByDept(), pb = prev.byDept;
  var rows = [['部署', '通し数', '前年', '増減', '前年比',
               '件数', '前年', '前年比（件数）']];
  DATA.depts.concat(['合計']).forEach(function (d) {
    var ds = d === '合計' ? DATA.depts : [d];
    var ct = 0, pt = 0, cc = 0, pc = 0;
    ds.forEach(function (x) {
      ct += pick(cur, x, months, 0); pt += pick(pb, x, months, 0);
      cc += pick(cur, x, months, 1); pc += pick(pb, x, months, 1);
    });
    rows.push([d, ct, pt, ct - pt, ratio(ct, pt), cc, pc, ratio(cc, pc)]);
  });
  csv(rows, DATA.year + '年_前年比較_部署別.xlsx');
}

/* 月別 × 部署別（画面の表と同じ、部署ごとに区分4段） */
function csvYear() {
  var kinds = matrixKinds(), yoy = kinds.length > 2;
  var prev = yoy ? yearSummary(DATA.year - 1) : null;
  var cur = yoy ? curByDept() : null, pb = prev ? prev.byDept : null;
  var have = DATA.months || [];
  var rows = [['営業部', '区分']
    .concat(MONTHS.map(function (m) { return m + '月'; })).concat(['年間合計'])];
  DATA.depts.concat(['合計']).forEach(function (d) {
    var ds = d === '合計' ? DATA.depts : [d];
    kinds.forEach(function (kind) {
      rows.push([d, kind.label].concat(MONTHS.concat([0]).map(function (m) {
        var year = !m;
        if (kind.k === 'tsu' || kind.k === 'cnt') {
          return cell(d, year ? null : m)[kind.k === 'tsu' ? 0 : 1];
        }
        var ms = year ? have : (have.indexOf(m) >= 0 ? [m] : null);
        if (!ms || !ms.length) { return ''; }
        var c = 0, p = 0;
        ds.forEach(function (x) { c += pick(cur, x, ms, 0); p += pick(pb, x, ms, 0); });
        return kind.k === 'pre' ? p : kind.k === 'yoy' ? ratio(c, p) : c - p;
      })));
    });
  });
  csv(rows, DATA.year + '年_月別営業部別.xlsx');
}

/* 稼働率（機械 × 月） */
function csvOper() {
  var o = oper();
  if (!o || !o.machines.length) { return; }
  var months = DATA.months || [];
  var head = ['機械'].concat(months.map(function (m) { return m + '月'; })).concat(['年間']);
  var rows = [head];
  rows.push(['稼働日数'].concat(months.map(function (m) {
    var v = operMonth(m); return v ? v.work : 0;
  })).concat([months.reduce(function (s, m) {
    var v = operMonth(m); return s + (v ? v.work : 0);
  }, 0)]));
  o.machines.concat(['全機械']).forEach(function (mc) {
    var list = mc === '全機械' ? o.machines : [mc];
    var th = 0, tb = 0;
    var line = [mc].concat(months.map(function (m) {
      var v = operMonth(m);
      var hh = v ? list.reduce(function (s, k) { return s + (v.tot[k] || 0); }, 0) : 0;
      var b = operBase(m, list.length);
      th += hh; tb += b;
      return rate(hh, b);
    }));
    rows.push(line.concat([rate(th, tb)]));
  });
  csv(rows, DATA.year + '年_機械別_月間稼働率.xlsx');
}

/* 稼働率（日 × 機械） */
function csvOperDay() {
  var o = oper(), v = operMonth(S.monthO);
  if (!o || !v) { return; }
  var ms = o.machines, last = new Date(DATA.year, S.monthO, 0).getDate();
  var rows = [['日', '曜', '区分'].concat(ms).concat(['全機械（有効時間）'])];
  for (var d = 1; d <= last; d++) {
    var w = new Date(DATA.year, S.monthO - 1, d).getDay();
    var why = v.off[String(d)] || (w === 0 ? '日曜' : w === 6 ? '土曜' : '');
    var hh = v.days[String(d)] || {};
    rows.push([d, WDAY[w], why || '稼働日'].concat(ms.map(function (m) {
      return hh[m] || 0;
    })).concat([ms.reduce(function (s, m) { return s + (hh[m] || 0); }, 0)]));
  }
  var base = o.shift * v.work;
  rows.push(['月合計', '', '稼働 ' + v.work + '日'].concat(ms.map(function (m) {
    return v.tot[m] || 0;
  })).concat([ms.reduce(function (s, m) { return s + (v.tot[m] || 0); }, 0)]));
  rows.push(['稼働率', '', '分母 ' + base + '時間'].concat(ms.map(function (m) {
    return rate(v.tot[m] || 0, base);
  })).concat([rate(ms.reduce(function (s, m) { return s + (v.tot[m] || 0); }, 0),
                   base * ms.length)]));
  csv(rows, DATA.year + '年' + S.monthO + '月_日ごとの稼働率.xlsx');
}

/* 損紙率（機械 × 月） */
function csvWaste() {
  var w = waste();
  if (!w || !w.machines.length) { return; }
  var months = DATA.months || [];
  var rows = [['機械', '区分'].concat(months.map(function (m) { return m + '月'; }))
    .concat(['年間'])];
  w.machines.concat(['全機械']).forEach(function (mc) {
    var all = mc === '全機械';
    var cells = months.map(function (m) { return wasteOf(m, all ? null : mc); });
    var t = wasteSumOf0(cells);
    var col = function (i) {
      return cells.map(function (v) { return v[i]; }).concat([t[i]]);
    };
    rows.push([mc, '出庫枚数'].concat(col(W_OUT)));
    rows.push([mc, '実印刷枚数'].concat(col(W_PRN)));
    rows.push([mc, '基準印刷予備'].concat(cells.map(function (v) {
      return v[W_OUT] - v[W_PRN];
    })).concat([t[W_OUT] - t[W_PRN]]));
    rows.push([mc, 'やれ枚数'].concat(col(W_YARE)));
    rows.push([mc, '損紙率'].concat(cells.map(function (v) {
      return wasteRate(v[W_OUT], v[W_YARE]);
    })).concat([wasteRate(t[W_OUT], t[W_YARE])]));
    rows.push([mc, '予備率'].concat(cells.map(spareRate)).concat([spareRate(t)]));
  });
  csv(rows, DATA.year + '年_機械別_月間損紙率.xlsx');
}

/* 日ごとの損紙率（日付順・機械順） */
function csvWasteDay() {
  var w = waste(), v = wasteMonth(S.monthWD);
  if (!w || !v) { return; }
  var ms = w.machines, last = new Date(DATA.year, S.monthWD, 0).getDate();
  var head = ['日', '曜'];
  ms.concat(['全機械']).forEach(function (m) {
    head.push(m + ' 出庫枚数', m + ' 実印刷枚数', m + ' やれ枚数',
              m + ' 損紙率', m + ' 予備率');
  });
  var rows = [head];
  var line = function (label, wd, get) {
    var r = [label, wd];
    var put = function (x) {
      r.push(x[W_OUT], x[W_PRN], x[W_YARE],
             wasteRate(x[W_OUT], x[W_YARE]), spareRate(x));
    };
    ms.forEach(function (m) { put(get(m) || [0, 0, 0, 0]); });
    put(wasteSumOf0(ms.map(get)));
    return r;
  };
  for (var d = 1; d <= last; d++) {
    var day = v.days[String(d)] || {};
    rows.push(line(d, WDAY[new Date(DATA.year, S.monthWD - 1, d).getDay()],
                   function (m) { return day[m]; }));
  }
  rows.push(line('月合計', '', function (m) { return v.tot[m]; }));
  csv(rows, DATA.year + '年' + S.monthWD + '月_日ごとの損紙率.xlsx');
}

/* マスの配列を足す */
function wasteSumOf0(list) {
  var s = [0, 0, 0, 0];
  list.forEach(function (x) { if (x) { x.forEach(function (v, i) { s[i] += v; }); } });
  return s;
}

/* 損紙の多い案件 */
function csvWasteRank() {
  var rows = recs(S.monthW || null, null).filter(function (r) {
    return r.yare > 0 && r.ob;
  });
  var byRate = S.wasteSort === 'rate';
  rows.sort(function (a, b) {
    if (byRate) {
      var ra = a.out ? a.yare / a.out : -1, rb = b.out ? b.yare / b.out : -1;
      return rb - ra || b.yare - a.yare;
    }
    return b.yare - a.yare || b.out - a.out;
  });
  var one = !!S.monthW;
  var out = [['順位', '管理番号', 'ｸﾗｲｱﾝﾄ名', '品名', '営業部', '機械']
    .concat(one ? [] : ['月']).concat(['出庫枚数', 'やれ枚数', '損紙率'])];
  rows.forEach(function (r, i) {
    out.push([i + 1, r.no, r.client, r.name, r.dept,
      (r.machines || []).map(shortMachine).join(' / ')]
      .concat(one ? [] : [r.m + '月'])
      .concat([r.out, r.yare, wasteRate(r.out, r.yare)]));
  });
  var to = rows.reduce(function (s, r) { return s + r.out; }, 0);
  var ty = rows.reduce(function (s, r) { return s + r.yare; }, 0);
  out.push(['合計', '', '', '', '', ''].concat(one ? [] : [''])
    .concat([to, ty, wasteRate(to, ty)]));
  csv(out, DATA.year + '年' + (one ? S.monthW + '月' : '') + '_損紙の多い案件.xlsx');
}

function csvMonth() {
  var rs = filtered();
  var all = S.dept === '全社', pre = all ? ['営業部'] : [];
  var rows = [pre.concat(['管理番号', 'ｸﾗｲｱﾝﾄ名', '品名', '今年の動向', '無しの場合の代替対策',
    '対策通し数', '営業担当ｺｰﾄﾞ', '印刷日', '色数', '通し数'])];
  rs.forEach(function (r) {
    var m = memoOf(r);
    rows.push((all ? [r.dept] : []).concat([r.no, r.client, r.name, m.trend, m.plan, m.tsu, r.code,
      r.date.replace(/-/g, '/'), r.color, r.tsu]));
  });
  var tot = rs.reduce(function (s, r) { return s + r.tsu; }, 0);
  var pt = rs.reduce(function (s, r) { return s + planTsu(r); }, 0);
  var pad = all ? [''] : [];
  rows.push(pad.concat(['合計', '', rs.length + '件', '', '', pt, '', '', '', tot]));
  rows.push(pad.concat(['差引（通し数 − 対策通し数）', '', '', '', '', tot - pt, '', '', '充足率',
    tot ? (pt / tot * 100).toFixed(1) + '%' : '']));
  csv(rows, DATA.year + '年' + S.month + '月_' + S.dept + '_印刷実績.xlsx');
}

function csvClient() {
  var one = !!S.monthC, rs = clientRows();
  var rows = [['順位', 'ｸﾗｲｱﾝﾄ名', '営業部', '件数', '通し数'].concat(
    one ? ['品名（管理番号 / 通し数）'] : MONTHS.map(function (m) { return m + '月'; }))];
  rs.forEach(function (e, i) {
    rows.push([i + 1, e.name, e.dept, e.cnt, e.tsu].concat(one
      ? [e.items.map(function (x) { return (x.name || '') + '（' + x.no + ' / ' + x.tsu + '）'; }).join(' ／ ')]
      : MONTHS.map(function (m) { return e.months[m] || 0; })));
  });
  csv(rows, DATA.year + '年' + (one ? S.monthC + '月' : '') + '_' + S.deptC + '_得意先別.xlsx');
}

/* 画面切替・その他 */
function switchView(v) {
  S.view = v;
  $$('nav.tabs button').forEach(function (b) { b.classList.toggle('on', b.dataset.view === v); });
  $$('section.view').forEach(function (s) { s.classList.toggle('on', s.id === 'v-' + v); });
}

$$('nav.tabs button').forEach(function (b) { b.onclick = function () { switchView(b.dataset.view); }; });
/* 印刷。いま開いているタブだけが紙に出る（他のタブは隠れている） */
var VIEWNAME = { year: '年間サマリー', month: '月別明細', client: '得意先別',
                 oper: '稼働率', waste: '損紙・予備', raw: '日報明細',
                 src: '元ファイル' };
function setPrintTitle() {
  var t = $('#printTitle');
  if (!t || !DATA) { return; }
  t.innerHTML = esc(DATA.year + '年 稼動日報 印刷実績　' + (VIEWNAME[S.view] || ''))
    + '<small>' + esc('読込 ' + DATA.generated) + '</small>';
}

function printView(hint) {
  setPrintTitle();
  if (hint) {
    say('印刷画面の「送信先（プリンター）」で <b>PDFとして保存</b> を'
      + '選んでください', 6000);
    setTimeout(function () { window.print(); }, 500);
  } else {
    window.print();
  }
}

$('#btnPrint').onclick = function () { printView(false); };
$('#btnPrintTop').onclick = function () { printView(false); };
$('#btnPdf').onclick = function () { printView(true); };
$$('#pillChart button').forEach(function (b) {
  b.onclick = function () { S.chart = b.dataset.c; renderChart(); };
});
$('#selChartDept').onchange = function () {
  S.chartDept = $('#selChartDept').value; renderChart();
};
$('#btnCsvYear').onclick = csvYear;
$('#btnCsvCmp').onclick = csvCmp;
$('#btnCsvMonth').onclick = csvMonth;
$('#btnCsvClient').onclick = csvClient;
$('#btnCsvOper').onclick = csvOper;
$('#btnCsvOperDay').onclick = csvOperDay;
$('#btnCsvWaste').onclick = csvWaste;
$('#btnCsvWasteDay').onclick = csvWasteDay;
$('#btnCsvWasteRank').onclick = csvWasteRank;
$$('#pillWasteDay button').forEach(function (b) {
  b.onclick = function () {
    S.wasteDay = b.dataset.d; renderWasteControls(); renderWasteDay();
  };
});
$$('#pillWaste button').forEach(function (b) {
  b.onclick = function () { S.waste = b.dataset.w; renderWaste(); };
});
$$('#pillWasteSort button').forEach(function (b) {
  b.onclick = function () {
    S.wasteSort = b.dataset.s; renderWasteControls(); renderWasteRank();
  };
});
$$('#pillOper button').forEach(function (b) {
  b.onclick = function () { S.oper = b.dataset.o; renderOperDayControls(); renderOperDay(); };
});
$('#btnCsvRaw').onclick = csvRaw;

$('#btnSetting').onclick = function () {
  fetch('/api/config').then(function (r) { return r.json(); }).then(function (c) {
    var list = Array.isArray(c.src) ? c.src.slice() : (c.src ? [c.src] : []);
    (c.srcAlt || []).forEach(function (x) { if (list.indexOf(x) < 0) list.push(x); });
    $('#inSrc').value = list.join('\n');
    $('#inYear').value = c.year || '';
    $('#inClosed').value = (c.closed || []).join('\n');
    // 社内サーバーで共有しているときは、読むフォルダを画面から変えられない
    // （サーバー上の好きなフォルダを読めてしまわないようにするため）
    var ro = !!c.shared;
    $('#inSrc').readOnly = ro;
    $('#inYear').readOnly = ro;
    $('#inClosed').readOnly = ro;
    $('#btnPick').style.display = ro ? 'none' : '';
    $('#dlgOk').style.display = ro ? 'none' : '';
    $('#sharedNote').style.display = ro ? '' : 'none';
    // U: が割り当てられていないPC向けの予備パスと、実際に読んだフォルダを知らせる
    var alt = (c.srcAlt || []).map(esc).join('<br>　　');
    $('#altNote').innerHTML = (alt
      ? '上のフォルダが見つからないときは、次のパスも自動で試します（U: が割り当てられていないPC向け）:<br>　　' + alt + '<br>'
      : '') + (DATA ? '今回読み込んだフォルダ: <b>' + esc(DATA.source) + '</b>' : '');
    $('#dlg').showModal();
  }).catch(function () {
    $('#inSrc').value = DATA ? DATA.source : '';
    $('#inYear').value = '';
    $('#inClosed').value = '';
    $('#dlg').showModal();
  });
};
$('#btnPick').onclick = function () {
  // サーバー（このPC）側でフォルダ選択ダイアログを出してもらう
  var b = $('#btnPick'), label = b.textContent;
  b.disabled = true; b.textContent = '選んでください…';
  function back() { b.disabled = false; b.textContent = label; }
  var lines = $('#inSrc').value.split('\n').map(function (x) { return x.trim(); }).filter(Boolean);
  fetch('/api/pick', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ initial: lines[0] || '' })
  }).then(function (r) { return r.json(); }).then(function (k) {
    back();
    if (!k.ok) { alert(k.error); return; }
    if (!k.path) { return; }                       // 選ばずに閉じた
    if (lines.indexOf(k.path) < 0) { lines.push(k.path); }
    $('#inSrc').value = lines.join('\n');
  }).catch(function (e) {
    back();
    alert('サーバーに接続できませんでした。起動.bat から開き直してください。\n' + e);
  });
};
$('#dlgCancel').onclick = function () { $('#dlg').close(); };
$('#dlgOk').onclick = function () {
  $('#dlg').close();
  doRebuild($('#inSrc').value, $('#inYear').value, $('#inClosed').value);
};
$('#btnRebuild').onclick = function () { doRebuild(); };

function doRebuild(src, year, closed) {
  var b = $('#btnRebuild'); b.disabled = true; b.textContent = '読込中…';
  fetch('/api/rebuild', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ src: src, year: year, closed: closed })
  }).then(function (r) { return r.json(); }).then(function (j) {
    b.disabled = false; b.textContent = 'データ更新';
    if (!j.ok) { alert('読み込みに失敗しました:\n' + j.error); return; }
    S.year = null;                       // 年の一覧が変わるので選び直す
    var chartTimer = null;
window.addEventListener('resize', function () {
  // 幅に合わせて目盛りごと引き直す（続けて変わるので少し待つ）
  clearTimeout(chartTimer);
  chartTimer = setTimeout(function () { if (DATA) { renderChart(); } }, 200);
});

boot();
  }).catch(function (e) {
    b.disabled = false; b.textContent = 'データ更新';
    alert('サーバーに接続できませんでした。起動.bat から開き直してください。\n' + e);
  });
}

boot();
