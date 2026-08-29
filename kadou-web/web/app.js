/* 稼動日報 印刷実績ビューア
   data.json（build.py が生成）を読み、月別・営業部別に表示する。
   記入欄の内容は /api/memo（memo.json）に保存する。元の .xls は読み取りのみ。 */
'use strict';

var YEARS = null, DATA = null, MEMO = {},
    DEPTKEY = { '本社営業部': 'hq', '東京営業部': 'tk', '池袋営業部': 'ik',
                '生産管理部（工務）': 'km', 'その他': 'ot' };
var S = { view: 'year', year: null, month: null, dept: '全社', deptC: '全社', monthC: 0,
          monthR: null, machineR: '全機械', q: '', qc: '', qr: '',
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
  Promise.all([
    fetch('years.json?t=' + Date.now()).then(function (r) { return r.ok ? r.json() : null; })
      .catch(function () { return null; }),
    fetch('/api/memo').then(function (r) { return r.ok ? r.json() : {}; })
      .catch(function () { return localMemo(); })
  ]).then(function (a) {
    YEARS = a[0];
    MEMO = a[1] || {};
    indexMemo();
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
  e.classList.add('on');
  clearTimeout(sayTimer);
  sayTimer = setTimeout(function () {
    e.classList.remove('on');
    e.textContent = '保存しました';
  }, ms || 2000);
}

function localMemo() { try { return JSON.parse(localStorage.getItem('kadouMemo') || '{}'); } catch (e) { return {}; } }

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
    if (p.length === 4) { MEMOIDX[key2(p[0], p[1], p[3])] = MEMO[k]; }
  });
}

function memoOf(r) {
  return MEMO[key(r)] || MEMOIDX[key2(DATA.year, r.m, r.no)]
    || { trend: '', plan: '', tsu: '' };
}
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
  renderClientControls(); renderClient(); renderRawControls(); renderRaw(); renderSrc();
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
  $('#cmpMonthCard').style.display = ok ? '' : 'none';
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

  // ── 月別（部署別の前年比） ──
  var t = '<thead><tr><th>部署</th>';
  months.forEach(function (m) { t += '<th class="num">' + m + '月</th>'; });
  t += '<th class="num">合計</th></tr></thead><tbody>';
  rows.forEach(function (d) {
    var all = d === '合計';
    var ds = all ? DATA.depts : [d];
    t += '<tr' + (all ? ' class="total"' : '') + '><th class="rh">' + esc(d) + '</th>';
    months.concat(['計']).forEach(function (m) {
      var ms = m === '計' ? months : [m];
      var c = 0, p = 0;
      ds.forEach(function (x) { c += pick(cur, x, ms, 0); p += pick(pb, x, ms, 0); });
      t += '<td class="num" title="' + esc(d) + ' ' + (m === '計' ? '合計' : m + '月')
        + '　今年 ' + num(c) + ' ／ 前年 ' + num(p) + '">'
        + mark(ratio(c, p), c, p) + '<br><small>' + mark(delta(c, p), c, p)
        + '</small></td>';
    });
    t += '</tr>';
  });
  $('#cmpMonth').innerHTML = t + '</tbody>';
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
    });
    svg.appendChild(svText(L + band * i + band / 2, T + h + 18, m + '月', 'lab'));
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
    groups: [{ name: DATA.year + '年', series: mk(cur) },
             { name: prev.year + '年', hatch: true, series: mk(pb) }],
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

  // 月×営業部 マトリクス
  var t = $('#matrix'), h = '<thead><tr><th>部署</th>';
  MONTHS.forEach(function (m) { h += '<th class="num">' + m + '月</th>'; });
  h += '<th class="num">年間合計</th></tr></thead><tbody>';
  DATA.depts.forEach(function (d) {
    h += '<tr><td><b>' + esc(d) + '</b></td>';
    var sum = 0, cnt = 0;
    MONTHS.forEach(function (m) {
      var rs = recs(m, d), v = rs.reduce(function (s, r) { return s + r.tsu; }, 0);
      sum += v; cnt += rs.length;
      h += '<td class="num"><a href="#" data-m="' + m + '" data-d="' + esc(d) + '" class="jump">'
        + (v ? num(v) : '－') + '</a><br><small style="color:#5b6478">'
        + (rs.length ? rs.length + '件' : '') + '</small></td>';
    });
    h += '<td class="num"><b>' + num(sum) + '</b><br><small style="color:#5b6478">' + cnt + '件</small></td></tr>';
  });
  h += '<tr class="total"><td>合計</td>';
  var gt = 0, gc = 0;
  MONTHS.forEach(function (m) {
    var rs = recs(m, null), v = rs.reduce(function (s, r) { return s + r.tsu; }, 0);
    gt += v; gc += rs.length;
    h += '<td class="num">' + (v ? num(v) : '－') + '<br><small>' + (rs.length ? rs.length + '件' : '') + '</small></td>';
  });
  h += '<td class="num">' + num(gt) + '<br><small>' + gc + '件</small></td></tr></tbody>';
  t.innerHTML = h;
  $$('#matrix a.jump').forEach(function (a) {
    a.onclick = function (e) {
      e.preventDefault();
      S.month = +a.dataset.m; S.dept = a.dataset.d;
      switchView('month'); renderMonthControls(); renderMonth();
    };
  });
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

var saveTimer = null;
function onEdit(f, now) {
  var tr = f.closest('tr'), r = DATA.records[+tr.dataset.i];
  var k = key(r), m = MEMO[k] || { trend: '', plan: '', tsu: '' };
  m[f.dataset.f] = f.value;
  if (m.trend || m.plan || m.tsu) MEMO[k] = m; else delete MEMO[k];
  tr.querySelectorAll('td.inp').forEach(function (td) {
    td.classList.toggle('filled', !!(m.trend || m.plan || m.tsu));
  });
  updateTotals();
  clearTimeout(saveTimer);
  saveTimer = setTimeout(save, now ? 0 : 700);
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

function save() {
  indexMemo();
  try { localStorage.setItem('kadouMemo', JSON.stringify(MEMO)); } catch (e) { /* 容量超過は無視 */ }
  fetch('/api/memo', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(MEMO)
  }).then(function () { flash(); }).catch(function () { flash(); });
}

function flash() {
  var s = $('#saved'); s.classList.add('on');
  setTimeout(function () { s.classList.remove('on'); }, 1200);
}

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

function csvCmpM() {
  var prev = yearSummary(DATA.year - 1), months = DATA.months || [];
  if (!prev || !prev.byDept) { return; }
  var cur = curByDept(), pb = prev.byDept;
  var head = ['部署', '区分'].concat(months.map(function (m) { return m + '月'; }));
  var rows = [head.concat(['合計'])];
  DATA.depts.concat(['合計']).forEach(function (d) {
    var ds = d === '合計' ? DATA.depts : [d];
    var now = [], before = [];
    months.concat(['計']).forEach(function (m) {
      var ms = m === '計' ? months : [m];
      var c = 0, p = 0;
      ds.forEach(function (x) { c += pick(cur, x, ms, 0); p += pick(pb, x, ms, 0); });
      now.push(c); before.push(p);
    });
    rows.push([d, '今年'].concat(now));
    rows.push([d, '前年'].concat(before));
    rows.push([d, '前年比'].concat(now.map(function (c, i) {
      return ratio(c, before[i]);
    })));
  });
  csv(rows, DATA.year + '年_月別前年比_部署別.xlsx');
}

function csvYear() {
  var rows = [['営業部'].concat(MONTHS.map(function (m) { return m + '月'; })).concat(['年間合計'])];
  DATA.depts.forEach(function (d) {
    rows.push([d].concat(MONTHS.map(function (m) {
      return recs(m, d).reduce(function (s, r) { return s + r.tsu; }, 0);
    })).concat([recs(null, d).reduce(function (s, r) { return s + r.tsu; }, 0)]));
  });
  rows.push(['合計'].concat(MONTHS.map(function (m) {
    return recs(m, null).reduce(function (s, r) { return s + r.tsu; }, 0);
  })).concat([DATA.records.reduce(function (s, r) { return s + r.tsu; }, 0)]));
  csv(rows, DATA.year + '年_月別営業部別_通し数.xlsx');
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
                 raw: '日報明細', src: '元ファイル' };
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
$('#btnCsvCmpM').onclick = csvCmpM;
$('#btnCsvMonth').onclick = csvMonth;
$('#btnCsvClient').onclick = csvClient;
$('#btnCsvRaw').onclick = csvRaw;

$('#btnSetting').onclick = function () {
  fetch('/api/config').then(function (r) { return r.json(); }).then(function (c) {
    var list = Array.isArray(c.src) ? c.src.slice() : (c.src ? [c.src] : []);
    (c.srcAlt || []).forEach(function (x) { if (list.indexOf(x) < 0) list.push(x); });
    $('#inSrc').value = list.join('\n');
    $('#inYear').value = c.year || '';
    // 社内サーバーで共有しているときは、読むフォルダを画面から変えられない
    // （サーバー上の好きなフォルダを読めてしまわないようにするため）
    var ro = !!c.shared;
    $('#inSrc').readOnly = ro;
    $('#inYear').readOnly = ro;
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
$('#dlgOk').onclick = function () { $('#dlg').close(); doRebuild($('#inSrc').value, $('#inYear').value); };
$('#btnRebuild').onclick = function () { doRebuild(); };

function doRebuild(src, year) {
  var b = $('#btnRebuild'); b.disabled = true; b.textContent = '読込中…';
  fetch('/api/rebuild', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ src: src, year: year })
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
