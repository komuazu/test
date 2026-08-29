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
var $$ = function (s) { return Array.prototype.slice.call(document.querySelectorAll(s)); };
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

function localMemo() { try { return JSON.parse(localStorage.getItem('kadouMemo') || '{}'); } catch (e) { return {}; } }

/* ── 抽出ヘルパ ── */
function recs(month, dept) {
  return DATA.records.filter(function (r) {
    return (month == null || r.m === month) && (!dept || dept === '全社' || r.dept === dept);
  });
}
function key(r) { return DATA.year + '|' + r.m + '|' + r.dept + '|' + r.no; }
function memoOf(r) { return MEMO[key(r)] || { trend: '', plan: '', tsu: '' }; }
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

/* ── 月別グラフ。積み上げ（部署別）と、前年との比較を切り替えられる ── */
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
      + (S.chartDept === 'すべて' ? '部署ごと' : S.chartDept)
    : '部署の積み上げ';
  if (yoy) { renderChartYoY(prev); } else { renderChartStack(); }
}

function renderChartStack() {
  $('#chart').className = 'bars';
  var per = MONTHS.map(function (m) {
    var o = { m: m, total: 0 };
    DATA.depts.forEach(function (d) {
      o[d] = recs(m, d).reduce(function (s, r) { return s + r.tsu; }, 0);
      o.total += o[d];
    });
    return o;
  });
  $('#legend').innerHTML = DATA.depts.map(function (d) {
    return '<span><i class="' + DEPTKEY[d] + '"></i>' + esc(d) + '</span>';
  }).join('');
  var max = Math.max.apply(null, per.map(function (o) { return o.total; }).concat([1]));
  var c = $('#chart'); c.innerHTML = '';
  per.forEach(function (o) {
    var b = el('div', 'bar');
    var track = el('div', 'track'), st = el('div', 'stack');
    track.appendChild(el('div', 'val', o.total ? num(o.total) : ''));   // 棒のすぐ上に数値
    st.style.height = (o.total / max * 100) + '%';
    DATA.depts.forEach(function (d) {
      if (!o[d]) return;
      var g = el('div', 'seg ' + DEPTKEY[d]);
      g.style.height = (o[d] / o.total * 100) + '%';
      g.title = o.m + '月 ' + d + '　' + num(o[d]) + ' 通し';
      st.appendChild(g);
    });
    track.appendChild(st);
    b.appendChild(track);
    b.appendChild(el('div', 'lab', o.m + '月'));
    c.appendChild(b);
  });
}

/* 今年と前年を月ごとに並べる。前年は years.json の小さな集計から取る。
   「すべて」を選ぶと、全社に続けて部署ごとの小さなグラフを縦に並べる。 */
function renderChartYoY(prev) {
  var cur = curByDept(), pb = prev.byDept;
  var months = (DATA.months && DATA.months.length) ? DATA.months : MONTHS;
  $('#legend').innerHTML = '<span><i class="now"></i>' + DATA.year + '年</span>'
    + '<span><i class="was"></i>' + prev.year + '年</span>';

  var c = $('#chart');
  c.innerHTML = '';
  if (S.chartDept !== 'すべて') {
    c.className = 'bars';
    yoyBars(c, S.chartDept === '全社' ? DATA.depts : [S.chartDept],
            cur, pb, months, prev, false);
    return;
  }
  monthDeptBars(c, cur, pb, months, prev);
}

/* 月ごとに「今年の積み上げ」と「前年の積み上げ」を並べる。
   積み上げのグラフと同じ形（部署を積み上げた1本）を2本並べただけなので、
   見方を変えずに前年と見比べられる。前年は同じ色の斜線で区別する。 */
function monthDeptBars(c, cur, pb, months, prev) {
  c.className = 'bars deptyoy';
  var tot = months.map(function (m) {
    var now = 0, was = 0;
    DATA.depts.forEach(function (d) {
      now += pick(cur, d, [m], 0); was += pick(pb, d, [m], 0);
    });
    return { m: m, now: now, was: was };
  });
  var max = Math.max.apply(null, tot.map(function (o) {
    return Math.max(o.now, o.was);
  }).concat([1]));
  $('#legend').innerHTML = DATA.depts.map(function (d) {
    return '<span><i class="' + DEPTKEY[d] + '"></i>' + esc(d) + '</span>';
  }).join('') + '<span class="lg">塗り＝' + DATA.year + '年</span>'
    + '<span class="lg"><i class="hatch"></i>斜線＝' + prev.year + '年</span>';

  tot.forEach(function (o) {
    var b = el('div', 'bar'), track = el('div', 'track'), pair = el('div', 'pair');
    track.appendChild(el('div', 'val', o.now ? num(o.now) : ''));
    [[DATA.year, o.now, cur, ''], [prev.year, o.was, pb, ' was']].forEach(function (x) {
      var col = el('div', 'col');
      var st = el('div', 'stack' + x[3]);
      st.style.height = (x[1] / max * 100) + '%';
      DATA.depts.forEach(function (d) {
        var v = pick(x[2], d, [o.m], 0);
        if (!v) { return; }
        var g = el('div', 'seg ' + DEPTKEY[d]);
        g.style.height = (v / x[1] * 100) + '%';
        g.title = o.m + '月 ' + x[0] + '年 ' + d + '　' + num(v) + ' 通し';
        st.appendChild(g);
      });
      col.appendChild(st);
      col.title = o.m + '月 ' + x[0] + '年　' + num(x[1]) + ' 通し';
      pair.appendChild(col);
    });
    track.appendChild(pair);
    b.appendChild(track);
    b.appendChild(el('div', 'rate', o.was
      ? mark(ratio(o.now, o.was), o.now, o.was) : '－'));
    b.appendChild(el('div', 'lab', o.m + '月'));
    c.appendChild(b);
  });
}

/* 1つぶんの棒グラフ（今年と前年の2本組を月ごとに） */
function yoyBars(box, ds, cur, pb, months, prev, mini) {
  var per = months.map(function (m) {
    var now = 0, was = 0;
    ds.forEach(function (d) { now += pick(cur, d, [m], 0); was += pick(pb, d, [m], 0); });
    return { m: m, now: now, was: was };
  });
  var max = Math.max.apply(null, per.map(function (o) {
    return Math.max(o.now, o.was);
  }).concat([1]));
  per.forEach(function (o) {
    var b = el('div', 'bar');
    var track = el('div', 'track'), pair = el('div', 'pair');
    track.appendChild(el('div', 'val', o.now ? num(o.now) : ''));
    [['now', o.now, DATA.year], ['was', o.was, prev.year]].forEach(function (x) {
      var col = el('div', 'col ' + x[0]);
      var bar = el('b');
      bar.style.height = (x[1] / max * 100) + '%';
      bar.className = x[0] === 'was' ? 'was' : '';
      col.title = o.m + '月 ' + x[2] + '年　' + num(x[1]) + ' 通し';
      col.appendChild(bar);
      pair.appendChild(col);
    });
    track.appendChild(pair);
    b.appendChild(track);
    b.appendChild(el('div', 'rate', o.was
      ? mark(ratio(o.now, o.was), o.now, o.was) : '－'));
    b.appendChild(el('div', 'lab', o.m + '月'));
    box.appendChild(b);
  });
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
    + '<td class="num" id="sumDiff">' + num(tot - pt) + '</td><td></td><td></td>'
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
  if ($('#sumDiff')) $('#sumDiff').textContent = num(tot - pt);
  if ($('#sumRate')) $('#sumRate').textContent = tot ? (pt / tot * 100).toFixed(1) + '%' : '－';
}

function save() {
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
  csv(out, DATA.year + '年' + S.monthR + '月_日報明細_' + S.machineR + '.csv');
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
function csv(rows, filename) {
  var body = rows.map(function (r) {
    return r.map(function (c) {
      c = c == null ? '' : String(c);
      return /[",\r\n]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c;
    }).join(',');
  }).join('\r\n');
  var blob = new Blob(['﻿' + body], { type: 'text/csv;charset=utf-8' });
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
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
  csv(rows, DATA.year + '年_前年比較_部署別.csv');
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
  csv(rows, DATA.year + '年_月別前年比_部署別.csv');
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
  csv(rows, DATA.year + '年_月別営業部別_通し数.csv');
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
  csv(rows, DATA.year + '年' + S.month + '月_' + S.dept + '_印刷実績.csv');
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
  csv(rows, DATA.year + '年' + (one ? S.monthC + '月' : '') + '_' + S.deptC + '_得意先別.csv');
}

/* 画面切替・その他 */
function switchView(v) {
  S.view = v;
  $$('nav.tabs button').forEach(function (b) { b.classList.toggle('on', b.dataset.view === v); });
  $$('section.view').forEach(function (s) { s.classList.toggle('on', s.id === 'v-' + v); });
}

$$('nav.tabs button').forEach(function (b) { b.onclick = function () { switchView(b.dataset.view); }; });
$('#btnPrint').onclick = function () { window.print(); };
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
    boot();
  }).catch(function (e) {
    b.disabled = false; b.textContent = 'データ更新';
    alert('サーバーに接続できませんでした。起動.bat から開き直してください。\n' + e);
  });
}

boot();
