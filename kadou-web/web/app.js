/* 稼動日報 印刷実績ビューア
   data.json（build.py が生成）を読み、月別・営業部別に表示する。
   記入欄の内容は /api/memo（memo.json）に保存する。元の .xls は読み取りのみ。 */
'use strict';

var DATA = null, MEMO = {}, DEPTKEY = { '本社営業部': 'hq', '東京営業部': 'tk', '池袋営業部': 'ik' };
var S = { view: 'year', month: null, dept: '全社', deptC: '全社', q: '', qc: '' };
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

/* ── 起動 ── */
function boot() {
  Promise.all([
    fetch('data.json?t=' + Date.now()).then(function (r) { return r.ok ? r.json() : null; }).catch(function () { return null; }),
    fetch('/api/memo').then(function (r) { return r.ok ? r.json() : {}; }).catch(function () { return localMemo(); })
  ]).then(function (a) {
    DATA = a[0];
    MEMO = a[1] || {};
    if (!DATA) {
      $('#warns').innerHTML = '<div class="warn"><b>データがありません。</b>'
        + '右上の「設定」で稼動日報フォルダを指定し、「保存して読み直す」を押してください。</div>';
      return;
    }
    S.month = DATA.months[DATA.months.length - 1];
    render();
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
    + DATA.files.length + 'ファイル　／　' + num(DATA.records.length) + '件';
  var w = DATA.warnings || [];
  $('#warns').innerHTML = w.length
    ? '<div class="warn"><b>注意</b><br>' + w.map(esc).join('<br>') + '</div>' : '';
  renderYear(); renderMonthControls(); renderMonth(); renderClientControls(); renderClient(); renderSrc();
}

/* 年間サマリー */
function renderYear() {
  var all = DATA.records, tot = all.reduce(function (s, r) { return s + r.tsu; }, 0);
  var k = $('#kpis'); k.innerHTML = '';
  k.appendChild(el('div', 'kpi', '<div class="k">年間 通し数</div><div class="v">' + num(tot)
    + '<span class="u">通し</span></div>'));
  k.appendChild(el('div', 'kpi', '<div class="k">年間 件数（管理番号）</div><div class="v">' + num(all.length)
    + '<span class="u">件</span></div>'));
  DATA.depts.forEach(function (d) {
    var rs = all.filter(function (r) { return r.dept === d; });
    var t = rs.reduce(function (s, r) { return s + r.tsu; }, 0);
    k.appendChild(el('div', 'kpi ' + DEPTKEY[d],
      '<div class="k">' + d + '（' + DATA.deptCodes[d].join(' / ') + '）</div>'
      + '<div class="v">' + num(t) + '<span class="u">通し</span></div>'
      + '<div class="k">' + num(rs.length) + '件　構成比 '
      + (tot ? (t / tot * 100).toFixed(1) : '0.0') + '%</div>'));
  });

  // 積み上げ棒グラフ
  var per = MONTHS.map(function (m) {
    var o = { m: m, total: 0 };
    DATA.depts.forEach(function (d) {
      o[d] = recs(m, d).reduce(function (s, r) { return s + r.tsu; }, 0);
      o.total += o[d];
    });
    return o;
  });
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

  // 月×営業部 マトリクス
  var t = $('#matrix'), h = '<thead><tr><th>営業部</th>';
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
function renderMonthControls() {
  var s = $('#selMonth'); s.innerHTML = '';
  MONTHS.forEach(function (m) {
    var n = recs(m, null).length;
    var o = el('option', null, DATA.year + '年 ' + m + '月' + (n ? '（' + n + '件）' : '（データなし）'));
    o.value = m; o.disabled = !n;
    if (m === S.month) o.selected = true;
    s.appendChild(o);
  });
  s.onchange = function () { S.month = +s.value; renderMonth(); };

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
    + (all ? '<th>営業部</th>' : '')
    + '<th>管理番号</th><th>ｸﾗｲｱﾝﾄ名</th><th>品名</th>'
    + '<th class="inp">今年の動向</th><th class="inp">無しの場合の代替対策</th><th class="inp num">対策通し数</th>'
    + '<th class="c">営業担当ｺｰﾄﾞ</th><th class="c">印刷日</th><th class="c">色数</th><th class="num">通し数</th>'
    + '</tr></thead><tbody>';
  rs.forEach(function (r) {
    var m = memoOf(r), fill = (m.trend || m.plan || m.tsu) ? ' filled' : '';
    var tip = '機械: ' + r.machines.join(' / ') + (r.nos.length ? '\n統合した管理番号: ' + r.nos.join(', ') : '');
    h += '<tr data-i="' + r.i + '">'
      + (all ? '<td>' + esc(r.dept.replace('営業部', '')) + '</td>' : '')
      + '<td title="' + esc(tip) + '">' + (r.nos.length ? '<b>' + esc(r.no) + '</b><span class="badge">枝番'
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

  $$('#detail textarea, #detail input').forEach(function (f) {
    autosize(f);
    f.oninput = function () { autosize(f); onEdit(f); };
    f.onchange = function () { onEdit(f, true); };
  });
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
  var p = $('#pillDeptC'); p.innerHTML = '';
  ['全社'].concat(DATA.depts).forEach(function (d) {
    var b = el('button', S.deptC === d ? 'on' : '', d);
    b.onclick = function () { S.deptC = d; renderClientControls(); renderClient(); };
    p.appendChild(b);
  });
  $('#qc').value = S.qc;
  $('#qc').oninput = function () { S.qc = this.value; renderClient(); };
}

/* ｸﾗｲｱﾝﾄ名の表記ゆれ吸収キー（集計スキルと同じ NFKC 正規化）。
   「㈱アルファ商事」と「（株）アルファ商事」を同じ得意先として数える。 */
function ckey(s) {
  return String(s || '').normalize('NFKC').replace(/[\s\u3000]/g, '');
}

function clientRows() {
  var map = {};
  recs(null, S.deptC).forEach(function (r) {
    var k = r.client ? ckey(r.client) : '';
    var e = map[k] || (map[k] = { name: '（得意先名なし）', tsu: 0, cnt: 0, months: {}, names: {} });
    if (r.client) e.names[r.client] = (e.names[r.client] || 0) + 1;
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
    return e;
  });
  var q = S.qc.trim().toLowerCase();
  if (q) rows = rows.filter(function (e) { return e.name.toLowerCase().indexOf(q) >= 0; });
  rows.sort(function (a, b) { return b.tsu - a.tsu; });
  return rows;
}

function renderClient() {
  var rows = clientRows(), tot = rows.reduce(function (s, e) { return s + e.tsu; }, 0);
  $('#cTitle').innerHTML = DATA.year + '年　' + esc(S.deptC)
    + ' <small>得意先 ' + rows.length + '社　通し数 ' + num(tot) + '</small>';
  var t = $('#clients');
  if (!rows.length) { t.innerHTML = '<tbody><tr><td class="empty">該当なし</td></tr></tbody>'; return; }
  var h = '<thead><tr><th class="c">順位</th><th>ｸﾗｲｱﾝﾄ名</th><th class="num">件数</th><th class="num">通し数</th><th class="num">構成比</th>';
  MONTHS.forEach(function (m) { h += '<th class="num">' + m + '月</th>'; });
  h += '</tr></thead><tbody>';
  rows.forEach(function (e, i) {
    h += '<tr><td class="c">' + (i + 1) + '</td><td' + (e.alias ? ' title="同一とみなした表記: ' + esc(e.alias.join(' / ')) + '"' : '') + '>'
      + esc(e.name) + (e.alias ? '<span class="badge">表記' + e.alias.length + '</span>' : '') + '</td>'
      + '<td class="num">' + e.cnt + '</td><td class="num"><b>' + num(e.tsu) + '</b></td>'
      + '<td class="num">' + (tot ? (e.tsu / tot * 100).toFixed(1) : '0.0') + '%</td>';
    MONTHS.forEach(function (m) { h += '<td class="num">' + (e.months[m] ? num(e.months[m]) : '') + '</td>'; });
    h += '</tr>';
  });
  h += '<tr class="total"><td></td><td>合計</td><td class="num">'
    + rows.reduce(function (s, e) { return s + e.cnt; }, 0) + '</td><td class="num">' + num(tot) + '</td><td class="num">100.0%</td>';
  MONTHS.forEach(function (m) {
    var v = rows.reduce(function (s, e) { return s + (e.months[m] || 0); }, 0);
    h += '<td class="num">' + (v ? num(v) : '') + '</td>';
  });
  h += '</tr></tbody>';
  t.innerHTML = h;
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
  $('#srcPath').textContent = '読込元: ' + DATA.source
    + '　／　読み込みは一時フォルダへコピーして行うため、元のExcelファイルは変更されません。';

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
  var rows = [['順位', 'ｸﾗｲｱﾝﾄ名', '件数', '通し数'].concat(MONTHS.map(function (m) { return m + '月'; }))];
  clientRows().forEach(function (e, i) {
    rows.push([i + 1, e.name, e.cnt, e.tsu].concat(MONTHS.map(function (m) { return e.months[m] || 0; })));
  });
  csv(rows, DATA.year + '年_' + S.deptC + '_得意先別.csv');
}

/* 画面切替・その他 */
function switchView(v) {
  S.view = v;
  $$('nav.tabs button').forEach(function (b) { b.classList.toggle('on', b.dataset.view === v); });
  $$('section.view').forEach(function (s) { s.classList.toggle('on', s.id === 'v-' + v); });
}

$$('nav.tabs button').forEach(function (b) { b.onclick = function () { switchView(b.dataset.view); }; });
$('#btnPrint').onclick = function () { window.print(); };
$('#btnCsvYear').onclick = csvYear;
$('#btnCsvMonth').onclick = csvMonth;
$('#btnCsvClient').onclick = csvClient;

$('#btnSetting').onclick = function () {
  fetch('/api/config').then(function (r) { return r.json(); }).then(function (c) {
    $('#inSrc').value = c.src; $('#inYear').value = c.year; $('#dlg').showModal();
  }).catch(function () {
    $('#inSrc').value = DATA ? DATA.source : ''; $('#inYear').value = DATA ? DATA.year : 2026;
    $('#dlg').showModal();
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
    boot();
  }).catch(function (e) {
    b.disabled = false; b.textContent = 'データ更新';
    alert('サーバーに接続できませんでした。起動.bat から開き直してください。\n' + e);
  });
}

boot();
