# -*- coding: utf-8 -*-
"""
稼動日報(.xls) 読み取りコア

kadou-report スキルで確定済みの抽出・統合ルールをそのまま実装したもの。
Web アプリ用に「1か月だけ」ではなく「年内の全月」を一度に取れるようにしてある。

【最重要】このモジュールは元の .xls を一切書き換えない。
  ・読み込みは必ず一時フォルダへコピーしてから行う（read_workbook）
  ・元フォルダに対する書き込み処理はこのファイルのどこにも存在しない
"""
import datetime
import re
import shutil
import tempfile
import unicodedata
from collections import OrderedDict
from pathlib import Path

import xlrd

# ────────── 確定ルール（kadou-report スキルと同一） ──────────
DEPTS = OrderedDict([('本社営業部', [1110, 1120]),
                     ('東京営業部', [2100, 2140]),
                     ('池袋営業部', [3810, 3820])])
CODE2DEPT = {c: d for d, cs in DEPTS.items() for c in cs}

NEED = ('日付', '管理番号', '営業担当ｺｰﾄﾞ', '品名', '通し枚数', '表版数', '裏版数')
OPTIONAL = ('ｸﾗｲｱﾝﾄ名',)
TIME_LABELS = ('有効時間', '準備合計', '色合わせ', '印刷時間', 'その他')
EPOCH = datetime.date(1899, 12, 30)

SHEET_RE = re.compile(r'^(?:20)?(\d{2})\s*[年月]\s*(\d{1,2})\s*月$')


def norm(s):
    return str(s).replace('　', '').strip()


def ckey(s):
    """ｸﾗｲｱﾝﾄ名の表記ゆれ吸収キー（NFKC・空白除去）"""
    if not s:
        return ''
    return unicodedata.normalize('NFKC', str(s)).replace(' ', '').replace('　', '')


class Desc:
    """昇順ソートの中で文字列だけ降順にするラッパ"""
    __slots__ = ('v',)

    def __init__(self, v):
        self.v = v

    def __lt__(self, o):
        return self.v > o.v

    def __eq__(self, o):
        return self.v == o.v


def basno(no):
    """管理番号から枝番を落とした基番号  8632175-1-2 → 8632175"""
    s = str(no).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s.split('-')[0]


# ────────── 読み込み（元ファイル保護） ──────────
def read_workbook(path):
    """元ファイルを一時フォルダへコピーしてから開く。

    xlrd 自体は読み取り専用だが、「元のエクセルファイルは壊さない」という
    運用ルールを機械的に保証するため、必ずコピー経由で読む。
    元ファイルはオープンすらされない（shutil.copy2 の読み取りのみ）。
    """
    with tempfile.TemporaryDirectory(prefix='kadou_ro_') as tmp:
        dst = Path(tmp) / Path(path).name
        shutil.copy2(path, dst)          # 元 → 一時。逆方向は行わない
        return xlrd.open_workbook(str(dst))


def cell(sh, r, c):
    if c is None or c >= sh.ncols:
        return None
    v = sh.cell_value(r, c)
    return None if v == '' else v


def header_map(sh):
    """見出し行を探して {列名: 最初に現れた列番号} を返す。

    列位置は月によって違うため必ず見出し文字で特定する。
    「通し枚数」「表版数」「裏版数」は本刷ブロックと集計ブロックの2か所に
    同名で出るので、最初に現れた方（本刷ブロック）を採用する。
    """
    for hr in range(0, min(8, sh.nrows)):
        names = [norm(sh.cell_value(hr, c)) for c in range(sh.ncols)]
        if '日付' in names and '管理番号' in names:
            m = {}
            for c, n in enumerate(names):
                if n and n not in m:
                    m[n] = c
            missing = [k for k in NEED if k not in m]
            if missing:
                raise ValueError('見出し列が見つかりません: ' + ', '.join(missing))
            for k in OPTIONAL:
                m.setdefault(k, None)
            return hr, m
    raise ValueError('見出し行（「日付」「管理番号」を含む行）が見つかりません')


def month_sheets(wb, year2):
    """その年の月次シートを [(月, シート名), ...] で返す。

    「25年9月30」のような日付付き補助シートは SHEET_RE に一致しないので自動的に除外される。
    """
    out = []
    for name in wb.sheet_names():
        m = SHEET_RE.match(norm(name))
        if m and int(m.group(1)) == year2:
            out.append((int(m.group(2)), name))
    return sorted(out)


def extract_sheet(sh, machine, src):
    """1シートから明細行を抽出する"""
    hr, M = header_map(sh)
    out = []
    for r in range(hr + 1, sh.nrows):
        # 日次集計ブロック（有効時間・準備合計・色合わせ・印刷時間・その他・受注件数）は明細ではない
        if norm(sh.cell_value(r, 27) if sh.ncols > 27 else '') in TIME_LABELS + ('受注件数',):
            continue
        d = cell(sh, r, M['日付'])
        if not (isinstance(d, float) and d > 40000):   # 日付シリアルの行だけが明細
            continue
        code = cell(sh, r, M['営業担当ｺｰﾄﾞ'])
        code = int(code) if isinstance(code, float) else code
        if code not in CODE2DEPT:
            continue
        out.append({
            'dept':    CODE2DEPT[code],
            'code':    code,
            'no':      cell(sh, r, M['管理番号']),
            'base':    basno(cell(sh, r, M['管理番号'])),
            'date':    EPOCH + datetime.timedelta(days=int(d)),
            'client':  cell(sh, r, M['ｸﾗｲｱﾝﾄ名']),
            'name':    cell(sh, r, M['品名']),
            'f':       cell(sh, r, M['表版数']) or 0,
            'b':       cell(sh, r, M['裏版数']) or 0,
            'tsu':     cell(sh, r, M['通し枚数']) or 0,
            'machine': machine,
            'src':     src,
            'seq':     r,
        })
    return out


def extract_year(path, year2, machine=None):
    """1ファイルから、その年の全月分の明細を {月: [行, ...]} で返す"""
    wb = read_workbook(path)
    machine = machine or Path(path).stem
    src = Path(path).name
    got = OrderedDict()
    for month, sheet in month_sheets(wb, year2):
        rows = extract_sheet(wb.sheet_by_name(sheet), machine, src)
        got.setdefault(month, []).extend(rows)
    return got


# ────────── 統合・並び替え（確定ルール） ──────────
def consolidate(rows):
    """管理番号ごとに統合する。枝番（-1 / -1-2）は同一管理番号として扱う。

    通し数 = グループ内の合計（表刷り＋裏刷り）
    色数   = 表版数・裏版数それぞれの最大値（同じ版で複数回通した分を重複させない）
    印刷日 = グループ内の初日
    """
    g = OrderedDict()
    for r in rows:
        k = (r['dept'], r['base'])
        e = g.setdefault(k, {'dept': r['dept'], 'base': r['base'], 'code': r['code'],
                             'dates': [], 'f': 0, 'b': 0, 'tsu': 0, 'n': 0,
                             'nos': set(), 'machines': set(), 'members': []})
        e['members'].append(r)
        e['dates'].append(r['date'])
        e['nos'].add(str(r['no']))
        e['machines'].add(r['machine'])
        e['f'] = max(e['f'], r['f'])
        e['b'] = max(e['b'], r['b'])
        e['tsu'] += r['tsu']
        e['n'] += 1
    for e in g.values():
        # ｸﾗｲｱﾝﾄ名: 非空値の最頻値 → 同数なら最も長い → なお同じなら文字列順
        names = [str(m['client']) for m in e['members'] if m['client']]
        e['client'] = (max(set(names), key=lambda s: (names.count(s), len(s), s))
                       if names else None)
        # 品名: 初日 → 管理番号 → 品名 の順で最小の行のもの
        first = min(e['members'], key=lambda m: (m['date'], str(m['no']), str(m['name'] or '')))
        e['name'] = first['name']
        del e['members']
    return list(g.values())


def sort_groups(gs):
    """①営業担当ｺｰﾄﾞ降順 ②ｸﾗｲｱﾝﾄ名降順（空欄は末尾） ③印刷日降順"""
    gs.sort(key=lambda g: (-g['code'],
                           g['client'] in (None, ''),
                           Desc(ckey(g['client'])),
                           -min(g['dates']).toordinal(),
                           g['base']))
    return gs
