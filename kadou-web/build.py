# -*- coding: utf-8 -*-
"""
稼動日報フォルダ → Web アプリ用データ生成

  python build.py --src "C:\\Users\\h-tag\\Desktop\\13年稼動～25年稼動"
  python build.py --src "U:\\...\\★新・26年稼動" --year 2026

フォルダ配下の .xls をすべて読み、月次シート（例「26年1月」）を全部拾って、
営業部別・月別に統合したデータを年ごとに書き出す。

  web/years.json      … 見つかった年の一覧（画面の年切替に使う）
  web/data_2025.json  … 2025年ぶんのデータ（年ごとに1ファイル）

--year を省略すると、フォルダに入っている年を全部作る。
13年〜25年のように複数年が1フォルダにある場合も、ファイルは1回開くだけで済む。

【元ファイル保護】
  ・読み込みは kadou_core.read_workbook（一時フォルダへコピーしてから開く）のみ
  ・--src で指定したフォルダには一切書き込まない（出力先は必ずこのアプリ側）
  ・Excel 編集中の一時ファイル ~$*.xls は読み飛ばす
"""
import argparse
import datetime
import json
import re
import sys
import traceback
import unicodedata
from collections import OrderedDict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from kadou_core import (ALL_DEPTS, DEPTS, OTHER, consolidate,     # noqa: E402
                        extract_all, sort_groups)

HERE = Path(__file__).resolve().parent
WEB = HERE / 'web'


def fold_name(s):
    """フォルダ名の表記ゆれを吸収するキー

    「13年稼動～25年稼動」の「～」は、全角チルダ(U+FF5E)と波ダッシュ(U+301C)の
    2種類が混在しやすく、見た目が同じでも別の文字として扱われる。
    全角半角・空白の違いも含めて同一視する。
    """
    t = unicodedata.normalize('NFKC', str(s))
    for ch in '〜～~⁓∼':
        t = t.replace(ch, '~')
    return t.replace(' ', '').replace('　', '').lower()


_QUOTES = '"' + "'" + '“”「」'      # 貼り付けで付いてくる引用符


def clean_src(s):
    r"""指定されたフォルダのパスを整える（前後の空白と引用符を落とす）

    エクスプローラーの「パスのコピー」は "C:\..." のように引用符付きで入る。
    そのまま設定に入れると実在するフォルダでも見つからなくなるため、剥がしておく。
    """
    return str(s).strip().strip(_QUOTES)


def resolve_src(src):
    """稼動日報フォルダの一覧から、実在するものを全部返す。

    src は文字列でも複数フォルダのリストでもよい。次の2つを兼ねる。
      ・データが複数の場所にある（例: ローカルの13〜25年 ＋ U: の26年）→ まとめて読む
      ・同じフォルダへの別の行き方（U: と UNC）→ 片方しか存在しないので自然に片方だけ使う
    どちらの場合も、同じファイルを二重に数えないよう後段で重複を除く。

    そのままの名前で見つからないときは、親フォルダの中を表記ゆれを吸収して探し直す。

    戻り値: ([実在したフォルダ, ...], [指定された全候補, ...])
    """
    cands = [src] if isinstance(src, str) else list(src)
    cands = [x for x in (clean_src(c) for c in cands if c) if x]
    found, seen = [], set()
    for c in cands:
        p = Path(c)
        if not p.is_dir():
            p = match_by_name(p)
            if p is None:
                continue
        key = str(p)
        if key not in seen:
            seen.add(key)
            found.append(p)
    return found, cands


def match_by_name(p):
    """親フォルダの中から、表記ゆれを無視して同じ名前のフォルダを探す"""
    parent = p.parent
    if not parent.is_dir():
        return None
    want = fold_name(p.name)
    try:
        for d in sorted(parent.iterdir()):
            if d.is_dir() and fold_name(d.name) == want:
                return d
    except OSError:
        pass
    return None


def hint_subdirs(cands):
    """フォルダが見つからないとき、親フォルダにある候補を案内する

    候補のパスをさかのぼって、実在する一番深いフォルダの中身を見せる。
    どこまでは合っていて、どこから違うのかが分かるようにするため。
    """
    here = Path.cwd().resolve()
    for c in cands:
        for parent in Path(c).parents:
            if str(parent) in ('.', '', '/') or not parent.is_dir():
                continue
            if parent.resolve() == here:      # アプリ自身のフォルダは案内にならない
                break
            subs = sorted(d.name for d in parent.iterdir() if d.is_dir())
            if subs:
                return 'ここまでは実在します: %s\n  その中にあるフォルダ: %s' % (
                    parent, '、'.join(subs[:12]))
            break
    return ''


NEW, OLD = '新', '旧'
MIN_YEAR = 2013          # これより前の年は読まない（古い試験用の日報が混ざるため）
MAX_DEPTH = 6            # 年フォルダを探しにいく深さ

# 「★新・26年稼動」「26年稼働」「13年稼働」といった年フォルダの名前。
# 稼動と稼働のどちらでもよい。
YEAR_DIR = re.compile(r'^(★新[・･])?(\d{2})年稼[動働]$')


def inner_dir(d):
    """★新・16年稼動\★新・16年稼動 のような同名の入れ子は、中側を使う"""
    nested = d / d.name
    return nested if nested.is_dir() else d


def year_dir_kind(name):
    """フォルダ名が年フォルダの形なら (年, 種類) を返す。違えば None

      ★新・NN年稼動 … その年の正（新しい方）
      NN年稼働      … ★新 に無い年・月を補う古い方
    """
    m = YEAR_DIR.match(fold_name(name))
    if not m:
        return None
    return 2000 + int(m.group(2)), NEW if m.group(1) else OLD


def scan_year_dirs(root, maxdepth=MAX_DEPTH):
    """配下をたどって年フォルダを全部見つける

    入れ物のフォルダ（「13年稼動～25年稼動」など）の名前は見ない。年が変わって
    「13年～26年稼働」に変わっても、「27年稼働」が増えても、設定を直さずに
    最新のデータが読めるようにするため。

    戻り値: [(年, 種類, 深さ, フォルダ), ...]
    """
    out = []

    def walk(d, depth):
        if depth > maxdepth:
            return
        try:
            kids = sorted(x for x in d.iterdir() if x.is_dir())
        except OSError:
            return
        for k in kids:
            hit = year_dir_kind(k.name)
            if hit:
                out.append((hit[0], hit[1], depth, inner_dir(k)))
                continue                     # 年フォルダの中はもう探さない
            walk(k, depth + 1)

    walk(Path(root), 1)
    return out


def year_folders(src):
    """実際に読む年フォルダを決める

    同じ年・同じ種類が何か所にもあるときは、いちばん浅いものだけを使う。
    控えのフォルダ（☆平版印刷課印刷稼働日報 の下など）に置かれた写しは階層が
    深いので、名前で除外しなくても自然に外れる。同じ通し数を二重に数えないため。

    src 自身が年フォルダのこともある（U: の ★新・26年稼動 を直接指定した場合）。
    年フォルダが1つも無ければ空を返し、呼び出し側が従来どおり配下を全部読む。

    戻り値: [(フォルダ, NEW か OLD), ...]
    """
    src = Path(src)
    hit = year_dir_kind(src.name)
    if hit:
        return [(inner_dir(src), hit[1])]
    if not src.is_dir():
        return []
    best = {}
    for year, kind, depth, path in scan_year_dirs(src):
        if year < MIN_YEAR:
            continue
        k = (year, kind)
        if k not in best or depth < best[k][0]:
            best[k] = (depth, path)
    return [(p, kind) for (_y, kind), (_d, p) in sorted(best.items())]


def report_files(folder, kind):
    """年フォルダ直下の稼動日報だけを返す

    サブフォルダ（日報取込用 など）は同じ日報の写しなので見ない。
    集約表・稼働率・断裁機・軽ｵﾌ印刷機などは名前で外れる。
    """
    out = []
    for p in sorted(folder.glob('*.xls')):
        if p.name.startswith('~$'):
            continue
        if kind == NEW:
            if p.name.startswith('新') and '稼動日報' in p.name:
                out.append(p)                      # 新25・菊全UV稼動日報【新台】.xls も拾う
        elif p.name.endswith('稼動日報.xls'):
            out.append(p)                          # 三菱稼動日報.xls / 小森1号機稼動日報.xls
    return out


def file_key(p):
    """同じファイル（名前・サイズ・更新日時が同じ）かどうかの目印"""
    try:
        st = p.stat()
        return (p.name, st.st_size, int(st.st_mtime))
    except OSError:
        return (p.name, None, None)


def collect_files(srcs):
    """読む対象の稼動日報を、年フォルダごとに1セットだけ集める

    「☆平版印刷課印刷稼働日報」のような控えのフォルダや、年フォルダの中の
    「日報取込用」は同じ日報の写しで、そのまま読むと同じ通し数を何度も数えて
    しまう（2016年は4回数えていた）。年フォルダの直下だけを見ることで防ぐ。

    戻り値: [(パス, NEW か OLD), ...]
    """
    if isinstance(srcs, (str, Path)):
        srcs = [srcs]

    best, plain = {}, []
    for order, src in enumerate(srcs):
        src = Path(src)
        hit = year_dir_kind(src.name)
        if hit:                                    # src 自身が年フォルダ
            found = [(hit[0], hit[1], 0, inner_dir(src))]
        else:
            found = scan_year_dirs(src) if src.is_dir() else []
        if not found:
            plain.append(src)                      # 年フォルダが無い形（テスト用など）
            continue
        for year, kind, depth, path in found:
            if year < MIN_YEAR:
                continue
            k = (year, kind)
            # 浅いものを優先し、同じ深さなら設定で先に書いたフォルダを採る。
            # デスクトップの控えと U: の本体のように、同じ年が別のフォルダにも
            # あるとき、両方読むと同じ通し数を二重に数えてしまうため。
            if k not in best or (depth, order) < best[k][:2]:
                best[k] = (depth, order, path, kind)

    got, seen = [], set()
    for _key, (_d, _o, folder, kind) in sorted(best.items()):
        for p in report_files(folder, kind):
            fk = file_key(p)
            if fk not in seen:
                seen.add(fk)
                got.append((p, kind))
    for src in plain:
        for p in find_xls([src]):
            fk = file_key(p)
            if fk not in seen:
                seen.add(fk)
                got.append((p, NEW))
    return got


def find_xls(srcs):
    """対象フォルダ配下の .xls を集める

    ・Excel 編集中の一時ファイル ~$*.xls は除外する
    ・複数フォルダを指定した場合、同じファイル（名前・サイズ・更新日時が同じ）は
      1つだけ採る。U: と UNC の両方が通るPCで二重に数えないため。
    """
    if isinstance(srcs, (str, Path)):
        srcs = [srcs]
    files, seen = [], set()
    for src in srcs:
        for p in sorted(Path(src).rglob('*.xls*')):
            if p.name.startswith('~$') or p.suffix.lower() not in ('.xls', '.xlsx', '.xlsm'):
                continue
            try:
                st = p.stat()
                key = (p.name, st.st_size, int(st.st_mtime))
            except OSError:
                key = (p.name, None, None)
            if key in seen:
                continue
            seen.add(key)
            files.append(p)
    return files


def read_folder(files):
    """稼動日報を1回ずつ開いて、年→月→明細行 にまとめる

    files は collect_files が返す [(パス, NEW か OLD), ...]。
    同じ年・月が「★新・NN年稼動」と「NN年稼働」の両方にあるときは ★新 だけを採る
    （★新 が改訂版。両方入れると同じ通し数を二重に数えてしまう）。

    戻り値: (rows, daily, cols, file_info, warnings)
      rows  = {年4桁: {月: [明細行, ...]}}
      daily = {年4桁: {月: [日次集計行, ...]}}
    """
    part = {NEW: (OrderedDict(), OrderedDict()), OLD: (OrderedDict(), OrderedDict())}
    cols, file_info, warnings = [], [], []

    for path, kind in files:
        if path.suffix.lower() != '.xls':
            warnings.append('%s は旧形式(.xls)ではないため読み飛ばしました。' % path.name)
            continue
        skipped = []
        try:
            got, days, cc = extract_all(path, skipped=skipped)
        except Exception as e:                                   # noqa: BLE001
            warnings.append('%s を読めませんでした: %s' % (path.name, e))
            traceback.print_exc(file=sys.stderr)
            continue
        for sheet, why in skipped:
            warnings.append('%s の「%s」は形式が違うため飛ばしました: %s'
                            % (path.name, sheet, why))

        for n in cc:
            if n not in cols:
                cols.append(n)

        krows, kdaily = part[kind]
        # このファイル自身の年別の明細行数を数える（他のファイルの分と混ぜない）
        per_year, total = OrderedDict(), 0
        for year2, months in got.items():
            year = 2000 + year2
            e = per_year.setdefault(year, {'months': [], 'rows': 0})
            for month, rr in months.items():
                krows.setdefault(year, OrderedDict()).setdefault(month, []).extend(rr)
                e['months'].append(month)
                e['rows'] += len(rr)
                total += len(rr)
        for year2, months in days.items():
            for month, dd in months.items():
                kdaily.setdefault(2000 + year2, OrderedDict()).setdefault(month, []).extend(dd)

        if not per_year:
            warnings.append('%s に月次シート（「26年1月」形式）がありません。' % path.name)
        file_info.append({'name': path.name, 'machine': path.stem, 'path': str(path),
                          'kind': kind,
                          'years': {str(y): {'months': sorted(e['months']), 'rows': e['rows']}
                                    for y, e in per_year.items()},
                          'rows': total})
        print('  読込 %-40s %-24s 明細%6d行'
              % (path.name,
                 '、'.join('%d年' % y for y in sorted(per_year)) or '月次シートなし',
                 total))

    rows, daily, notes = prefer_new(part)
    return rows, daily, cols, file_info, warnings + notes


def prefer_new(part):
    """★新 にある年・月は ★新 だけを採り、無い月だけ古いフォルダで補う

    2015年は8月から ★新・15年稼動 が始まっており、1〜7月は「15年稼働」にしかない。
    重なる8〜12月を両方入れると二重になるため、月ごとに ★新 を優先する。

    戻り値: (rows, daily, [お知らせ, ...])
    """
    new_rows, new_daily = part[NEW]
    old_rows, old_daily = part[OLD]
    covered = {(y, m) for y, ms in new_rows.items() for m in ms}

    rows, daily = OrderedDict(), OrderedDict()
    for src_rows, src_daily, is_new in ((new_rows, new_daily, True),
                                        (old_rows, old_daily, False)):
        for year, months in src_rows.items():
            for month, rr in months.items():
                if not is_new and (year, month) in covered:
                    continue                       # ★新 が正。古い方は数えない
                rows.setdefault(year, OrderedDict()).setdefault(month, []).extend(rr)
        for year, months in src_daily.items():
            for month, dd in months.items():
                if not is_new and (year, month) in covered:
                    continue
                daily.setdefault(year, OrderedDict()).setdefault(month, []).extend(dd)

    notes = []
    for year in sorted({y for y, _m in covered}):
        dup = sorted(m for m in old_rows.get(year, ()) if (year, m) in covered)
        if dup:
            notes.append('%d年の%sは「★新・%02d年稼動」の内容を使いました'
                         '（古いフォルダの同じ月は数えていません）。'
                         % (year, '、'.join('%d月' % m for m in dup), year % 100))
    return rows, daily, notes


def build_year(year, by_month, daily_by_month, cols, files, src, warnings):
    """1年ぶんのデータを組み立てる（統合ルールは kadou-report と同じ）"""
    records, stats = [], []
    # 「その他」に入った営業担当ｺｰﾄﾞ（画面の見出しに出す）
    other_codes = sorted({r['code'] for rr in by_month.values() for r in rr
                          if r['dept'] == OTHER})
    for month in sorted(by_month):
        rows = by_month[month]
        rows.sort(key=lambda x: (x['date'], str(x['no']), x['seq']))
        for dept in ALL_DEPTS:
            drows = [r for r in rows if r['dept'] == dept]
            groups = sort_groups(consolidate(drows))
            for i, g in enumerate(groups):
                records.append({
                    'i':       len(records),
                    'm':       month,
                    'dept':    dept,
                    'ord':     i,
                    'code':    g['code'],
                    'no':      g['base'],
                    'client':  g['client'] or '',
                    'name':    g['name'] or '',
                    'date':    min(g['dates']).isoformat(),
                    'color':   '%d/%d' % (g['f'], g['b']),
                    'tsu':     int(g['tsu']),
                    'rows':    g['n'],
                    'nos':     sorted(g['nos']) if len(g['nos']) > 1 else [],
                    'machines': sorted(g['machines']),
                    'multi':   len(set(g['dates'])) > 1,
                    # 日報の元の明細行（この管理番号の内訳）。空セルは落として軽くする
                    'det':     [dict([('__機械', mm['machine'])]
                                     + [(k, v) for k, v in mm['raw'].items() if v not in (None, '')])
                                for mm in g['members']],
                })
            # 検証用: 統合の前後で通し数が変わっていないこと
            stats.append({'m': month, 'dept': dept,
                          'detail': len(drows), 'groups': len(groups),
                          'tsuDetail': int(sum(r['tsu'] for r in drows)),
                          'tsuGroups': int(sum(g['tsu'] for g in groups))})

    warns = list(warnings)
    bad = [s for s in stats if s['tsuDetail'] != s['tsuGroups']]
    if bad:
        warns.append('統合前後で通し数が一致しない月があります: '
                     + ', '.join('%d月/%s' % (s['m'], s['dept']) for s in bad))

    # この年のシートを持つファイルだけを「元ファイル」として載せる
    yfiles = []
    for f in files:
        e = f['years'].get(str(year))
        if e:
            yfiles.append({'name': f['name'], 'machine': f['machine'],
                           'months': e['months'], 'rows': e['rows']})
    return {
        'year':      year,
        'generated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
        'source':    str(src),
        'depts':     list(ALL_DEPTS),
        'deptCodes': dict({d: cs for d, cs in DEPTS.items()}, **{OTHER: other_codes}),
        'months':    sorted(by_month),
        'files':     yfiles,
        'stats':     stats,
        'warnings':  warns,
        'cols':      cols,
        'daily':     [dict(d, m=m) for m in sorted(daily_by_month) for d in daily_by_month[m]],
        'records':   records,
    }


def build(src, year=None, outdir=None):
    """フォルダを読み、年ごとの data_YYYY.json と years.json を書き出す"""
    srcs, cands = resolve_src(src)
    if not srcs:
        msg = 'フォルダが見つかりません:\n    ' + '\n    '.join(cands)
        hint = hint_subdirs(cands)
        if hint:
            msg += '\n  ' + hint
        raise SystemExit(msg)

    outdir = Path(outdir or WEB)
    targets = collect_files(srcs)
    if not targets:
        raise SystemExit('稼動日報の .xls が1つも見つかりません:\n    '
                         + '\n    '.join(str(x) for x in srcs))

    rows, daily, cols, files, warnings = read_folder(targets)
    # 見つからなかったフォルダは警告にしない。U: が使えるPCではUNC側が、
    # 使えないPCではU:側が必ず存在しないため、毎回出ると邪魔になる。
    # 「元ファイル」タブでだけ確認できるようにする。
    skipped = [c for c in cands if not Path(c).is_dir()]
    if not rows:
        raise SystemExit('月次シート（「26年1月」形式）が1つも見つかりませんでした: %s' % src)

    too_old = sorted(y for y in rows if y < MIN_YEAR)
    if too_old:
        warnings.append('%s は %d年より前のため読み込みませんでした。'
                        % ('、'.join('%d年' % y for y in too_old), MIN_YEAR))
    years = sorted(y for y in rows if y >= MIN_YEAR)
    if not years:
        raise SystemExit('%d年以降のデータが見つかりませんでした。' % MIN_YEAR)
    if year:
        if year not in years:
            raise SystemExit('%d年 のデータがありません。このフォルダにあるのは %s です。'
                             % (year, '、'.join('%d年' % y for y in years)))
        years = [year]

    outdir.mkdir(parents=True, exist_ok=True)
    source = ' ／ '.join(str(x) for x in srcs)
    summary = []
    for y in years:
        data = build_year(y, rows[y], daily.get(y, {}), cols, files, source, warnings)
        data['sources'] = [str(x) for x in srcs]
        data['skipped'] = skipped
        (outdir / ('data_%d.json' % y)).write_text(
            json.dumps(data, ensure_ascii=False), encoding='utf-8')
        # 前年と比べるための小さな集計。画面はこれだけを使うので、前年の
        # data_<年>.json（10MB前後）を読み込まずに前年比が出せる。
        by_dept = OrderedDict()
        for st in data['stats']:
            by_dept.setdefault(st['dept'], {})[str(st['m'])] = [st['tsuGroups'],
                                                                st['groups']]
        summary.append({
            'year':    y,
            'months':  data['months'],
            'records': len(data['records']),
            'detail':  sum(s['detail'] for s in data['stats']),
            'tsu':     sum(s['tsuGroups'] for s in data['stats']),
            'depts':   list(data['depts']),
            'byDept':  by_dept,          # {部署: {"月": [通し数, 件数]}}
        })

    if not year:
        # 年を絞らずに作り直したときは、今回作らなかった年のファイルを消す。
        # 前回の読み込みで出ていた年が、画面に残り続けないようにするため。
        for p in outdir.glob('data_*.json'):
            try:
                if int(p.stem.split('_')[1]) not in years:
                    p.unlink()
            except (ValueError, IndexError, OSError):
                pass

    manifest = {
        'generated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
        'source':    source,
        'sources':   [str(x) for x in srcs],
        'skipped':   skipped,
        'years':     summary,
        'current':   summary[-1]['year'],
        'files':     files,
        'warnings':  warnings,
    }
    (outdir / 'years.json').write_text(
        json.dumps(manifest, ensure_ascii=False), encoding='utf-8')
    return manifest


def main():
    p = argparse.ArgumentParser(description='稼動日報フォルダから Web アプリ用データを作成します')
    p.add_argument('--src', required=True, action='append',
                   help='稼動日報フォルダ（複数指定可: --src A --src B）')
    p.add_argument('--year', type=int, help='対象年（省略すると入っている年を全部）')
    p.add_argument('--out', default=str(WEB), help='出力先フォルダ（既定 web）')
    a = p.parse_args()

    print('対象フォルダ:')
    for c in (clean_src(x) for x in a.src):
        p_ = Path(c)
        if p_.is_dir():
            print('    %s' % c)
        else:
            hit = match_by_name(p_)
            print('    %s' % c
                  + ('\n      → %s として読み込みます' % hit if hit is not None
                     else '  （見つかりません）'))
    print('対象年      : %s' % ('%d年' % a.year if a.year else 'フォルダ内の全部'))
    m = build(a.src, a.year, a.out)

    print('\n── 集計結果 ' + '─' * 56)
    print('%-8s %8s %8s %14s   %s' % ('年', '明細行', '統合件数', '通し数', '月'))
    for y in m['years']:
        print('%-8s %8d %8d %14s   %s'
              % ('%d年' % y['year'], y['detail'], y['records'], format(y['tsu'], ','),
                 '、'.join('%d月' % x for x in y['months'])))
    if len(m['years']) > 1:
        print('%-8s %8d %8d %14s' % ('合計',
                                     sum(y['detail'] for y in m['years']),
                                     sum(y['records'] for y in m['years']),
                                     format(sum(y['tsu'] for y in m['years']), ',')))
    for w in m['warnings']:
        print('  [注意] %s' % w)
    print('\n出力: %s （years.json と data_<年>.json）' % a.out)


if __name__ == '__main__':
    main()
