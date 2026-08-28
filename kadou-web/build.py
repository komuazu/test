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
import sys
import traceback
import unicodedata
from collections import OrderedDict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from kadou_core import DEPTS, consolidate, extract_all, sort_groups   # noqa: E402

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
    cands = [c for c in cands if c]
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


def read_folder(srcs):
    """フォルダ配下の .xls を1回ずつ開いて、年→月→明細行 にまとめる

    戻り値: (rows, daily, cols, file_info, warnings)
      rows  = {年4桁: {月: [明細行, ...]}}
      daily = {年4桁: {月: [日次集計行, ...]}}
    """
    rows, daily, cols = OrderedDict(), OrderedDict(), []
    file_info, warnings = [], []

    for path in find_xls(srcs):
        if path.suffix.lower() != '.xls':
            warnings.append('%s は旧形式(.xls)ではないため読み飛ばしました。' % path.name)
            continue
        try:
            got, days, cc = extract_all(path)
        except Exception as e:                                   # noqa: BLE001
            warnings.append('%s を読めませんでした: %s' % (path.name, e))
            traceback.print_exc(file=sys.stderr)
            continue

        for n in cc:
            if n not in cols:
                cols.append(n)

        # このファイル自身の年別の明細行数を数える（他のファイルの分と混ぜない）
        per_year, total = OrderedDict(), 0
        for year2, months in got.items():
            year = 2000 + year2
            e = per_year.setdefault(year, {'months': [], 'rows': 0})
            for month, rr in months.items():
                rows.setdefault(year, OrderedDict()).setdefault(month, []).extend(rr)
                e['months'].append(month)
                e['rows'] += len(rr)
                total += len(rr)
        for year2, months in days.items():
            for month, dd in months.items():
                daily.setdefault(2000 + year2, OrderedDict()).setdefault(month, []).extend(dd)

        if not per_year:
            warnings.append('%s に月次シート（「26年1月」形式）がありません。' % path.name)
        file_info.append({'name': path.name, 'machine': path.stem, 'path': str(path),
                          'years': {str(y): {'months': sorted(e['months']), 'rows': e['rows']}
                                    for y, e in per_year.items()},
                          'rows': total})
        print('  読込 %-40s %-24s 明細%6d行'
              % (path.name,
                 '、'.join('%d年' % y for y in sorted(per_year)) or '月次シートなし',
                 total))
    return rows, daily, cols, file_info, warnings


def build_year(year, by_month, daily_by_month, cols, files, src, warnings):
    """1年ぶんのデータを組み立てる（統合ルールは kadou-report と同じ）"""
    records, stats = [], []
    for month in sorted(by_month):
        rows = by_month[month]
        rows.sort(key=lambda x: (x['date'], str(x['no']), x['seq']))
        for dept in DEPTS:
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
        'depts':     list(DEPTS),
        'deptCodes': {d: cs for d, cs in DEPTS.items()},
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
    if not find_xls(srcs):
        raise SystemExit('.xls ファイルが1つも見つかりません:\n    '
                         + '\n    '.join(str(x) for x in srcs))

    rows, daily, cols, files, warnings = read_folder(srcs)
    # 見つからなかったフォルダは警告にしない。U: が使えるPCではUNC側が、
    # 使えないPCではU:側が必ず存在しないため、毎回出ると邪魔になる。
    # 「元ファイル」タブでだけ確認できるようにする。
    skipped = [c for c in cands if not Path(c).is_dir()]
    if not rows:
        raise SystemExit('月次シート（「26年1月」形式）が1つも見つかりませんでした: %s' % src)

    years = sorted(rows)
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
        summary.append({
            'year':    y,
            'months':  data['months'],
            'records': len(data['records']),
            'detail':  sum(s['detail'] for s in data['stats']),
            'tsu':     sum(s['tsuGroups'] for s in data['stats']),
        })

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
    for c in a.src:
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
