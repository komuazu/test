# -*- coding: utf-8 -*-
"""
稼動日報フォルダ → Web アプリ用データ(data.json) 生成

  python build.py --src "U:\\製造本部\\...\\★新・26年稼動" --year 2026

フォルダ内の .xls をすべて読み、その年の月次シート（例「26年1月」）を全部拾って
営業部別・月別に統合したデータを web/data.json に書き出す。

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
from collections import OrderedDict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from kadou_core import DEPTS, consolidate, extract_year, sort_groups   # noqa: E402

HERE = Path(__file__).resolve().parent


def find_xls(src):
    """対象フォルダ配下の .xls を集める（Excel の一時ファイルは除外）"""
    files = []
    for p in sorted(Path(src).rglob('*.xls*')):
        if p.name.startswith('~$') or p.suffix.lower() not in ('.xls', '.xlsx', '.xlsm'):
            continue
        files.append(p)
    return files


def build(src, year, out):
    src = Path(src)
    if not src.is_dir():
        raise SystemExit('フォルダが見つかりません: %s' % src)

    year2 = year % 100
    files = find_xls(src)
    if not files:
        raise SystemExit('.xls ファイルが1つも見つかりません: %s' % src)

    by_month = OrderedDict()      # 月 → 明細行
    daily_by_month = OrderedDict()  # 月 → 日次集計行（有効時間・準備合計 など）
    all_cols = []                 # 日報の見出し列（出てきた順）
    file_info, warnings = [], []

    for path in files:
        if path.suffix.lower() != '.xls':
            warnings.append('%s は旧形式(.xls)ではないため読み飛ばしました。' % path.name)
            continue
        try:
            got, days, cols = extract_year(path, year2)
        except Exception as e:                                   # noqa: BLE001
            warnings.append('%s を読めませんでした: %s' % (path.name, e))
            traceback.print_exc(file=sys.stderr)
            continue
        for n_ in cols:
            if n_ not in all_cols:
                all_cols.append(n_)
        n = 0
        for month, rows in got.items():
            by_month.setdefault(month, []).extend(rows)
            n += len(rows)
        for month, dd in days.items():
            daily_by_month.setdefault(month, []).extend(dd)
        if not got:
            warnings.append('%s に %d年 の月次シートがありません。' % (path.name, year))
        file_info.append({'name': path.name,
                          'machine': path.stem,
                          'path': str(path),
                          'months': sorted(got),
                          'rows': n})
        print('  読込 %-40s 月:%-28s 明細%5d行'
              % (path.name, ','.join('%d月' % m for m in sorted(got)) or 'なし', n))

    if not by_month:
        raise SystemExit('%d年 の該当データが1件も見つかりませんでした。' % year)

    records, stats = [], []
    for month in sorted(by_month):
        rows = by_month[month]
        rows.sort(key=lambda x: (x['date'], str(x['no']), x['seq']))
        for dept in DEPTS:
            drows = [r for r in rows if r['dept'] == dept]
            groups = sort_groups(consolidate(drows))
            for i, g in enumerate(groups):
                d0 = min(g['dates'])
                records.append({
                    'i':       len(records),
                    'm':       month,
                    'dept':    dept,
                    'ord':     i,
                    'code':    g['code'],
                    'no':      g['base'],
                    'client':  g['client'] or '',
                    'name':    g['name'] or '',
                    'date':    d0.isoformat(),
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

    bad = [s for s in stats if s['tsuDetail'] != s['tsuGroups']]
    if bad:
        warnings.append('統合前後で通し数が一致しない月があります: '
                        + ', '.join('%d月/%s' % (s['m'], s['dept']) for s in bad))

    data = {
        'year':      year,
        'generated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
        'source':    str(src),
        'depts':     list(DEPTS),
        'deptCodes': {d: cs for d, cs in DEPTS.items()},
        'months':    sorted(by_month),
        'files':     file_info,
        'stats':     stats,
        'warnings':  warnings,
        'cols':      all_cols,
        'daily':     [dict(d, m=m) for m in sorted(daily_by_month) for d in daily_by_month[m]],
        'records':   records,
    }
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')
    return data


def main():
    p = argparse.ArgumentParser(description='稼動日報フォルダから Web アプリ用データを作成します')
    p.add_argument('--src', required=True, help='稼動日報フォルダ（例: ★新・26年稼動）')
    p.add_argument('--year', type=int, default=2026, help='対象年（既定 2026）')
    p.add_argument('--out', default=str(HERE / 'web' / 'data.json'), help='出力先 data.json')
    a = p.parse_args()

    print('対象フォルダ: %s' % a.src)
    print('対象年      : %d年' % a.year)
    data = build(a.src, a.year, a.out)

    print('\n── 集計結果 ' + '─' * 52)
    print('%-6s %8s %8s %14s' % ('月', '明細行', '統合件数', '通し数'))
    for month in data['months']:
        ss = [s for s in data['stats'] if s['m'] == month]
        print('%-6s %8d %8d %14s' % ('%d月' % month,
                                     sum(s['detail'] for s in ss),
                                     sum(s['groups'] for s in ss),
                                     format(sum(s['tsuGroups'] for s in ss), ',')))
    print('%-6s %8d %8d %14s' % ('合計',
                                 sum(s['detail'] for s in data['stats']),
                                 sum(s['groups'] for s in data['stats']),
                                 format(sum(s['tsuGroups'] for s in data['stats']), ',')))
    for w in data['warnings']:
        print('  [注意] %s' % w)
    print('\n出力: %s' % a.out)


if __name__ == '__main__':
    main()
