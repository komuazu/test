# -*- coding: utf-8 -*-
"""抽出・統合まわりの自動テスト（ブラウザ不要）

    python tests/core_test.py

ダミーの稼動日報を作って、確定ルールどおりに読めているかを確かめる。
"""
import datetime
import shutil
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))

import build                                                     # noqa: E402
from kadou_core import consolidate, extract_all, sort_groups     # noqa: E402
from make_fixture import main as make_sample, make_multi         # noqa: E402

ok, ng = [], []


def check(cond, msg):
    (ok if cond else ng).append(msg)
    print(('  OK  ' if cond else '  NG  ') + msg)


def main():
    tmp = Path(tempfile.mkdtemp(prefix='kadou_test_'))
    try:
        sample, multi = tmp / 'sample', tmp / 'multi'
        make_sample(sample)
        make_multi(multi)

        # ── フォルダ名の表記ゆれ（全角チルダ／波ダッシュ） ──
        wave = tmp / '13年稼動〜25年稼動'          # 実在するのは波ダッシュ U+301C
        shutil.copytree(multi, wave)
        tilde = str(tmp / '13年稼動～25年稼動')     # 指定は全角チルダ U+FF5E
        found, _ = build.resolve_src([tilde])
        check([str(x) for x in found] == [str(wave)],
              '波ダッシュ「〜」のフォルダを全角チルダ「～」の指定で見つけられる')
        check(build.fold_name('13年稼動～25年') == build.fold_name('13年稼動〜25年'),
              '「～」と「〜」を同じ名前として扱う')
        check(build.resolve_src([str(tmp / 'ありません')])[0] == [],
              '本当に無いフォルダは見つけない')

        # ── 同じファイルを二重に数えない ──
        dup = tmp / 'dup'
        shutil.copytree(sample, dup)
        one = build.find_xls([sample])
        two = build.find_xls([sample, dup])
        check(len(one) == len(two) == 3,
              '同じ内容のフォルダを2つ指定してもファイルは3件のまま  → %d / %d'
              % (len(one), len(two)))

        # ── 年をまたいで読む ──
        got, days, cols = extract_all(sample / '新26・A全UV稼動日報.xls')
        check(sorted(got) == [25, 26], '1つのファイルから25年と26年の両方を読む')
        check(sorted(got[26]) == [1, 2], '26年は1月と2月')
        check('26年1月30' not in str(cols), '日付付きの補助シートは対象外')
        check(any(d['label'] == '有効時間' for d in days[26][1]),
              '日次集計ブロック（有効時間など）を別に拾う')

        rows = got[26][1]
        check(len(rows) == 4, '26年1月の明細は4行（対象外コード9999と日次集計は除く）')
        check(all(r['tsu'] != 999999 for r in rows),
              '通し枚数は本刷ブロックを採る（集計ブロックの値を拾わない）')
        check(all(r['code'] in (1110, 1120, 2100, 2140, 3810, 3820) for r in rows),
              '対象の営業担当ｺｰﾄﾞだけ')

        # ── 管理番号の統合ルール ──
        gs = sort_groups(consolidate([r for r in rows if r['dept'] == '本社営業部']))
        g = [x for x in gs if x['base'] == '8632175'][0]
        check(g['tsu'] == 20000, '枝番を統合して通し数は合計  → %d' % g['tsu'])
        check(len(g['nos']) == 2, '8632175-1 と 8632175-1-2 が1件にまとまる')
        check((g['f'], g['b']) == (4, 4), '色数は表・裏それぞれの最大値  → %d/%d' % (g['f'], g['b']))
        check(min(g['dates']) == datetime.date(2026, 1, 6), '印刷日はグループ内の初日')
        check(g['client'] == '（株）アルファ商事',
              'ｸﾗｲｱﾝﾄ名は同数なら長い方  → %s' % g['client'])
        check(g['name'] == 'カタログ 春号', '品名は初日の行のもの  → %s' % g['name'])

        # ── フォルダ全体のビルド ──
        out = tmp / 'out'
        m = build.build([str(multi), str(sample)], None, out)
        years = [y['year'] for y in m['years']]
        check(years == [2023, 2024, 2025, 2026], '2フォルダから4年ぶんできる  → %s' % years)
        check(all((out / ('data_%d.json' % y)).exists() for y in years),
              '年ごとに data_<年>.json ができる')
        check((out / 'years.json').exists(), 'years.json ができる')

        y26 = [y for y in m['years'] if y['year'] == 2026][0]
        check(y26['tsu'] == 1797000, '2026年の通し数 1,797,000  → %s' % format(y26['tsu'], ','))
        check(y26['detail'] == 70 and y26['records'] == 68,
              '2026年は明細70行 → 68件に統合')

        # 統合の前後で通し数が変わらないこと（全年）
        import json
        bad = []
        for y in years:
            d = json.loads((out / ('data_%d.json' % y)).read_text(encoding='utf-8'))
            for st in d['stats']:
                if st['tsuDetail'] != st['tsuGroups']:
                    bad.append((y, st['m'], st['dept']))
        check(not bad, '全年・全月で統合前後の通し数が一致')

        # ── 元ファイルを変更していないこと ──
        before = sorted((p.name, p.stat().st_size, int(p.stat().st_mtime))
                        for p in build.find_xls([sample, multi]))
        build.build([str(multi), str(sample)], None, out)
        after = sorted((p.name, p.stat().st_size, int(p.stat().st_mtime))
                       for p in build.find_xls([sample, multi]))
        check(before == after, '読み込んでも元の .xls は変わらない（サイズ・更新日時）')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print('\n合格 %d / 不合格 %d' % (len(ok), len(ng)))
    return 1 if ng else 0


if __name__ == '__main__':
    sys.exit(main())
