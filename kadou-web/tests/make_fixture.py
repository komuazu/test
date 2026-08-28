# -*- coding: utf-8 -*-
"""検証用のダミー稼動日報 .xls を作る（実データの構造を模したテスト用）"""
import datetime
import sys
from pathlib import Path

import xlwt

EPOCH = datetime.date(1899, 12, 30)


def serial(d):
    return (d - EPOCH).days


def sheet(wb, title, rows, with_client=True, daily_block=True):
    """rows: (日, 管理番号, ｸﾗｲｱﾝﾄ名, 品名, ｺｰﾄﾞ, 通し, 表, 裏)"""
    ws = wb.add_sheet(title)
    ws.write(0, 0, '稼動日報')
    head = ['日付', '管理番号'] + (['ｸﾗｲｱﾝﾄ名'] if with_client else []) \
        + ['品名', '営業担当ｺｰﾄﾞ', '通し枚数', '表版数', '裏版数']
    for c, h in enumerate(head):
        ws.write(3, c, h)
    # 集計ブロック側にも同名の列を置く（本刷ブロックが優先されることの確認用）
    for c, h in zip((20, 21, 22), ('通し枚数', '表版数', '裏版数')):
        ws.write(3, c, h)

    r = 4
    for d, no, client, name, code, tsu, f, b in rows:
        c = 0
        ws.write(r, c, serial(d)); c += 1
        ws.write(r, c, no); c += 1
        if with_client:
            ws.write(r, c, client or ''); c += 1
        ws.write(r, c, name); c += 1
        ws.write(r, c, code); c += 1
        ws.write(r, c, tsu); c += 1
        ws.write(r, c, f); c += 1
        ws.write(r, c, b)
        # 集計ブロックにはダミー値（採用されてはいけない）
        ws.write(r, 20, 999999); ws.write(r, 21, 99); ws.write(r, 22, 99)
        r += 1
    if daily_block:
        # 日次集計ブロック: 27列目にラベル。明細として拾われてはいけない
        for lbl in ('有効時間', '準備合計', '色合わせ', '印刷時間', 'その他', '受注件数'):
            ws.write(r, 0, serial(rows[0][0]))
            ws.write(r, 27, lbl)
            ws.write(r, 5, 123456)
            r += 1
    return ws


def main(outdir):
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    D = datetime.date

    # ── 1号機 ──────────────────────────────
    wb = xlwt.Workbook(encoding='utf-8')
    sheet(wb, '26年1月', [
        (D(2026, 1, 6),  '8632175-1',   '㈱アルファ商事', 'カタログ 春号',     1110, 12000, 4, 4),
        (D(2026, 1, 7),  '8632175-1-2', '（株）アルファ商事', 'カタログ 春号 追刷', 1110,  8000, 4, 0),
        (D(2026, 1, 6),  '8700001',     'ベータ工業㈱',   'パンフレット',       2100, 30000, 4, 4),
        (D(2026, 1, 9),  '8700002',     '',               '封筒',               3810,  5000, 2, 0),
        (D(2026, 1, 9),  '8700003',     'ガンマ社',       '対象外の部署',       9999, 77777, 4, 4),
    ])
    sheet(wb, '26年1月30', [                       # 日付付き補助シート → 無視されるべき
        (D(2026, 1, 30), '9999999', 'ダミー', '拾われてはいけない', 1110, 111111, 4, 4),
    ])
    sheet(wb, '26年2月', [
        (D(2026, 2, 3),  '8700010', 'デルタ印刷', 'ポスター', 1120, 20000, 6, 0),
        (D(2026, 2, 17), '8700011', 'ベータ工業㈱', 'DM',      2140, 45000, 4, 4),
    ], with_client=False)                          # ｸﾗｲｱﾝﾄ名列が無い月
    sheet(wb, '25年12月', [                        # 前年 → 対象外
        (D(2025, 12, 2), '8600000', '前年社', '前年案件', 1110, 999, 4, 0),
    ])
    wb.save(str(outdir / '新26・A全UV稼動日報.xls'))

    # ── 2号機 ──────────────────────────────
    wb = xlwt.Workbook(encoding='utf-8')
    sheet(wb, '26年1月', [
        # 同一管理番号が複数日にまたがる（印刷日=初日・複数日フラグ）
        (D(2026, 1, 20), '8700100', 'イプシロン㈱', '会報誌 表',   1110, 15000, 4, 0),
        (D(2026, 1, 21), '8700100', 'イプシロン㈱', '会報誌 裏',   1110, 15000, 0, 4),
        (D(2026, 1, 22), '8700200', 'ゼータ',       'チラシ',       3820,  9000, 4, 4),
    ])
    sheet(wb, '26年3月', [
        (D(2026, 3, 5), '8700300', 'イータ商会', '取扱説明書', 2100, 60000, 2, 2),
    ])
    wb.save(str(outdir / '新26・菊全UV稼動日報.xls'))

    # ── 3号機: 4〜8月のまとまったデータ（画面確認用） ──
    wb = xlwt.Workbook(encoding='utf-8')
    clients = ['㈱アルファ商事', 'ベータ工業㈱', 'デルタ印刷', 'イプシロン㈱', 'イータ商会', 'シータ物産']
    codes = [1110, 1120, 2100, 2140, 3810, 3820]
    items = ['カタログ', 'パンフレット', 'DM', 'ポスター', '会報誌', '取扱説明書']
    for mo in range(4, 9):
        rows = []
        for i in range(12):
            day = 1 + (i * 2 + mo) % 26
            rows.append((D(2026, mo, day),
                         '87%03d%02d' % (mo * 7 + i, i),
                         clients[(mo + i) % len(clients)],
                         '%s %d月号' % (items[(mo * 2 + i) % len(items)], mo),
                         codes[(mo + i * 5) % len(codes)],
                         3000 + ((mo * 137 + i * 911) % 47) * 1000,
                         [2, 4, 4, 6][(i + mo) % 4],
                         [0, 0, 4, 4][(i + mo) % 4]))
        sheet(wb, '26年%d月' % mo, rows)
    wb.save(str(outdir / '新26・小森1号機稼動日報.xls'))

    print('作成:', outdir)


def make_multi(outdir):
    """複数年（年ごとのサブフォルダ）を模したダミーを作る

    「13年稼動～25年稼動」のように、1つのフォルダの下に年別フォルダが並ぶ形。
    """
    outdir = Path(outdir)
    D = datetime.date
    clients = ['㈱アルファ商事', 'ベータ工業㈱', 'デルタ印刷', 'イプシロン㈱', 'イータ商会']
    codes = [1110, 1120, 2100, 2140, 3810, 3820]
    items = ['カタログ', 'パンフレット', 'DM', 'ポスター', '会報誌']

    for year in (2023, 2024, 2025):
        d = outdir / ('★新・%d年稼動' % (year % 100))
        d.mkdir(parents=True, exist_ok=True)
        for mi, machine in enumerate(('A全UV', '菊全UV')):
            wb = xlwt.Workbook(encoding='utf-8')
            for mo in range(1, 13):
                rows = []
                for i in range(5):
                    rows.append((D(year, mo, 1 + (i * 3 + mo) % 26),
                                 '%d%02d%02d%d' % (year % 100, mo, i, mi),
                                 clients[(mo + i + mi) % len(clients)],
                                 '%s %d月号' % (items[(mo + i) % len(items)], mo),
                                 codes[(mo * 2 + i + mi) % len(codes)],
                                 2000 + ((year + mo * 137 + i * 911) % 40) * 1000,
                                 [2, 4, 4, 6][(i + mo) % 4],
                                 [0, 0, 4, 4][(i + mo) % 4]))
                sheet(wb, '%d年%d月' % (year % 100, mo), rows)
            wb.save(str(d / ('新%d・%s稼動日報.xls' % (year % 100, machine))))
    print('作成（複数年）:', outdir)


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[2] == '--multi':
        make_multi(sys.argv[1])
    else:
        main(sys.argv[1] if len(sys.argv) > 1 else 'tests/sample')
