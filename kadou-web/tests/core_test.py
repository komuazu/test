# -*- coding: utf-8 -*-
"""抽出・統合と設定まわりの自動テスト（ブラウザ不要）

    python tests/core_test.py

ダミーの稼動日報を作って、確定ルールどおりに読めているかを確かめる。
"""
import datetime
import json
import os
import shutil
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))

import build                                                     # noqa: E402
import server                                                    # noqa: E402
from kadou_core import consolidate, extract_all, sort_groups     # noqa: E402
from make_fixture import main as make_sample, make_multi         # noqa: E402

ok, ng = [], []


def check(cond, msg):
    (ok if cond else ng).append(msg)
    print(('  OK  ' if cond else '  NG  ') + msg)



def check_pick():
    """設定画面の「フォルダを選ぶ」が server と正しくつながっているか

    本物のフォルダ選択ダイアログは人が操作しないと閉じないので、
    ダイアログを出す部分だけ差し替えて、その前後の受け渡しを確かめる。
    """
    import json                                                  # noqa: PLC0415
    import threading                                             # noqa: PLC0415
    import urllib.request                                        # noqa: PLC0415
    from functools import partial                                # noqa: PLC0415
    from http.server import ThreadingHTTPServer                  # noqa: PLC0415

    picked = 'C:' + chr(92) + '選んだフォルダ'
    got = []
    real = server.pick_folder
    server.Handler.cfg = server.load_config()
    server.pick_folder = lambda initial=None: (got.append(initial), picked)[1]

    port = server.free_port(8791)
    srv = ThreadingHTTPServer(('127.0.0.1', port), partial(server.Handler))
    threading.Thread(target=srv.serve_forever, daemon=True).start()

    def post(body):
        req = urllib.request.Request('http://127.0.0.1:%d/api/pick' % port,
                                     data=json.dumps(body).encode('utf-8'),
                                     headers={'Content-Type': 'application/json'})
        return json.loads(urllib.request.urlopen(req).read().decode('utf-8'))

    try:
        # data_<年>.json は10MB前後になる。HTTP/1.0 のままだと送信途中で接続が
        # 切れて、画面が読み込み中のまま止まることがあった（実際に半々で再現）。
        check(server.Handler.protocol_version == 'HTTP/1.1',
              '大きなデータを送り切れるよう HTTP/1.1 で応答する')

        r = post({'initial': 'C:' + chr(92) + 'Users'})
        check(r == {'ok': True, 'path': picked}, '選んだフォルダが画面に返る')
        check(got == ['C:' + chr(92) + 'Users'], '今の設定を開始位置としてダイアログに渡す')

        server.pick_folder = lambda initial=None: None
        check(post({}) == {'ok': True, 'path': None}, '選ばずに閉じたら何もしない')

        def boom(initial=None):
            raise RuntimeError('tkinter なし')
        server.pick_folder = boom
        r = post({})
        check(r['ok'] is False and '直接ご記入' in r['error'],
              'ダイアログを出せない環境では手入力を案内する')
    finally:
        server.pick_folder = real
        srv.shutdown()


def check_koyomi():
    """祝日・稼働日と、稼働率で使う機械名のまとめ方"""
    import koyomi                                                # noqa: PLC0415

    # 内閣府の一覧と突き合わせる（改正をまたぐ年を選んである）
    want = {
        2013: '01-01 元日/01-14 成人の日/02-11 建国記念の日/03-20 春分の日/'
              '04-29 昭和の日/05-03 憲法記念日/05-04 みどりの日/05-05 こどもの日/'
              '05-06 振替休日/07-15 海の日/09-16 敬老の日/09-23 秋分の日/'
              '10-14 体育の日/11-03 文化の日/11-04 振替休日/11-23 勤労感謝の日/'
              '12-23 天皇誕生日',
        2019: '01-01 元日/01-14 成人の日/02-11 建国記念の日/03-21 春分の日/'
              '04-29 昭和の日/04-30 国民の休日/05-01 天皇の即位の日/05-02 国民の休日/'
              '05-03 憲法記念日/05-04 みどりの日/05-05 こどもの日/05-06 振替休日/'
              '07-15 海の日/08-11 山の日/08-12 振替休日/09-16 敬老の日/'
              '09-23 秋分の日/10-14 体育の日/10-22 即位礼正殿の儀の行われる日/'
              '11-03 文化の日/11-04 振替休日/11-23 勤労感謝の日',
        2021: '01-01 元日/01-11 成人の日/02-11 建国記念の日/02-23 天皇誕生日/'
              '03-20 春分の日/04-29 昭和の日/05-03 憲法記念日/05-04 みどりの日/'
              '05-05 こどもの日/07-22 海の日/07-23 スポーツの日/08-08 山の日/'
              '08-09 振替休日/09-20 敬老の日/09-23 秋分の日/11-03 文化の日/'
              '11-23 勤労感謝の日',
        2026: '01-01 元日/01-12 成人の日/02-11 建国記念の日/02-23 天皇誕生日/'
              '03-20 春分の日/04-29 昭和の日/05-03 憲法記念日/05-04 みどりの日/'
              '05-05 こどもの日/05-06 振替休日/07-20 海の日/08-11 山の日/'
              '09-21 敬老の日/09-22 国民の休日/09-23 秋分の日/10-12 スポーツの日/'
              '11-03 文化の日/11-23 勤労感謝の日',
    }
    for y, exp in sorted(want.items()):
        got = sorted('%02d-%02d %s' % (d.month, d.day, n)
                     for d, n in koyomi.holidays(y).items())
        check(got == sorted(exp.split('/')),
              '%d年の祝日が内閣府の一覧と一致（振替休日・国民の休日を含む）' % y)

    check(len(koyomi.workdays(2026, 7)) == 22 and len(koyomi.workdays(2026, 8)) == 20,
          '土日祝を除いた稼働日数（2026年7月22日・8月20日）')
    cl = koyomi.parse_closed(['08-13..15  夏季休業', '2026-12-29  年末年始',
                              '2025-10-05  よその年'], 2026)
    check(sorted(str(d) for d in cl) == ['2026-08-13', '2026-08-14', '2026-08-15',
                                         '2026-12-29'],
          '会社の休業日: 範囲・毎年・その年だけを読み分ける')
    check(len(koyomi.workdays(2026, 8, cl)) == 18,
          '会社の休業日を分母から外す（2026年8月 20日 → 18日）')

    # 年をまたぐ休業日（年末年始）。前の年から続く範囲も拾えること
    cl26 = koyomi.parse_closed(['12-29..01-03  年末年始'], 2026)
    check(sorted(str(d) for d in cl26)
          == ['2026-01-01', '2026-01-02', '2026-01-03',
              '2026-12-29', '2026-12-30', '2026-12-31'],
          '年をまたぐ休業日（12-29..01-03）を前後どちらの年でも拾う')
    cl27 = koyomi.parse_closed(['2026-12-29..2027-01-03  年末年始'], 2027)
    check(sorted(str(d) for d in cl27)
          == ['2027-01-01', '2027-01-02', '2027-01-03'],
          '年を書いた年またぎの範囲も、翌年ぶんだけ正しく拾う')
    check(len(koyomi.workdays(2026, 1, cl26)) == 19,
          '年末年始を分母から外す（2026年1月 20日 → 19日）')

    # 「★新・NN年稼動」と「NN年稼働」の両方から読む年で、同じ機械を1台にまとめる
    check(build.short_machine('新・三菱稼動日報') == build.short_machine('三菱稼動日報')
          == '三菱', '「新・」が付いても同じ機械として数える')
    check(build.short_machine('新26・小森1号機稼動日報') == '小森1号機',
          '機械名から年と「稼動日報」を落とす')


def check_daily_date(sample):
    """日次集計ブロックの行に日付が無くても、直前の明細行の日付を引き継げるか

    本物の稼動日報はブロックの行の日付欄が空になっている。ダミー日報の
    「新26・菊全UV稼動日報.xls」の26年3月シートを日付なしで作ってある。
    """
    rows, daily, cols, files, warns = build.read_folder(build.collect_files([str(sample)]))
    dd = [d for d in daily[2026][3] if d['label'] == '有効時間']
    check(len(dd) >= 1, '日付なしのブロックでも日次集計として拾える')
    check(all(d['date'] == '2026-03-05' for d in dd),
          '直前の明細行の日付(2026-03-05)を引き継ぐ  → %s'
          % [d['date'] for d in dd])
    check(all(d['hours'] == 8.0 for d in dd),
          '日付なしでも有効時間の値を拾える  → %s' % [d['hours'] for d in dd])


def check_waste_parity(sample):
    """損紙率で数える範囲が、機械別の表と案件ごとで一致するか

    統合レコードは部署ごとに分かれるので、同じ管理番号が部署をまたぐと
    「出庫枚数がある案件」の判定がずれ、案件ランキングの合計やれ枚数が
    月合計と食い違う（実データの2015年で4,510枚ぶん落ちていた）。
    レコードの 'ob'（管理番号に出庫があるか。部署をまたいで判定）で
    そろえてある。
    """
    data = build.build([str(sample)], None, str(sample.parent / 'out_parity'))
    for y in data['years']:
        p = sample.parent / 'out_parity' / ('data_%d.json' % y['year'])
        d = json.loads(p.read_text(encoding='utf-8'))
        w = d.get('waste') or {}
        if not w.get('months'):
            continue
        for m, v in w['months'].items():
            tot = sum(x[1] for x in v['tot'].values())
            rank = sum(r['yare'] for r in d['records']
                       if r['m'] == int(m) and r['yare'] > 0 and r.get('ob'))
            check(tot == rank,
                  '%d年%s月: 案件ごとの合計やれ枚数が機械別の表と一致  → %d / %d'
                  % (y['year'], m, rank, tot))


def check_xlsx():
    """画面の表が、そのまま開ける Excel ファイルになるか"""
    import io                                                    # noqa: PLC0415
    import zipfile                                               # noqa: PLC0415
    import xlsx                                                  # noqa: PLC0415

    data = xlsx.build([['部署', '通し数', '前年比'],
                       ['本社営業部', 12345, '89.9%'],
                       ['合計', 12345, '89.9%']], '前年との比較')
    check(data[:2] == b'PK', 'Excelファイル(.xlsx)ができる  → %d バイト' % len(data))
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        names = z.namelist()
        sh = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
        wbx = z.read('xl/workbook.xml').decode('utf-8')
        z_styles = z.read('xl/styles.xml').decode('utf-8')
    check('xl/styles.xml' in names and 'xl/workbook.xml' in names,
          '中身が Excel の形になっている')
    # Excel が保存するときと同じ部品をそろえる。テーマが無いファイルは、
    # 開けても印刷（＝描画）でつまずくことがある。
    check('xl/theme/theme1.xml' in names, 'テーマ(theme1.xml)を同梱する')
    check('docProps/core.xml' in names and 'docProps/app.xml' in names,
          '文書情報(docProps)を同梱する')
    check('<v>12345</v>' in sh, '数値は文字ではなく数値のまま入る')
    check('s="1"' in sh and 'customWidth' in sh, '見出しの飾りと列幅が付く')
    check('autoFilter' in sh and 'state="frozen"' in sh,
          '見出し行の固定とオートフィルタが付く')
    check('前年との比較' in wbx, 'シート名が付く')
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        root = z.read('_rels/.rels').decode('utf-8')
        ct = z.read('[Content_Types].xml').decode('utf-8')
    check('theme/theme1.xml' in rels and 'relationships/theme' in rels,
          'テーマが workbook から参照されている')
    check('docProps/core.xml' in root and 'docProps/app.xml' in root,
          '文書情報が参照されている')
    check('theme+xml' in ct and 'core-properties+xml' in ct,
          '足した部品が [Content_Types].xml に登録されている')
    # 印刷まわり（ここが空だとExcelが自前の既定で組版することになる）
    check('<pageSetup' in sh and 'paperSize="9"' in sh and 'landscape' in sh
          and 'fitToWidth="1"' in sh, '印刷設定が入る（A4横・横1ページ）')
    check('<pageMargins' in sh and 'fitToPage="1"' in sh and '<printOptions' in sh,
          '余白と横1ページの指定が入る')
    check('_xlnm.Print_Titles' in wbx and '$1:$1' in wbx,
          '見出し行が各ページの先頭で繰り返される')
    # 印刷範囲(Print_Area)は入れない。Excel も openpyxl も普段は作らない形で、
    # 入れても出力は変わらないことを実際に印刷して確かめた。印刷でつまずく
    # 元を減らすため、余計なものは書かない。
    check('_xlnm.Print_Area' not in wbx, '印刷範囲(Print_Area)は書かない')
    check('Meiryo' in z_styles and 'charset val="128"' in z_styles,
          '日本語のフォントで書き出す')
    # ヘッダーは開いたときではなく印刷するときに読まれる。フォントの指定を
    # &"名前,字体" の形で書かないと、開けるのに印刷でつまずく。
    check('&amp;"Meiryo,Regular"' in sh,
          'ヘッダーのフォントを &"名前,字体" の形で書く（印刷時に読まれる）')
    check('&amp;P / &amp;N' in sh, 'ヘッダーにページ番号が入る')

    # ヘッダーの中では & が命令の合図なので、文字としての & は && にする
    d3 = xlsx.build([['得意先'], ['A&B商事']], 'A&B 集計')
    with zipfile.ZipFile(io.BytesIO(d3)) as z3:
        sh3 = z3.read('xl/worksheets/sheet1.xml').decode('utf-8')
    check('&amp;&amp;B 集計' in sh3,
          'シート名の & はヘッダーで && に直す（印刷が崩れないように）')
    # 壊れた値が混ざってもファイルとして壊れない
    d2 = xlsx.build([['名前', '数'], ['制御\x0b文字\x00入り', float('nan')],
                     ['ふつう', 12]], '2026/8 [試] 小森1号機')
    with zipfile.ZipFile(io.BytesIO(d2)) as z2:
        sh2 = z2.read('xl/worksheets/sheet1.xml').decode('utf-8')
        wb2 = z2.read('xl/workbook.xml').decode('utf-8')
    check('制御文字入り' in sh2 and '\x0b' not in sh2,
          '制御文字は落とす（混ざるとExcelがファイルを開けない）')
    check('<v>nan</v>' not in sh2, 'nan・inf を数値として書かない')
    check('name="2026-8 試 小森1号機"' in wb2,
          'シート名に使えない文字を直す  → 2026-8 試 小森1号機')


def check_cache(tmp):
    """2回目以降は、変わっていない年を読み直さないか

    去年より前の日報は内容が決まっている。書き変わるのは当月ぶんだけなので、
    2回目からは今年のファイルだけ読めば足りる。実データ（13年・78ファイル）で
    起動が25秒→0秒になった仕組み。速いだけでなく、全部読み直したときと
    中身が同じであることが大事。
    """
    import contextlib                                            # noqa: PLC0415
    import io                                                    # noqa: PLC0415

    src = tmp / 'cache_src'
    make_multi(src)                      # 2023〜2025年、年ごとに2ファイル
    out, full = tmp / 'cache_out', tmp / 'cache_full'

    def run(**kw):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            m = build.build([str(src)], None, str(out), **kw)
        return m, buf.getvalue().count('  読込 ')

    def data(d, y):
        x = json.loads((Path(d) / ('data_%d.json' % y)).read_text(encoding='utf-8'))
        x.pop('generated', None)
        return x

    def tsu(m, y):
        return [x['tsu'] for x in m['years'] if x['year'] == y][0]

    m1, n1 = run()
    check(n1 == 6, '1回目は全部読む  → %d件' % n1)
    base = tsu(m1, 2024)

    m2, n2 = run()
    check(n2 == 0, '2回目は1件も読まない（変わっていないため）  → %d件' % n2)
    check([y['year'] for y in m2['years']] == [y['year'] for y in m1['years']],
          '読まなくても年の一覧は同じ')
    check(tsu(m2, 2024) == base, '読まなくても集計は同じ')

    # ある年の日報が書き変わったら、その年だけ読み直す
    f = src / '★新・24年稼動' / '新24・A全UV稼動日報.xls'
    f.touch()
    m3, n3 = run()
    check(n3 == 2, '書き変わった年のファイルだけ読む（2024年の2件）  → %d件' % n3)
    check(tsu(m3, 2024) == base, '読み直しても集計は変わらない')

    # ファイルが消えたら、その年から外れる
    keep = f.read_bytes()
    f.unlink()
    m4, _ = run()
    check(tsu(m4, 2024) < base, 'ファイルが消えたら、その年の集計が減る')

    # 戻したら（＝新しいファイルが増えたのと同じ）、元に戻る
    f.write_bytes(keep)
    m5, n5 = run()
    check(tsu(m5, 2024) == base,
          'ファイルが増えたら、その年を作り直して元に戻る  → %s' % format(tsu(m5, 2024), ','))
    check(n5 == 2, '増えた年のファイルをそろえて読む  → %d件' % n5)

    # 出来上がりを消したら作り直す
    (out / 'data_2023.json').unlink()
    run()
    check((out / 'data_2023.json').exists(), '出来上がりを消したら作り直す')

    # 控えが壊れていたら全部読み直す
    (out / 'cache.json').write_text('こわれています', encoding='utf-8')
    _m, n6 = run()
    check(n6 == 6, '控えが壊れていたら全部読み直す  → %d件' % n6)

    # 読込元フォルダを変えたら、作り直さない年があっても一覧に出る。
    # 読込元は年ごとではなく「今回の読み込み」の情報なので、年のデータの中の
    # 古い値ではなく years.json を見る必要がある。
    _m7, n7 = run()
    check(n7 == 0, '（下の確認の前に）読み直しは起きていない  → %d件' % n7)
    with contextlib.redirect_stdout(io.StringIO()):
        m8 = build.build([str(src), str(tmp / 'ありません')], None, str(out))
    check(str(tmp / 'ありません') in (m8.get('skipped') or []),
          '無かったフォルダは、作り直さない年があっても一覧に出る')
    check(m8.get('sources') == [str(src)], '読込元フォルダも今回のものになる')

    # 何より大事: 途中で作り直した結果が、全部読み直したときと同じであること
    with contextlib.redirect_stdout(io.StringIO()):
        build.build([str(src)], None, str(full), fresh=True)
    years = sorted(int(p.stem.split('_')[1]) for p in Path(full).glob('data_*.json'))
    diff = [y for y in years if data(out, y) != data(full, y)]
    check(not diff, '少しずつ作り直した結果が、全部読み直したときと同じ  → %s'
          % (diff or '全年一致'))


def check_shared():
    """共有モード（社内サーバーに置いたとき）の守りが効いているか

    合い言葉を入れた人だけが見られること、画面から稼動日報フォルダを
    変えたり選んだりできないことを確かめる。サーバー上の好きなフォルダを
    読まれてしまわないようにするため。
    """
    import threading                                             # noqa: PLC0415
    import urllib.error                                          # noqa: PLC0415
    import urllib.request                                        # noqa: PLC0415
    from functools import partial                                # noqa: PLC0415
    from http.server import ThreadingHTTPServer                  # noqa: PLC0415

    keep = (server.SHARED, server.PASSWORD, server.TRUST_LOCAL)
    server.SHARED, server.PASSWORD = True, 'あいことば'
    server.TRUST_LOCAL = False        # ほかの端末から来た人として確かめる
    server.SESSIONS.clear()
    server.Handler.cfg = server.load_config()
    port = server.free_port(8795)
    srv = ThreadingHTTPServer(('127.0.0.1', port), partial(server.Handler))
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    base = 'http://127.0.0.1:%d' % port

    def post(path, body, cookie=None):
        req = urllib.request.Request(base + path,
                                     data=json.dumps(body).encode('utf-8'),
                                     headers={'Content-Type': 'application/json'})
        if cookie:
            req.add_header('Cookie', cookie)
        try:
            r = urllib.request.urlopen(req)
            return r.headers, json.loads(r.read().decode('utf-8'))
        except urllib.error.HTTPError as e:
            return e.headers, json.loads(e.read().decode('utf-8'))

    try:
        page = urllib.request.urlopen(base + '/').read().decode('utf-8')
        check('合い言葉' in page and 'app.js' not in page,
              '合い言葉を入れるまで中身を見せない')
        check(post('/api/login', {'password': 'ちがう'})[1]['ok'] is False,
              '違う合い言葉では入れない')
        hd, j = post('/api/login', {'password': 'あいことば'})
        cookie = hd.get('Set-Cookie', '').split(';')[0]
        check(j['ok'] and cookie.startswith('kadou='), '合い言葉が合えば入れる')
        page = urllib.request.urlopen(
            urllib.request.Request(base + '/', headers={'Cookie': cookie})
        ).read().decode('utf-8')
        check('app.js' in page, '入ったあとは画面が出る')
        check(post('/api/pick', {}, cookie)[1]['ok'] is False,
              '共有のときは「フォルダを選ぶ」を使えない')
        check(post('/api/memo', {}, None)[1]['ok'] is False,
              '合い言葉なしでは記入欄も触れない')

        # サーバーにしているPC自身からは、これまでどおり
        server.TRUST_LOCAL = True
        page = urllib.request.urlopen(base + '/').read().decode('utf-8')
        check('app.js' in page, 'サーバーのPC自身からは合い言葉なしで開ける')
        check(post('/api/pick', {})[1].get('error', '').find('共有') < 0,
              'サーバーのPC自身からはフォルダを選べる')
        server.TRUST_LOCAL = False
    finally:
        srv.shutdown()
        server.SHARED, server.PASSWORD, server.TRUST_LOCAL = keep
        server.SESSIONS.clear()


def check_memo(tmp):
    """記入欄の内容が、アプリのフォルダの外にも控えられるか"""
    keep = (server.MEMO, server.MEMO_PREV, server.MEMO_HOME)
    server.MEMO = tmp / 'memo.json'
    server.MEMO_PREV = tmp / 'memo_前回.json'
    server.MEMO_HOME = tmp / 'ひかえ' / 'memo.json'
    try:
        server.save_memo({'a': {'trend': '継続'}})
        server.save_memo({'a': {'trend': '継続'}, 'b': {'trend': '増加'}})
        check(server.MEMO.exists() and server.MEMO_HOME.exists(),
              '記入欄をアプリのフォルダとユーザーフォルダの両方に書く')
        check(len(json.loads(server.MEMO_PREV.read_text(encoding='utf-8'))) == 1,
              '上書きする前の内容も残す（打ち間違いの戻し用）')
        server.MEMO.unlink()                       # フォルダを入れ替えた状況
        check(len(server.load_memo()) == 2,
              'アプリのフォルダ側が無くなっても控えから戻せる')
    finally:
        server.MEMO, server.MEMO_PREV, server.MEMO_HOME = keep


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

        # ── 引用符付きで貼られたパス（エクスプローラーの「パスのコピー」） ──
        quoted = '"' + str(wave) + '"'
        check([str(x) for x in build.resolve_src([quoted])[0]] == [str(wave)],
              '前後に " が付いたパスでもフォルダを見つけられる')
        check(build.clean_src('  ' + quoted + '  ') == str(wave),
              'パスの前後の空白と引用符を落とす')
        check(server.clean_src(quoted) == build.clean_src(quoted),
              'server.py も build.py と同じようにパスを整える')
        # ── 年フォルダごとに1セットだけ読む（同じ日報の写しを数えない） ──
        base = build.collect_files([multi])
        check(len(base) == 6 and all(k == build.NEW for _p, k in base),
              '★新の年フォルダ直下だけを読む  → %d件' % len(base))

        shutil.copytree(multi / '★新・23年稼動', multi / '☆控え' / '★新・23年稼動')
        shutil.copytree(multi / '★新・24年稼動', multi / '★新・23年稼動' / '日報取込用')
        check(len(build.collect_files([multi])) == 6,
              '控えフォルダと「日報取込用」の写しは読まない  → %d件'
              % len(build.collect_files([multi])))

        # ── ★新 が正。古いフォルダに同じ月があっても増えない ──
        a = build.build([str(multi)], 2023, tmp / 'out_a')['years'][0]
        (multi / '23年稼働').mkdir()
        shutil.copy(multi / '★新・23年稼動' / '新23・A全UV稼動日報.xls',
                    multi / '23年稼働' / '三菱稼動日報.xls')
        b = build.build([str(multi)], 2023, tmp / 'out_b')['years'][0]
        check(a['tsu'] == b['tsu'] and a['detail'] == b['detail'],
              '古いフォルダに同じ月があっても通し数は増えない  → %s'
              % format(b['tsu'], ','))

        part = {build.NEW: ({2015: {8: ['新8月']}}, {}),
                build.OLD: ({2015: {7: ['旧7月'], 8: ['旧8月']}}, {})}
        rows, _daily, notes = build.prefer_new(part)
        check(rows[2015][8] == ['新8月'] and rows[2015][7] == ['旧7月'],
              '重なる月は★新、★新に無い月は古いフォルダから採る')
        check(notes and '2015年の8月' in notes[0], '入れ替えた月をお知らせに出す')

        # ── 入れ物のフォルダ名が変わっても年フォルダを見つける ──
        box = tmp / '入れ物' / '13年～26年稼働'          # 来年こう変わる想定
        shutil.copytree(multi / '★新・25年稼動', box / '★新・25年稼動')
        got = build.collect_files([tmp / '入れ物'])
        check(len(got) == 2 and all(k == build.NEW for _p, k in got),
              '入れ物のフォルダ名を見ずに年フォルダを見つける  → %d本' % len(got))

        shutil.copytree(multi / '★新・25年稼動', tmp / '入れ物' / '07年稼動')
        check(len(build.collect_files([tmp / '入れ物'])) == 2,
              '%d年より前の年フォルダは読まない' % build.MIN_YEAR)

        # 同じ年が別のフォルダにもあるとき（中身は同じでも日時が違う写し）
        other = tmp / 'もう一つ'
        shutil.copytree(multi / '★新・25年稼動', other / '★新・25年稼動')
        for f in (other / '★新・25年稼動').glob('*.xls'):
            os.utime(f, (1000000000, 1000000000))
        got = build.collect_files([tmp / '入れ物', other])
        check(len(got) == 2, '同じ年が別のフォルダにもあるときは1か所だけ読む  → %d本'
              % len(got))
        check(all('もう一つ' in str(x) for x, _k in got),
              '浅いところにある年フォルダを使う（控えは階層が深いので外れる）')

        # 深さが同じなら、設定で先に書いたフォルダを使う
        other2 = tmp / 'さらに別'
        shutil.copytree(multi / '★新・25年稼動', other2 / '★新・25年稼動')
        for f in (other2 / '★新・25年稼動').glob('*.xls'):
            os.utime(f, (1100000000, 1100000000))
        got = build.collect_files([other, other2])
        check(len(got) == 2 and all('もう一つ' in str(x) for x, _k in got),
              '深さが同じなら、設定で先に書いたフォルダを使う')

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
        check(len(rows) == 5, '26年1月の明細は5行（日次集計ブロックは除く）')
        check(all(r['tsu'] != 999999 for r in rows),
              '通し枚数は本刷ブロックを採る（集計ブロックの値を拾わない）')
        check([r['dept'] for r in rows if r['code'] == 9999] == ['その他'],
              '3営業部のどれでもないコード9999は「その他」に入る')
        check(sum(1 for r in rows if r['dept'] != 'その他') == 4,
              '3営業部の行はこれまでどおり4行')

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
        check(y26['tsu'] == 1874777,
              '2026年の通し数 1,874,777（その他77,777を含む）  → %s'
              % format(y26['tsu'], ','))
        check(y26['detail'] == 71 and y26['records'] == 69,
              '2026年は明細71行 → 69件に統合')

        man = json.loads((out / 'years.json').read_text(encoding='utf-8'))
        y26s = [y for y in man['years'] if y['year'] == 2026][0]
        tot = sum(v[0] for dd in y26s['byDept'].values() for v in dd.values())
        cnt = sum(v[1] for dd in y26s['byDept'].values() for v in dd.values())
        check(tot == y26s['tsu'],
              '前年比較用の月×部署の集計が年間通し数と一致  → %s' % format(tot, ','))
        check(cnt == y26s['records'], '同じく件数も年間件数と一致  → %s' % format(cnt, ','))

        d26 = json.loads((out / 'data_2026.json').read_text(encoding='utf-8'))
        check(d26['depts'][-1] == 'その他', '並びの最後が「その他」')
        check(d26['depts'] == ['本社営業部', '東京営業部', '池袋営業部',
                               '生産管理部（工務）', 'その他'],
              '営業部の並び  → %s' % '、'.join(d26['depts']))
        from kadou_core import CODE2DEPT                          # noqa: PLC0415
        check(CODE2DEPT.get(6930) == '生産管理部（工務）',
              'ｺｰﾄﾞ6930は生産管理部（工務）')
        check(d26['deptCodes']['その他'] == [9999],
              '「その他」に入ったコードを画面用に持たせる  → %s'
              % d26['deptCodes']['その他'])
        check(sum(r['tsu'] for r in d26['records'] if r['dept'] == 'その他') == 77777,
              '「その他」の通し数が集計に入る')

        # 統合の前後で通し数が変わらないこと（全年）
        bad = []
        for y in years:
            d = json.loads((out / ('data_%d.json' % y)).read_text(encoding='utf-8'))
            for st in d['stats']:
                if st['tsuDetail'] != st['tsuGroups']:
                    bad.append((y, st['m'], st['dept']))
        check(not bad, '全年・全月で統合前後の通し数が一致')

        # ── 「フォルダを選ぶ」の配線（ダイアログは出さずに差し替えて確かめる） ──
        check_pick()

        # ── Excel 書き出しと、記入欄の控え ──
        check_koyomi()
        check_daily_date(sample)
        check_waste_parity(sample)
        check_xlsx()
        check_memo(tmp)
        check_shared()

        # ── 2回目以降の読み込みを省くしくみ ──
        check_cache(tmp)

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
