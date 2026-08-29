# -*- coding: utf-8 -*-
"""抽出・統合と設定まわりの自動テスト（ブラウザ不要）

    python tests/core_test.py

ダミーの稼動日報を作って、確定ルールどおりに読めているかを確かめる。
"""
import datetime
import json
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

        d26 = json.loads((out / 'data_2026.json').read_text(encoding='utf-8'))
        check(d26['depts'][-1] == 'その他', '営業部の並びの最後が「その他」')
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
