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
import koyomi                                                     # noqa: E402
from kadou_core import (ALL_DEPTS, DEPTS, HOURS_LABEL, OTHER, consolidate,  # noqa: E402
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
SHIFT_HOURS = 22         # 1日あたりの就業可能時間（昼勤11h＋夜勤11hの2直）
YARE_COL = 'やれ枚数'      # 損紙。日報の本刷ブロックの列名
OUT_COL = '出庫枚数'       # 実印刷枚数＋基準印刷予備。損紙率の分母
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
    """同じファイル（名前・大きさ・更新日時が同じ）かどうかの目印

    更新日時は秒ではなくナノ秒まで見る。秒までだと、日報を保存した直後に
    読み込んだときに「変わっていない」と誤って判断してしまう。
    2回目以降に読み直すかどうかの判断にも、この目印を使っている。
    """
    try:
        st = p.stat()
        return (p.name, st.st_size, st.st_mtime_ns)
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


def short_machine(name):
    """機械名を短くする（'新26・小森1号機稼動日報' → '小森1号機'）

    ファイル名は年ごとに「新25・」「新26・」と頭が変わるので、年と「稼動日報」を
    落として、画面の見出しが長くならないようにする。

    「★新・NN年稼動」と「NN年稼働」の両方から読む年（2015年など）では、同じ機械が
    「新・三菱稼動日報」と「三菱稼動日報」の2つの名前で入ってくる。頭の「新」は
    年の数字が付かないこともあるので、数字の有無によらず落として1台にまとめる。
    """
    s = re.sub(r'稼[働動](?:日報|実績)', '', str(name))
    s = re.sub(r'^\s*新[・･\-_\s]*(?:(?:20)?\d{2}\s*年?[・･\-_\s]*)?', '', s)
    s = re.sub(r'^\s*(?:20)?\d{2}\s*年?[・･\-_\s]*', '', s)
    return s.strip(' 　・･-_') or str(name).strip()


def build_oper(year, daily_by_month, closed_lines):
    """稼働率のもとになるデータを作る

      分母 = SHIFT_HOURS（22時間） × その月の稼働日数
             稼働日 = 平日 − 祝日（振替休日・国民の休日を含む） − 会社の休業日
      分子 = 稼動日報の「有効時間」の合計（機械が実際に動いていた時間）

    月の合計はシートの月で数える。日ごとの内訳は、日付がその月に入る行だけを
    並べる（月をまたぐ行がわずかにあるため、合計と日ごとの和は必ずしも一致
    しない。画面ではその差を「日付不明」として出す）。
    """
    closed = koyomi.parse_closed(closed_lines, year)
    months, machines = OrderedDict(), set()
    for month in sorted(daily_by_month):
        per_day, tot = OrderedDict(), OrderedDict()
        for d in daily_by_month[month]:
            if d['label'] != HOURS_LABEL or not d.get('hours'):
                continue
            mc = short_machine(d['machine'])
            machines.add(mc)
            tot[mc] = tot.get(mc, 0) + d['hours']
            iso = d.get('date') or ''
            if iso[:7] == '%d-%02d' % (year, month):
                day = per_day.setdefault(int(iso[8:10]), OrderedDict())
                day[mc] = day.get(mc, 0) + d['hours']
        # 土日は画面側で分かるので、平日の休み（祝日・会社の休業日）だけ持たせる
        off = dict((str(d), why) for d, why in koyomi.month_days(year, month, closed)
                   if why and why not in ('土曜', '日曜'))
        months[str(month)] = {
            'work': len(koyomi.workdays(year, month, closed)),
            'off':  off,
            'tot':  dict((k, round(v, 2)) for k, v in tot.items()),
            'days': dict((str(d), dict((k, round(v, 2)) for k, v in h.items()))
                         for d, h in sorted(per_day.items())),
        }
    return {'shift': SHIFT_HOURS, 'machines': sorted(machines), 'months': months}


def col_of(row, name):
    """明細行から数値の列を取る（本刷ブロック側）

    見出しが2か所にある列は2つめに「（集計）」が付くので、付かない方が本刷。
    通し枚数と同じ扱いにそろえてある。
    """
    v = row['raw'].get(name)
    return v if isinstance(v, (int, float)) else 0


def bases_with_out(rows):
    """その月に出庫枚数の入っている管理番号（枝番をまとめた base）

    部署をまたいで同じ管理番号が入力されることがある（2015年6月の 7312579 など）。
    統合レコードは部署ごとに分かれるので、**部署に関係なく管理番号で**判定しないと、
    月別の表と案件ランキングで数える範囲がずれる。
    """
    return {r['base'] for r in rows if col_of(r, OUT_COL)}


def build_waste(year, by_month):
    """損紙率・予備率のもとになるデータ（機械ごと・月ごと・日ごと）

      出庫枚数 = 実印刷枚数 ＋ 基準印刷予備
      損紙率   = やれ枚数 ÷ 出庫枚数
                 （やれ枚数 = 出庫枚数からはみ出して使ってしまった紙）
      予備率   = 基準印刷予備 ÷ 実印刷枚数
                 （基準印刷予備 = 出庫枚数 − 実印刷枚数）

    出庫枚数は用紙を出すたびに1行へまとめて入るので、明細行すべてには入って
    いない。ただしその行にも日付と機械があるので、月別・日別・機械別に足せる。
    2015年より前の日報には出庫枚数の列が無く、その年は損紙率も予備率も出せない。

    実印刷枚数は「出庫枚数の入っている行の通し枚数」を使う。出庫が入らない行
    （同じ紙の2回目以降の刷り）まで足すと、基準印刷予備が合わなくなるため。
    実データでは 出庫枚数 = 通し枚数 + ＪＰ予備紙 が97.3%の行で成り立つ。

    数える範囲は「出庫枚数の入っている案件（管理番号）」だけにそろえる。やれ枚数
    だけ数えて出庫枚数を数えないと、率が跳ね上がってしまうため（出庫枚数がほとんど
    無い2014年で 66429% になっていた）。どれだけの案件を数えられたかは cover に
    入れて、画面で少なすぎるときに知らせる。

    1マスに [出庫枚数, やれ枚数, 通し枚数, 実印刷枚数] を入れる。月ごとの合計
    (tot)と、日ごとの内訳(days)を持たせる。日ごとは明細行の日付で振り分ける
    （月をまたぐ行はその月の合計にだけ入るので、画面では差を「日付不明」に出す）。
    """
    months, machines = OrderedDict(), set()
    njob = njob_out = 0
    for month in sorted(by_month):
        # その月に出庫枚数のある案件（管理番号）を先に拾う
        with_out = bases_with_out(by_month[month])
        jobs = {r['base'] for r in by_month[month]}
        njob += len(jobs)
        njob_out += len(with_out)

        per, days = OrderedDict(), OrderedDict()
        for r in by_month[month]:
            mc = short_machine(r['machine'])
            machines.add(mc)
            if r['base'] not in with_out:
                continue                    # 出庫枚数の無い案件は数えない
            out = col_of(r, OUT_COL)
            v = [out, col_of(r, YARE_COL), r['tsu'] or 0,
                 (r['tsu'] or 0) if out else 0]
            e = per.setdefault(mc, [0] * len(v))
            for i, x in enumerate(v):
                e[i] += x
            d = r['date']
            if d and d.year == year and d.month == month:
                day = days.setdefault(str(d.day), OrderedDict())
                e2 = day.setdefault(mc, [0] * len(v))
                for i, x in enumerate(v):
                    e2[i] += x
        cast = lambda h: dict((k, [int(x) for x in v]) for k, v in h.items())  # noqa: E731
        months[str(month)] = {
            'tot':  cast(per),
            'days': dict((dd, cast(h))
                         for dd, h in sorted(days.items(), key=lambda x: int(x[0]))),
        }
    return {'machines': sorted(machines), 'months': months,
            'cover': round(njob_out / njob, 4) if njob else 0}


def build_year(year, by_month, daily_by_month, cols, files, src, warnings,
               closed=None):
    """1年ぶんのデータを組み立てる（統合ルールは kadou-report と同じ）"""
    records, stats = [], []
    # 「その他」に入った営業担当ｺｰﾄﾞ（画面の見出しに出す）
    other_codes = sorted({r['code'] for rr in by_month.values() for r in rr
                          if r['dept'] == OTHER})
    for month in sorted(by_month):
        rows = by_month[month]
        rows.sort(key=lambda x: (x['date'], str(x['no']), x['seq']))
        # 損紙率の数える範囲（build_waste と同じ判定。部署をまたいでも同じ結果）
        month_out = bases_with_out(rows)
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
                    'yare':    int(sum(col_of(mm, YARE_COL) for mm in g['members'])),
                    'out':     int(sum(col_of(mm, OUT_COL) for mm in g['members'])),
                    # 実印刷枚数（出庫枚数が入っている行の通し枚数）
                    'prn':     int(sum(mm['tsu'] or 0 for mm in g['members']
                                       if col_of(mm, OUT_COL))),
                    # この管理番号に出庫枚数があるか（部署をまたいで見る）。
                    # 損紙率で数える範囲を、機械別の表とそろえるための目印
                    'ob':      1 if g['base'] in month_out else 0,
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
        # 並びは機械・日付・項目で決め打ちにする。読んだ順のままだと、
        # 前回のまま使う年があるかどうかで中身の並びが変わってしまう。
        'daily':     [dict(d, m=m) for m in sorted(daily_by_month)
                      for d in sorted(daily_by_month[m],
                                      key=lambda x: (str(x.get('machine') or ''),
                                                     str(x.get('date') or ''),
                                                     str(x.get('label') or '')))],
        'oper':      build_oper(year, daily_by_month, closed),
        'waste':     build_waste(year, by_month),
        'records':   records,
    }


CACHE_NAME = 'cache.json'
CACHE_VERSION = 4     # 目印の作り方を変えたら上げる


def load_cache(outdir):
    """前回の読み込みの控えを読む

    中身:
      files … {パス: {"key": [大きさ, 更新日時], "info": 元ファイルの一覧に出す情報}}
      years … {"2013": {"summary": 年の要約, "files": [パス, ...]}}
      cols  … 日報の見出し列（出てきた順）
    """
    p = Path(outdir) / CACHE_NAME
    try:
        c = json.loads(p.read_text(encoding='utf-8'))
        if c.get('version') == CACHE_VERSION:
            return c
    except (OSError, ValueError):
        pass
    return {'version': CACHE_VERSION, 'files': {}, 'years': {}, 'cols': []}


def save_cache(outdir, cache):
    cache['version'] = CACHE_VERSION
    try:
        (Path(outdir) / CACHE_NAME).write_text(
            json.dumps(cache, ensure_ascii=False), encoding='utf-8')
    except OSError as e:                                         # noqa: BLE001
        print('  控えを書けませんでした（次回も全部読み直します）: %s' % e)


def merge_read(a, b):
    """read_folder の結果を1つにまとめる"""
    rows, daily, cols, info, warns = a
    r2, d2, c2, i2, w2 = b
    for y, months in r2.items():
        for m, rr in months.items():
            rows.setdefault(y, OrderedDict()).setdefault(m, []).extend(rr)
    for y, months in d2.items():
        for m, dd in months.items():
            daily.setdefault(y, OrderedDict()).setdefault(m, []).extend(dd)
    for n in c2:
        if n not in cols:
            cols.append(n)
    return rows, daily, cols, info + i2, warns + w2


def read_incremental(targets, cache, outdir, force_year=None):
    """変わったところだけを読む

    去年より前の日報は内容が決まっているので、ファイルが変わっていなければ
    読み直さない。書き変わるのは当月の日報だけなので、2回目からは今年のぶん
    だけを読むことになる。

    新しく増えたファイルは、開いてみるまでどの年のものか分からない。そこで
      1回目 … 変わった・増えたファイルだけ読んで、関わる年を知る
      2回目 … その年に関わる残りのファイルを読む（年の全行がそろわないと
               集計できないため）
    の2段階にしてある。

    戻り値: (読み取り結果, 作り直す年, そのまま使う年, 消えたファイル, 読んだ数)
    """
    now = {str(p): list(file_key(p)) for p, _k in targets}
    old = cache.get('files') or {}
    cyears = cache.get('years') or {}

    changed = {p for p, k in now.items()
               if p not in old or list(old[p].get('key') or []) != k}
    gone = [p for p in old if p not in now]

    # 変わった・消えたファイルが、控えの上で関わっていた年
    dirty = set()
    for p in list(changed) + gone:
        for y in (old.get(p, {}).get('years') or []):
            dirty.add(int(y))
    # 出来上がりが残っていない年
    for ys in cyears:
        if not (Path(outdir) / ('data_%s.json' % ys)).exists():
            dirty.add(int(ys))
    if force_year:
        dirty.add(int(force_year))

    # ① 変わった・増えたファイルを読む（増えたぶんの年はここで分かる）
    done = set()
    got = read_folder([(p, k) for p, k in targets if str(p) in changed])
    done |= changed
    for y in got[0]:
        dirty.add(int(y))

    # ② 作り直す年に関わる残りのファイルを読む（年の全行がそろわないと集計できない）
    need = set()
    for ys, e in cyears.items():
        if int(ys) in dirty:
            need.update(e.get('files') or [])
    need -= done
    need &= set(now)
    if need:
        got = merge_read(got, read_folder([(p, k) for p, k in targets
                                           if str(p) in need]))
        done |= need

    keep = sorted(int(y) for y in cyears if int(y) not in dirty)
    return got, dirty, keep, gone, done


def build(src, year=None, outdir=None, closed=None, fresh=False):
    """フォルダを読み、年ごとの data_YYYY.json と years.json を書き出す

    去年より前の日報は内容が決まっているので、ファイルが変わっていなければ
    作り直さない（前回の出来上がりをそのまま使う）。書き変わるのは当月の
    日報だけなので、2回目からは今年のぶんだけを読むことになる。
    fresh=True で控えを無視して全部作り直す。
    """
    srcs, cands = resolve_src(src)
    if not srcs:
        msg = 'フォルダが見つかりません:\n    ' + '\n    '.join(cands)
        hint = hint_subdirs(cands)
        if hint:
            msg += '\n  ' + hint
        raise SystemExit(msg)

    outdir = Path(outdir or WEB)
    outdir.mkdir(parents=True, exist_ok=True)
    targets = collect_files(srcs)
    if not targets:
        raise SystemExit('稼動日報の .xls が1つも見つかりません:\n    '
                         + '\n    '.join(str(x) for x in srcs))

    cache = {'version': CACHE_VERSION, 'files': {}, 'years': {}, 'cols': []} \
        if (fresh or year) else load_cache(outdir)
    got, dirty, keep, gone, done = read_incremental(targets, cache, outdir, year)
    if keep:
        print('  前回のまま使う年: %s（%d件のファイルは読みません）'
              % ('、'.join('%d年' % y for y in keep), len(targets) - len(done)))
    rows, daily, cols, read_info, warnings = got
    for n in (cache.get('cols') or []):          # 読まなかったファイルの列も残す
        if n not in cols:
            cols.append(n)

    # 見つからなかったフォルダは警告にしない。U: が使えるPCではUNC側が、
    # 使えないPCではU:側が必ず存在しないため、毎回出ると邪魔になる。
    # 「元ファイル」タブでだけ確認できるようにする。
    skipped = [c for c in cands if not Path(c).is_dir()]

    # 今回読んだぶんで控えを更新する
    now_keys = {str(p): list(file_key(p)) for p, _k in targets}
    cfiles = dict(cache.get('files') or {})
    for p in gone:
        cfiles.pop(p, None)
    read_by_path = {f['path']: f for f in read_info}
    for path in done:
        f = read_by_path.get(path)
        cfiles[path] = {'key': now_keys[path],
                        'years': sorted(int(y) for y in (f['years'] if f else {})),
                        'info': f}
    # 読まなかったファイルの情報は控えから引き継ぐ
    files = [cfiles[p]['info'] for p in sorted(cfiles) if cfiles[p].get('info')]

    read_years = sorted(y for y in rows if y >= MIN_YEAR)
    too_old = sorted(y for y in rows if y < MIN_YEAR)
    if too_old:
        warnings.append('%s は %d年より前のため読み込みませんでした。'
                        % ('、'.join('%d年' % y for y in too_old), MIN_YEAR))
    keep = [y for y in keep if y not in rows]   # 行を読んだ年は必ず作り直す
    years = sorted(set(read_years) | set(keep))
    if not years:
        raise SystemExit('%d年以降のデータが見つかりませんでした。' % MIN_YEAR)
    if year:
        if year not in read_years:
            raise SystemExit('%d年 のデータがありません。このフォルダにあるのは %s です。'
                             % (year, '、'.join('%d年' % y for y in read_years)))
        years, keep = [year], []

    source = ' ／ '.join(str(x) for x in srcs)
    cyears = dict(cache.get('years') or {})
    summary = []
    for y in years:
        if y in keep:                            # 変わっていないので作り直さない
            summary.append(cyears[str(y)]['summary'])
            continue
        data = build_year(y, rows[y], daily.get(y, {}), cols, files, source, warnings,
                          closed)
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
        one = {
            'year':    y,
            'months':  data['months'],
            'records': len(data['records']),
            'detail':  sum(s['detail'] for s in data['stats']),
            'tsu':     sum(s['tsuGroups'] for s in data['stats']),
            'depts':   list(data['depts']),
            'byDept':  by_dept,          # {部署: {"月": [通し数, 件数]}}
        }
        summary.append(one)
        cyears[str(y)] = {'summary': one,
                          'files': sorted(p for p, e in cfiles.items()
                                          if y in (e.get('years') or []))}

    if not year:
        # 年を絞らずに作り直したときは、今回の対象でない年のファイルを消す。
        # 前回の読み込みで出ていた年が、画面に残り続けないようにするため。
        for p in outdir.glob('data_*.json'):
            try:
                if int(p.stem.split('_')[1]) not in years:
                    p.unlink()
            except (ValueError, IndexError, OSError):
                pass
        for ys in [y for y in cyears if int(y) not in years]:
            cyears.pop(ys, None)
        save_cache(outdir, {'files': cfiles, 'years': cyears, 'cols': cols})

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
    p.add_argument('--closed', action='append',
                   help='会社の休業日（--closed 08-13..15 のように複数指定可）')
    p.add_argument('--fresh', action='store_true',
                   help='前回の控えを使わず、全部読み直す')
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
    m = build(a.src, a.year, a.out, a.closed, a.fresh)

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
