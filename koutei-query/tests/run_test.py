#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
export_finish_size.py の動作確認。

テスト用の空 PostgreSQL に fixture.sql を流し、受注番号一覧（ゼロ埋め・小数・全角のゆれ入り）で
スクリプトを走らせて、出力 CSV を突き合わせる。SQL 版（query_finish_size.sql）も同じ結果になることを見る。

    python tests/run_test.py --dsn "host=... port=... dbname=... user=..."

※本番 DB に向けないこと（fixture.sql はテーブルを DROP する）。
"""
import argparse
import csv
import os
import subprocess
import sys
import tempfile

import psycopg2

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "..", "export_finish_size.py")
SQL = os.path.join(HERE, "..", "query_finish_size.sql")
HEADER = ["受注番号", "仕上りサイズ", "A3以下", "加工内容", "内外作区分", "委託先名", "用紙銘柄", "用紙規格", "斤量"]
SQL_HEADER = [h for h in HEADER if h != "A3以下"]  # SQL 版に A3 判定は無い

EXPECT = {
    # 受注番号(入力表記): (仕上りサイズ, A3以下, 加工内容, 内外作区分, 委託先名, 用紙銘柄, 用紙規格, 斤量)
    "08726258": ("A4 297×210", "○", "中綴じ12P / ミシン(筋)", "内作・外注", "八王子紙工", "A2マット", "A全判", "86.5 / 110"),
    "8726259": ("B2 728×515", "×", "折り", "", "", "オーロラコート", "菊全判", "93.5"),
    "8726260.0": ("A4", "○", "抜き・ポケット貼り・24P中綴じ", "外注", "松岡製本", "", "", ""),
    "8726261": ("", "不明", "", "", "", "", "", ""),
    "８７２６２６２": ("B5", "○", "折加工（二つ折り） / 二つ折り / 折加工1", "内作", "", "", "", ""),
    "8726263": ("A3", "○", "二つ折り", "外注", "松岡製本", "", "", ""),
    "8726264": ("", "不明", "", "", "", "", "", ""),
    "8726265": ("A5", "○", "", "", "", "", "", ""),
}
DETAIL_ROWS = 12  # tp-1..7 + wl-1 + ol-1 + ds-1 + pw-1,2


# SQL 版（query_finish_size.sql）に未反映の規則。Python 版が正で、SQL 版は検算用
KNOWN_SQL_GAPS = {"その他"}  # MIS の仕上加工名「その他」を加工内容に出さない


def as_sets(t, drop=()):
    """「 / 」区切りの並び順の違いを無視して比べる。"""
    return tuple(frozenset(x.split(" / ")) - set(drop) if x else frozenset() for x in t)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dsn", required=True)
    args = ap.parse_args()

    conn = psycopg2.connect(args.dsn)
    conn.autocommit = True
    cur = conn.cursor()
    cur.execute("SELECT current_database()")
    dbname = cur.fetchone()[0]
    if "test" not in dbname and "fixture" not in dbname and os.environ.get("KOUTEI_TEST_ALLOW_DB") != dbname:
        # 本番っぽい名前のDBに誤って流さないための安全弁
        sys.exit(f"DB 名 '{dbname}' はテスト用に見えません。KOUTEI_TEST_ALLOW_DB={dbname} を付けて明示してください。")
    cur.execute(open(os.path.join(HERE, "fixture.sql"), encoding="utf-8").read())
    conn.close()

    params = dict(p.split("=", 1) for p in args.dsn.split())
    env = dict(os.environ)
    env.update({
        "DATABASE_HOST": params.get("host", "localhost"),
        "DATABASE_PORT": params.get("port", "5432"),
        "DATABASE_NAME": params["dbname"],
        "DATABASE_USER": params.get("user", ""),
        "DATABASE_PASSWORD": params.get("password", ""),
        "PGPASSWORD": params.get("password", ""),
    })
    ok = True

    with tempfile.TemporaryDirectory() as td:
        orders = os.path.join(td, "orders.csv")
        with open(orders, "w", encoding="cp932", newline="") as f:
            f.write("受注番号（管理番号）,品名\r\n")
            for k in EXPECT:
                f.write(f"{k},x\r\n")
            f.write("0,受注番号として読めない行\r\n")
        out = os.path.join(td, "out.csv")
        r = subprocess.run([sys.executable, SCRIPT, "--orders", orders, "--out", out],
                           env=env, capture_output=True, text=True)
        print(r.stdout)
        if r.returncode != 0:
            print(r.stderr)
            sys.exit("スクリプトが失敗しました")
        assert "読めなかった値 1 件" in r.stdout, "「0」を飛ばした警告が出ていない"

        rows = list(csv.reader(open(out, encoding="utf-8-sig")))
        assert rows[0] == HEADER, rows[0]
        got = {r[0]: tuple(r[1:]) for r in rows[1:]}
        for k, exp in EXPECT.items():
            if got.get(k) != exp:
                ok = False
                print(f"NG {k}: 期待 {exp} / 実際 {got.get(k)}")
            else:
                print(f"OK {k}: {exp}")
        assert list(got.keys()) == list(EXPECT.keys()), "入力の並びが保たれていない"

        a3 = list(csv.reader(open(os.path.join(td, "out_A3以下.csv"), encoding="utf-8-sig")))
        assert a3[0] == HEADER and [r[0] for r in a3[1:]] == [k for k, e in EXPECT.items() if e[1] == "○"], a3
        assert "A3 以下の判定: ○ 5 件、× 1 件、不明 2 件" in r.stdout, r.stdout

        detail = list(csv.reader(open(os.path.join(td, "out_詳細.csv"), encoding="utf-8-sig")))
        tables = sorted({r[3] for r in detail[1:]})
        print("詳細CSV:", len(detail) - 1, "行", tables)
        assert not any(r[1] == "08799999" for r in detail[1:]), "一覧に無い受注番号を拾っている"
        assert len(detail) - 1 == DETAIL_ROWS, f"詳細行数 {len(detail) - 1}（期待 {DETAIL_ROWS}）"
        assert any(r[1] == "08726258（表紙）" and r[2] == "（表紙）" and r[0] == "08726258" for r in detail[1:]), "枝付きの受注番号が詳細に出ていない"

        # 読み取り専用の確認: 同じ接続方法で書き込みが拒否されること
        conn = psycopg2.connect(args.dsn, options="-c default_transaction_read_only=on")
        conn.set_session(readonly=True, autocommit=True)
        try:
            conn.cursor().execute("DELETE FROM timeline_processes")
            ok = False
            print("NG: 読み取り専用のはずが DELETE が通った")
        except psycopg2.errors.ReadOnlySqlTransaction:
            print("OK: 読み取り専用接続で DELETE は拒否された")
        conn.close()

        # SQL 版が同じ結果になること（並び順の違いは無視）
        sql_orders = os.path.join(td, "orders_sql.csv")
        with open(sql_orders, "w", encoding="utf-8", newline="") as f:
            f.write("受注番号\n" + "\n".join(EXPECT) + "\n")
        sql_out = os.path.join(td, "sql.csv")
        cmd = ["psql", "-h", params.get("host", "localhost"), "-p", params.get("port", "5432"),
               "-U", params.get("user", ""), "-d", params["dbname"], "-v", "ON_ERROR_STOP=1", "--csv", "-q",
               "-c", "CREATE TEMP TABLE target(order_no text)",
               "-c", f"\\copy target FROM '{sql_orders}' WITH (FORMAT csv, HEADER true)",
               "-f", SQL, "-o", sql_out]
        r = subprocess.run(cmd, env=env, capture_output=True, text=True)
        if r.returncode != 0:
            ok = False
            print("NG: SQL 版がエラー\n", r.stderr)
        else:
            srows = list(csv.reader(open(sql_out, encoding="utf-8")))
            assert srows[0] == SQL_HEADER, srows[0]
            # SQL 版は入力の表記をそのまま返すので、正規化キーで突き合わせる
            sys.path.insert(0, os.path.join(HERE, ".."))
            from export_finish_size import norm_order
            sgot = {norm_order(r[0])[0]: tuple(r[1:]) for r in srows[1:]}
            for k, exp in EXPECT.items():
                sk = norm_order(k)[0]
                exp = exp[:1] + exp[2:]  # A3 判定列を除く
                if as_sets(sgot.get(sk, ()), KNOWN_SQL_GAPS) != as_sets(exp):
                    ok = False
                    print(f"NG SQL版 {k}: 期待 {exp} / 実際 {sgot.get(sk)}")
            print("SQL版: Python 版と一致（既知の差「その他」を除く）" if ok else "SQL版: 不一致あり")

    # 入力の読み方（見出し判定・小数・枝）
    sys.path.insert(0, os.path.join(HERE, ".."))
    from export_finish_size import read_order_numbers, norm_order
    with tempfile.TemporaryDirectory() as td:
        p = os.path.join(td, "o.csv")
        open(p, "w", encoding="utf-8").write("2026年9月 未確認一覧\n08726258\n8726274.5\n0\n")
        got_keys = list(read_order_numbers(p).keys())
        assert got_keys == ["8726258"], got_keys
    assert norm_order("08704276②") == ("8704276", "②")
    assert norm_order("087033142②") == ("87033142", "②")
    assert norm_order("８７２６２６２.0") == ("8726262", "")
    assert norm_order("8726274.5")[0] == ""
    print("OK: 見出し判定・小数・枝の扱い")

    # events[] に別の受注番号の要素が混ざっていても拾わない
    from export_finish_size import new_record, absorb_row_json
    rec = new_record("timeline_processes", "x", "08726259")
    absorb_row_json(rec, {"finish_size": "B2", "events": [
        {"order_number": "08726259", "finish_process": "折り"},
        {"order_number": "08799999", "finish_size": "OTHER", "finish_process": "中綴じ", "outsource_name": "他社"},
        {"finish_process": "その他"}, "junk", None]})
    assert rec["sizes"] == ["B2"] and rec["contents"] == ["折り"] and rec["companies"] == [], rec
    assert rec["flags"] == ["events:08799999"], rec["flags"]
    print("OK: 別受注番号の events 要素は飛ばす")

    # A3 以下の判定
    from export_finish_size import judge_a3, judge_a3_all
    cases = {"A4": "○", "A4 297×210": "○", "A3": "○", "A3ノビ": "×", "B4": "○", "B3": "×", "B5以下": "○", "B1": "×",
             "B2 728×515": "×", "規格外": "?", "153×75": "○", "153*75": "○", "220*520": "×", "297*420": "○", "300×423": "○",
             "ハガキ": "○", "長3": "○", "A全判": "×", "菊全": "×", "A4変形": "○", "B5 182×257": "○", "A2": "×",
             "29.7×21cm": "○", "名刺 91×55": "○", "A3二つ折り": "○", "348*527": "×", "": "?"}
    bad = {t: judge_a3(t) for t, e in cases.items() if judge_a3(t) != e}
    assert not bad, bad
    assert judge_a3_all(["A4", "B2 728×515"]) == "×" and judge_a3_all(["規格外"]) == "不明"
    assert judge_a3_all([]) == "不明" and judge_a3_all(["A4", "規格外"]) == "○"
    print("OK: A3 以下の判定")

    print("\nALL OK" if ok else "\nFAILED")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
