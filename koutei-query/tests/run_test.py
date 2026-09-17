#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
export_finish_size.py の動作確認。

テスト用の空 PostgreSQL に fixture.sql を流し、受注番号一覧（ゼロ埋め・小数・全角のゆれ入り）で
スクリプトを走らせて、出力 CSV を突き合わせる。

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

EXPECT = {
    # 受注番号(入力表記): (仕上りサイズ, 加工内容, 内外作区分, 委託先名)
    "08726258": ("A4 297×210", "中綴じ12P / ミシン(筋)", "内作・外注", "八王子紙工"),
    "8726259": ("B2 728×515", "", "", ""),
    "8726260.0": ("A4", "抜き・ポケット貼り・24P中綴じ", "外注", "松岡製本"),
    "8726261": ("", "", "", ""),
    "８７２６２６２": ("B5", "折加工（二つ折り） / 二つ折り / 折加工1", "内作", ""),
    "8726263": ("A3", "二つ折り", "外注", "松岡製本"),
    "8726264": ("", "", "", ""),
}


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
    })

    with tempfile.TemporaryDirectory() as td:
        orders = os.path.join(td, "orders.csv")
        with open(orders, "w", encoding="cp932", newline="") as f:
            f.write("受注番号（管理番号）,品名\r\n")
            for k in EXPECT:
                f.write(f"{k},x\r\n")
        out = os.path.join(td, "out.csv")
        r = subprocess.run([sys.executable, SCRIPT, "--orders", orders, "--out", out],
                           env=env, capture_output=True, text=True)
        print(r.stdout)
        if r.returncode != 0:
            print(r.stderr)
            sys.exit("スクリプトが失敗しました")

        rows = list(csv.reader(open(out, encoding="utf-8-sig")))
        assert rows[0] == ["受注番号", "仕上りサイズ", "加工内容", "内外作区分", "委託先名"], rows[0]
        got = {r[0]: tuple(r[1:]) for r in rows[1:]}
        ok = True
        for k, exp in EXPECT.items():
            if got.get(k) != exp:
                ok = False
                print(f"NG {k}: 期待 {exp} / 実際 {got.get(k)}")
            else:
                print(f"OK {k}: {exp}")
        assert list(got.keys()) == list(EXPECT.keys()), "入力の並びが保たれていない"

        detail = list(csv.reader(open(os.path.join(td, "out_詳細.csv"), encoding="utf-8-sig")))
        tables = sorted({r[2] for r in detail[1:]})
        print("詳細CSV:", len(detail) - 1, "行", tables)
        assert not any(r[1] == "08799999" for r in detail[1:]), "一覧に無い受注番号を拾っている"
        assert len(detail) - 1 == 10, f"詳細行数 {len(detail) - 1}（期待 10）"

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

    print("\nALL OK" if ok else "\nFAILED")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
