#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本番ダンプ（koutei リポジトリ web_app/data/pg_backup.sql）を読み込んだテスト DB で、
Python 版と SQL 版を全受注番号について走らせ、結果を突き合わせる。

    python tests/compare_on_dump.py --dsn "host=... port=... dbname=koutei_dump_test user=..."

読み取りのみ（ダンプの読み込みは別途 psql -f で済ませておく）。本番には向けない。
"""
import argparse
import collections
import csv
import os
import subprocess
import sys
import tempfile

import psycopg2

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "..", "export_finish_size.py")
SQL = os.path.join(HERE, "..", "query_finish_size.sql")


def sets(row):
    return tuple(frozenset(x.split(" / ")) if x else frozenset() for x in row)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dsn", required=True)
    args = ap.parse_args()
    params = dict(p.split("=", 1) for p in args.dsn.split())
    if "test" not in params["dbname"]:
        sys.exit("テスト用 DB（名前に test を含む）だけに向けてください")

    conn = psycopg2.connect(args.dsn, options="-c default_transaction_read_only=on")
    conn.set_session(readonly=True, autocommit=True)
    cur = conn.cursor()
    cur.execute("SELECT to_regclass('outsourcing_list')")
    if cur.fetchone()[0] is None:
        # ダンプに無いテーブル。SQL 版が参照するので空で作る（テスト DB なので可）
        w = psycopg2.connect(args.dsn)
        w.autocommit = True
        w.cursor().execute("CREATE TABLE outsourcing_list (id SERIAL, order_number TEXT, finish_size TEXT, outsourcing_company TEXT, processing_content TEXT)")
        w.close()
        print("outsourcing_list が無いので空テーブルを作りました（テスト DB）")
    norm = "regexp_replace(COALESCE(substring(regexp_replace(COALESCE({c}, ''), '\\s', '', 'g') from '^\\d+'), ''), '^0+', '')"
    cur.execute(f"""
        SELECT DISTINCT k FROM (
            SELECT {norm.format(c='"orderNumber"')} k FROM timeline_processes
            UNION SELECT {norm.format(c='"orderNumber"')} FROM waiting_list
            UNION SELECT {norm.format(c='order_number')} FROM delivery_schedule
            UNION SELECT {norm.format(c='order_number')} FROM processing_works
        ) x WHERE k <> '' ORDER BY k
    """)
    keys = [r[0] for r in cur.fetchall()]
    conn.close()
    print(f"ダンプ内の受注番号: {len(keys)} 件")

    env = dict(os.environ)
    env.update({"DATABASE_HOST": params.get("host", "localhost"), "DATABASE_PORT": params.get("port", "5432"),
                "DATABASE_NAME": params["dbname"], "DATABASE_USER": params.get("user", ""),
                "DATABASE_PASSWORD": params.get("password", ""), "PGPASSWORD": params.get("password", "")})
    with tempfile.TemporaryDirectory() as td:
        orders = os.path.join(td, "orders.csv")
        with open(orders, "w", encoding="utf-8", newline="") as f:
            f.write("受注番号\n" + "\n".join(keys) + "\n")
        out = os.path.join(td, "py.csv")
        r = subprocess.run([sys.executable, SCRIPT, "--orders", orders, "--out", out], env=env, capture_output=True, text=True)
        print(r.stdout[-600:])
        if r.returncode:
            print(r.stderr)
            sys.exit(1)
        py = {row[0]: row[1:] for row in list(csv.reader(open(out, encoding="utf-8-sig")))[1:]}

        sql_out = os.path.join(td, "sql.csv")
        cmd = ["psql", "-h", params.get("host", "localhost"), "-p", params.get("port", "5432"), "-U", params.get("user", ""),
               "-d", params["dbname"], "-v", "ON_ERROR_STOP=1", "--csv", "-q",
               "-c", "CREATE TEMP TABLE target(order_no text)",
               "-c", f"\\copy target FROM '{orders}' WITH (FORMAT csv, HEADER true)",
               "-f", SQL, "-o", sql_out]
        r = subprocess.run(cmd, env=env, capture_output=True, text=True)
        if r.returncode:
            print("SQL 版がエラー:", r.stderr)
            sys.exit(1)
        sq = {row[0]: row[1:] for row in list(csv.reader(open(sql_out, encoding="utf-8")))[1:]}

    diff = 0
    for k in keys:
        if sets(py.get(k, [""] * 7)) != sets(sq.get(k, [""] * 7)):
            diff += 1
            if diff <= 15:
                print(f"差: {k}\n  py : {py.get(k)}\n  sql: {sq.get(k)}")
    stat = collections.Counter()
    for k in keys:
        row = py.get(k, [""] * 7)
        stat["仕上りサイズあり"] += bool(row[0])
        stat["加工内容あり"] += bool(row[1])
        stat["内作"] += row[2] == "内作"
        stat["外注"] += row[2] == "外注"
        stat["内作・外注"] += row[2] == "内作・外注"
        stat["委託先あり"] += bool(row[3])
        stat["用紙銘柄あり"] += bool(row[4])
        stat["B1だけ"] += row[0] == "B1"
    print("Python 版の内訳:", dict(stat))
    print(f"Python 版と SQL 版の差: {diff} / {len(keys)} 件")
    sys.exit(1 if diff else 0)


if __name__ == "__main__":
    main()
