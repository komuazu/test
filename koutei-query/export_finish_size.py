#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
koutei-kanr30（社内工程管理システム）の本番DBから、受注番号ごとの
「仕上りサイズ・加工内容・内外作区分・委託先名」を読み出して CSV にする。

* 読み取り専用。DB もコードも一切書き換えない
  （接続を default_transaction_read_only=on にして開くので、書き込み文は DB 側で拒否される）
* 本番機（丸本PC）の PostgreSQL は localhost:5432 にしか口を開けていないので、
  このスクリプトは丸本PC の上で動かす（詳しくは README.md）

使い方（丸本PC、コマンドプロンプト）:

    cd <このフォルダ>
    python export_finish_size.py --orders 仕上りサイズ未確認_受注番号一覧.csv ^
        --env C:\\Users\\116544\\Desktop\\UPDATE17\\web_app\\.env

出力:
    仕上りサイズ_koutei取得結果.csv        … 受注番号, 仕上りサイズ, 加工内容, 内外作区分, 委託先名, 用紙銘柄, 用紙規格, 斤量
    仕上りサイズ_koutei取得結果_詳細.csv   … どのテーブルのどの行から取ったか（検算用）
"""

import argparse
import csv
import getpass
import io
import json
import os
import re
import sys
import unicodedata
from collections import OrderedDict

try:
    import psycopg2
    import psycopg2.extras
except ImportError:  # pragma: no cover
    sys.exit("psycopg2 が入っていません。 pip install psycopg2-binary で入れてください。")


# ------------------------------------------------------------
# 受注番号のゆれ吸収
# ------------------------------------------------------------
def norm_order(value):
    """'08726258' / '8726258' / '8726258.0' / '８７２６２５８' を同じキーにする。"""
    if value is None:
        return ""
    s = unicodedata.normalize("NFKC", str(value)).strip()
    s = re.sub(r"\s+", "", s)
    if re.fullmatch(r"\d+\.0+", s):  # Excel 由来の 8726258.0
        s = s.split(".")[0]
    s = s.lstrip("0")
    return s


# SQL 側で同じ正規化をする式（"orderNumber" 列 / order_number 列に当てる）
SQL_NORM = "regexp_replace(regexp_replace(COALESCE({col}, ''), '\\s', '', 'g'), '^0+', '')"


def truthy(v):
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    return str(v).strip().lower() in ("1", "true", "yes", "on", "内作", "外注")


def clean(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("none", "null", "nan", "未設定", "-", "未定") else s


def add_unique(lst, value):
    value = clean(value)
    if value and value not in lst:
        lst.append(value)


def add_paper(rec, paper_type, standard_size, paper_weight):
    t = (clean(paper_type), clean(standard_size), clean(paper_weight))
    if any(t) and t not in rec["papers"]:
        rec["papers"].append(t)


# ------------------------------------------------------------
# 入力 CSV（受注番号一覧）
# ------------------------------------------------------------
def read_order_numbers(path):
    raw = open(path, "rb").read()
    text = None
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        sys.exit(f"受注番号一覧 {path} の文字コードが読めません（UTF-8 か CP932 にしてください）")

    rows = list(csv.reader(io.StringIO(text)))
    rows = [r for r in rows if any(c.strip() for c in r)]
    if not rows:
        sys.exit("受注番号一覧が空です")

    col = 0
    start = 0
    header = [unicodedata.normalize("NFKC", c).strip() for c in rows[0]]
    for i, h in enumerate(header):
        if "受注番号" in h or "管理番号" in h:
            col = i
            start = 1
            break
    else:
        if not re.fullmatch(r"\d+(\.0+)?", norm_order(header[0]) or "x"):
            start = 1  # 数字でない1行目は見出しとみなす

    result = OrderedDict()  # 正規化キー → 元の表記（最初に出たもの）
    for r in rows[start:]:
        if col >= len(r):
            continue
        key = norm_order(r[col])
        if key and key not in result:
            result[key] = r[col].strip()
    return result


# ------------------------------------------------------------
# DB 接続（読み取り専用）
# ------------------------------------------------------------
def load_env_file(path):
    """python-dotenv が無くても動くように、KEY=VALUE を素朴に読む。"""
    env = {}
    if not path or not os.path.exists(path):
        return env
    with open(path, encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip().strip('"').strip("'")
    return env


def find_env_file(explicit):
    if explicit:
        return explicit
    here = os.path.dirname(os.path.abspath(__file__))
    for cand in (
        os.path.join(os.getcwd(), ".env"),
        os.path.join(os.getcwd(), "web_app", ".env"),
        os.path.join(here, ".env"),
        os.path.join(here, "..", "web_app", ".env"),
        os.path.join(os.getcwd(), "web_app", ".env.marumoto"),
    ):
        if os.path.exists(cand):
            return cand
    return None


def connect(args):
    env = load_env_file(find_env_file(args.env))
    get = lambda k, d=None: args.__dict__.get(k.lower()) or os.environ.get(k) or env.get(k) or d
    host = get("DATABASE_HOST", "localhost")
    port = int(get("DATABASE_PORT", 5432))
    dbname = get("DATABASE_NAME", "koutei_kanri_dev")
    user = get("DATABASE_USER", "koutei_dev")
    password = None
    for k, src in (("DATABASE_PASSWORD", os.environ), ("PGPASSWORD", os.environ), ("DATABASE_PASSWORD", env)):
        if src.get(k) is not None:
            password = src[k]
            break
    if password is None:
        password = getpass.getpass(f"DB パスワード（{user}@{host}:{port}/{dbname}）: ")

    print(f"接続先: {dbname}@{host}:{port} ユーザー={user}（読み取り専用）")
    conn = psycopg2.connect(
        host=host,
        port=port,
        dbname=dbname,
        user=user,
        password=password,
        options="-c default_transaction_read_only=on",
        connect_timeout=10,
    )
    conn.set_session(readonly=True, autocommit=True)
    return conn


def existing_columns(cur, table):
    cur.execute(
        """
        SELECT column_name FROM information_schema.columns
        WHERE table_schema = 'public' AND table_name = %s
        """,
        (table,),
    )
    return {r[0] for r in cur.fetchall()}


# ------------------------------------------------------------
# 各テーブルからの読み出し
# ------------------------------------------------------------
def parse_json(v):
    if v is None:
        return {}
    if isinstance(v, dict):
        return v
    try:
        d = json.loads(v)
        return d if isinstance(d, dict) else {}
    except (TypeError, ValueError):
        return {}


def rows_from_event_table(cur, table, keys):
    """timeline_processes / waiting_list: 明細は data 列の JSON に入っている。"""
    cols = existing_columns(cur, table)
    if not cols or "data" not in cols:
        print(f"  {table}: テーブルまたは data 列が無いので飛ばします")
        return []
    ocol = '"orderNumber"' if "orderNumber" in cols else ("order_number" if "order_number" in cols else None)
    if ocol is None:
        print(f"  {table}: 受注番号の列が無いので飛ばします")
        return []
    extra = [c for c in ("date", "machine", "name", "updated_at") if c in cols]
    sel = ", ".join(["id", ocol, "data"] + extra)
    cur.execute(
        f"SELECT {sel} FROM {table} WHERE {SQL_NORM.format(col=ocol)} = ANY(%s)",
        (keys,),
    )
    out = []
    for row in cur.fetchall():
        rid, onum, data = row[0], row[1], row[2]
        d = parse_json(data)
        rec = {
            "table": table,
            "row_id": rid,
            "order_raw": onum,
            "key": norm_order(onum),
            "part": clean(d.get("part_type") or d.get("part_name") or d.get("partName")),
            "finish_size": clean(d.get("finish_size") or d.get("finished_size") or d.get("finishSize")),
            "papers": [],  # (銘柄, 規格, 斤量)
            "contents": [],
            "internal": truthy(d.get("internal_work")),
            "outsourcing": truthy(d.get("outsourcing")),
            "work_department": clean(d.get("work_department")),
            "companies": [],
            "company_contents": [],  # (委託先, 加工内容)
            "flags": [],
            "updated_at": row[3 + extra.index("updated_at")] if "updated_at" in extra else "",
        }
        for k in ("finish_processing_name", "finish_processing", "finishProcessing", "processing_content", "finish_process"):
            add_unique(rec["contents"], d.get(k))
        add_paper(rec, d.get("paper_type") or d.get("paperType"), d.get("standard_size"), d.get("paper_weight") or d.get("paperWeight"))
        for lst_key in ("order_paper_info", "papers"):  # 部品ごとの用紙（外注委託依頼書で使う形）
            lst = d.get(lst_key)
            if isinstance(lst, list):
                for pi in lst:
                    if isinstance(pi, dict):
                        add_paper(rec, pi.get("paper_type"), pi.get("standard_size"), pi.get("paper_weight"))
        for k in ("outsourcing_company", "outsource_name"):
            add_unique(rec["companies"], d.get(k))
        items = d.get("outsourcing_items")
        if isinstance(items, list):
            for it in items:
                if not isinstance(it, dict):
                    continue
                comp = clean(it.get("company") or it.get("company_name") or it.get("name"))
                cont = clean(it.get("processing_content") or it.get("content"))
                add_unique(rec["companies"], comp)
                add_unique(rec["contents"], cont)
                if comp or cont:
                    pair = (comp, cont)
                    if pair not in rec["company_contents"]:
                        rec["company_contents"].append(pair)
        for k in ("isReturnProcess", "isContinueProcess", "is_return", "is_continue"):
            if truthy(d.get(k)):
                rec["flags"].append(k)
        out.append(rec)
    print(f"  {table}: {len(out)} 行")
    return out


def rows_from_flat_table(cur, table, keys, colmap):
    """outsourcing_list / delivery_schedule / processing_works: 列に直接入っている。"""
    cols = existing_columns(cur, table)
    if not cols:
        print(f"  {table}: テーブルが無いので飛ばします")
        return []
    ocol = "order_number" if "order_number" in cols else ('"orderNumber"' if "orderNumber" in cols else None)
    if ocol is None:
        print(f"  {table}: 受注番号の列が無いので飛ばします")
        return []
    wanted = [c for c in colmap if c in cols]
    sel = ", ".join(["id", ocol] + wanted)
    cur.execute(
        f"SELECT {sel} FROM {table} WHERE {SQL_NORM.format(col=ocol)} = ANY(%s)",
        (keys,),
    )
    out = []
    for row in cur.fetchall():
        d = dict(zip(wanted, row[2:]))
        rec = {
            "table": table,
            "row_id": row[0],
            "order_raw": row[1],
            "key": norm_order(row[1]),
            "part": clean(d.get("part_name")),
            "finish_size": clean(d.get("finish_size")),
            "papers": [],
            "contents": [],
            "internal": False,
            "outsourcing": truthy(d.get("outsourcing")),
            "work_department": clean(d.get("work_department")),
            "companies": [],
            "company_contents": [],
            "flags": [],
            "updated_at": d.get("updated_at") or "",
        }
        add_unique(rec["contents"], d.get("processing_content"))
        add_unique(rec["companies"], d.get("outsourcing_company"))
        add_paper(rec, d.get("paper_type"), d.get("standard_size"), d.get("paper_weight"))
        if table == "processing_works":
            # 内作加工の作業一覧。classification は 折加工1／折加工2／中綴じ加工 などの機械区分
            rec["internal"] = True
            add_unique(rec["contents"], d.get("classification"))
        if table == "outsourcing_list":
            rec["outsourcing"] = True
            comp = clean(d.get("outsourcing_company"))
            cont = clean(d.get("processing_content"))
            if comp or cont:
                rec["company_contents"].append((comp, cont))
        elif rec["companies"]:
            comp = rec["companies"][0]
            cont = clean(d.get("processing_content"))
            rec["company_contents"].append((comp, cont))
        out.append(rec)
    print(f"  {table}: {len(out)} 行")
    return out


# ------------------------------------------------------------
# 受注番号ごとにまとめる
# ------------------------------------------------------------
def summarize(records):
    by_key = OrderedDict()
    for r in records:
        s = by_key.setdefault(
            r["key"],
            {"sizes": [], "contents": [], "companies": [], "internal": False, "outsourcing": False, "depts": [],
             "paper_types": [], "paper_sizes": [], "paper_weights": []},
        )
        add_unique(s["sizes"], r["finish_size"])
        for pt, ps, pw in r["papers"]:
            add_unique(s["paper_types"], pt)
            add_unique(s["paper_sizes"], ps)
            add_unique(s["paper_weights"], pw)
        for c in r["contents"]:
            add_unique(s["contents"], c)
        for c in r["companies"]:
            add_unique(s["companies"], c)
        add_unique(s["depts"], r["work_department"])
        s["internal"] = s["internal"] or r["internal"] or bool(r["work_department"])
        s["outsourcing"] = s["outsourcing"] or r["outsourcing"] or bool(r["companies"])
    return by_key


def classify(s):
    if s["internal"] and s["outsourcing"]:
        return "内作・外注"
    if s["outsourcing"]:
        return "外注"
    if s["internal"]:
        return "内作"
    return ""


def write_csv(path, header, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(header)
        w.writerows(rows)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--orders", required=True, help="受注番号一覧 CSV（1列目 or 「受注番号」列）")
    ap.add_argument("--out", default="仕上りサイズ_koutei取得結果.csv", help="出力 CSV")
    ap.add_argument("--env", help="koutei の .env（DATABASE_HOST などを書いたファイル）")
    ap.add_argument("--database_host", dest="database_host")
    ap.add_argument("--database_port", dest="database_port")
    ap.add_argument("--database_name", dest="database_name")
    ap.add_argument("--database_user", dest="database_user")
    args = ap.parse_args()

    orders = read_order_numbers(args.orders)
    keys = list(orders.keys())
    print(f"受注番号一覧: {len(keys)} 件（{args.orders}）")

    conn = connect(args)
    cur = conn.cursor()
    cur.execute("SHOW default_transaction_read_only")
    if cur.fetchone()[0] != "on":
        sys.exit("読み取り専用で開けませんでした。中止します。")
    cur.execute("SELECT current_database(), inet_server_addr(), inet_server_port(), now()")
    dbinfo = cur.fetchone()
    print(f"接続確認: db={dbinfo[0]} addr={dbinfo[1]} port={dbinfo[2]} now={dbinfo[3]}")

    print("読み出し中:")
    records = []
    records += rows_from_event_table(cur, "timeline_processes", keys)
    records += rows_from_event_table(cur, "waiting_list", keys)
    records += rows_from_flat_table(
        cur, "outsourcing_list", keys,
        ["finish_size", "outsourcing_company", "processing_content", "part_name", "updated_at"],
    )
    records += rows_from_flat_table(
        cur, "delivery_schedule", keys,
        ["finish_size", "work_department", "outsourcing", "outsourcing_company", "processing_content",
         "part_name", "paper_type", "standard_size", "paper_weight", "updated_at"],
    )
    records += rows_from_flat_table(
        cur, "processing_works", keys,
        ["classification", "processing_content", "updated_at"],
    )
    conn.close()

    summary = summarize(records)

    # 本体 CSV: 入力の並びのまま、無いものは空欄
    main_rows = []
    hit = 0
    for key, raw in orders.items():
        s = summary.get(key)
        if s:
            hit += 1
            main_rows.append([raw, " / ".join(s["sizes"]), " / ".join(s["contents"]), classify(s), " / ".join(s["companies"]),
                              " / ".join(s["paper_types"]), " / ".join(s["paper_sizes"]), " / ".join(s["paper_weights"])])
        else:
            main_rows.append([raw, "", "", "", "", "", "", ""])
    write_csv(args.out, ["受注番号", "仕上りサイズ", "加工内容", "内外作区分", "委託先名", "用紙銘柄", "用紙規格", "斤量"], main_rows)

    # 詳細 CSV: 行単位（検算用）
    base, ext = os.path.splitext(args.out)
    detail_path = f"{base}_詳細{ext or '.csv'}"
    detail_rows = []
    for r in sorted(records, key=lambda r: (keys.index(r["key"]) if r["key"] in keys else 10**9, r["table"])):
        detail_rows.append([
            orders.get(r["key"], r["key"]), r["order_raw"], r["table"], r["row_id"], r["part"], r["finish_size"],
            " / ".join(r["contents"]),
            "1" if r["internal"] else "", "1" if r["outsourcing"] else "", r["work_department"],
            " / ".join(r["companies"]),
            " / ".join(f"{c}：{t}" if c and t else (c or t) for c, t in r["company_contents"]),
            " / ".join(" ".join(x for x in t if x) for t in r["papers"]),
            ",".join(r["flags"]), str(r["updated_at"] or ""),
        ])
    write_csv(
        detail_path,
        ["受注番号", "DB上の受注番号", "テーブル", "行ID", "部品", "仕上りサイズ", "加工内容", "内作フラグ", "外注フラグ",
         "加工所(work_department)", "委託先名", "委託先ごとの加工内容", "用紙（銘柄 規格 斤量）", "返し/続き", "更新日時"],
        detail_rows,
    )

    with_size = sum(1 for k in keys if summary.get(k) and summary[k]["sizes"])
    with_proc = sum(1 for k in keys if summary.get(k) and summary[k]["contents"])
    with_paper = sum(1 for k in keys if summary.get(k) and summary[k]["paper_types"])
    print()
    print(f"出力: {args.out}")
    print(f"      {detail_path}")
    print(f"該当あり {hit} / {len(keys)} 件（仕上りサイズあり {with_size} 件、加工内容あり {with_proc} 件、用紙銘柄あり {with_paper} 件、該当なし {len(keys) - hit} 件）")


if __name__ == "__main__":
    main()
