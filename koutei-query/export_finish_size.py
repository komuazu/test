#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
koutei-kanr30（社内工程管理システム）の本番DBから、受注番号ごとの
「仕上りサイズ・加工内容・内外作区分・委託先名・用紙」を読み出して CSV にする。

* 読み取り専用。DB もコードも一切書き換えない
  （接続を default_transaction_read_only=on にして開くので、書き込み文は DB 側で拒否される）
* 本番機（丸本PC）の PostgreSQL は localhost:5432 にしか口を開けていないので、
  このスクリプトは丸本PC の上で動かす（詳しくは README.md）

使い方（丸本PC、コマンドプロンプト）:

    cd <このフォルダ>
    python export_finish_size.py --orders 仕上りサイズ未確認_受注番号一覧.csv ^
        --env C:\\Users\\116544\\Desktop\\UPDATA17\\web_app\\.env

出力:
    仕上りサイズ_koutei取得結果.csv        … 受注番号, 仕上りサイズ, A3以下, 加工内容, 内外作区分, 委託先名, 用紙銘柄, 用紙規格, 斤量
    仕上りサイズ_koutei取得結果_A3以下.csv … 上のうち A3 以下（判定○）だけ
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
except ImportError:  # pragma: no cover
    sys.exit("psycopg2 が入っていません。 pip install psycopg2-binary で入れてください。")


# ------------------------------------------------------------
# 受注番号のゆれ吸収
# ------------------------------------------------------------
FULLWIDTH_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")


def norm_order(value):
    """'08726258' / '8726258' / '8726258.0' / '８７２６２５８' / '08726258（表紙）' / '08726258②' を
    同じキー '8726258' にし、枝（（表紙）・②・複製1 など）を別に返す。
    NFKC はかけない（② が 2 になって別の受注番号に化けるため）。数字の部分だけ半角にする。"""
    if value is None:
        return "", ""
    s = re.sub(r"\s+", "", str(value))
    m = re.match(r"^([0-9０-９]+)(.*)$", s)
    if not m:
        return "", s
    digits = m.group(1).translate(FULLWIDTH_DIGITS).lstrip("0")
    rest = m.group(2)
    if re.fullmatch(r"\.0+", rest):  # Excel 由来の 8726258.0
        rest = ""
    elif re.match(r"^\.\d", rest):  # 8726274.5 のような小数は受注番号ではない
        return "", s
    if not digits:  # ゼロだけ（0, 000）は受注番号ではない
        return "", s
    return digits, rest


# SQL 側で同じ規則（空白を落とし、先頭の数字だけ取り、前ゼロを落とす）
SQL_NORM = (
    "regexp_replace(COALESCE(substring(regexp_replace(translate(COALESCE({col}, ''), '０１２３４５６７８９', '0123456789'), "
    "'\\s', '', 'g') from '^\\d+'), ''), '^0+', '')"
)

# 「値なし」とみなす文字列
EMPTY_WORDS = {"", "-", "－", "―", "未設定", "未定", "なし", "無し", "none", "null", "nan"}
# 委託先名として意味のないもの
NOT_A_COMPANY = EMPTY_WORDS | {"未手配", "内作"}
# 加工内容としては意味のないもの（MIS の仕上加工名「その他」、processing_works.classification の「その他」など）
GENERIC_CONTENT = {"その他", "内作", "外注"}
NOT_A_CONTENT = EMPTY_WORDS | GENERIC_CONTENT
# work_department のうち内作（自社の加工）と判定するもの（画面の選択肢: -/第二工場/POP課/本社/営業/その他）
INTERNAL_DEPARTMENTS = {"第二工場", "本社", "POP課"}


def truthy(v):
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    return str(v).strip().lower() in ("1", "true", "yes", "on")


def clean(v, empty=EMPTY_WORDS):
    """文字列だけを返す。dict/list や「-」「未設定」などは空扱い。"""
    if v is None or isinstance(v, (dict, list, bool)):
        return ""
    s = str(v).strip()
    return "" if s.lower() in empty else s


def add_unique(lst, value, empty=EMPTY_WORDS):
    value = clean(value, empty)
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
        k, rest = norm_order(header[0])
        if not k or rest:
            start = 1  # 受注番号として丸ごと読めない1行目（「2026年9月 一覧」など）は見出しとみなす

    result = OrderedDict()  # 正規化キー → 元の表記（最初に出たもの）
    dropped = []
    for r in rows[start:]:
        if col >= len(r):
            continue
        key, _ = norm_order(r[col])
        if not key:
            if r[col].strip():
                dropped.append(r[col].strip())
            continue
        if key not in result:
            result[key] = r[col].strip()
    if dropped:
        print(f"※ 受注番号として読めなかった値 {len(dropped)} 件を飛ばしました: {dropped[:10]}")
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

    def get(k, d=None):
        return args.__dict__.get(k.lower()) or os.environ.get(k) or env.get(k) or d

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


def new_record(table, row_id, order_raw, updated_at="", created_at=""):
    key, suffix = norm_order(order_raw)
    return {
        "table": table,
        "row_id": row_id,
        "order_raw": order_raw,
        "key": key,
        "suffix": suffix,           # 受注番号に付いていた枝（（表紙）・②・複製1 など）
        "part": "",
        "sizes": [],
        "papers": [],               # (銘柄, 規格, 斤量)
        "contents": [],
        "internal": False,
        "outsourcing": False,
        "work_department": "",
        "companies": [],
        "company_contents": [],     # (委託先, 加工内容)
        "flags": [],
        "updated_at": updated_at or "",
        "created_at": created_at or "",
    }


def absorb_event_dict(rec, d):
    """イベントの JSON（最上位でも events[] の要素でも同じ形）から項目を拾う。"""
    if not rec["part"]:
        rec["part"] = clean(d.get("part_type") or d.get("part_name") or d.get("partName"))
    add_unique(rec["sizes"], d.get("finish_size") or d.get("finished_size") or d.get("finishSize"))
    for k in ("finish_processing_name", "finish_processing", "finishProcessing", "processing_content", "finish_process"):
        add_unique(rec["contents"], d.get(k), NOT_A_CONTENT)
    rec["internal"] = rec["internal"] or truthy(d.get("internal_work"))
    rec["outsourcing"] = rec["outsourcing"] or truthy(d.get("outsourcing"))
    dept = clean(d.get("work_department"))
    if dept and not rec["work_department"]:
        rec["work_department"] = dept
    for k in ("outsourcing_company", "outsource_name"):
        add_unique(rec["companies"], d.get(k), NOT_A_COMPANY)
    items = d.get("outsourcing_items")
    if isinstance(items, list):
        for it in items:
            if not isinstance(it, dict):
                continue
            comp = clean(it.get("company") or it.get("company_name") or it.get("name"), NOT_A_COMPANY)
            cont = clean(it.get("processing_content") or it.get("content"), NOT_A_CONTENT)
            add_unique(rec["companies"], comp, NOT_A_COMPANY)
            add_unique(rec["contents"], cont, NOT_A_CONTENT)
            if (comp or cont) and (comp, cont) not in rec["company_contents"]:
                rec["company_contents"].append((comp, cont))
    add_paper(rec, d.get("paper_type") or d.get("paperType"), d.get("standard_size"), d.get("paper_weight") or d.get("paperWeight"))
    for lst_key in ("order_paper_info", "papers"):  # 部品ごとの用紙（外注委託依頼書で使う形）
        lst = d.get(lst_key)
        if isinstance(lst, list):
            for pi in lst:
                if isinstance(pi, dict):
                    add_paper(rec, pi.get("paper_type"), pi.get("standard_size"), pi.get("paper_weight"))
    for k in ("isReturnProcess", "isContinueProcess", "is_return", "is_continue"):
        if truthy(d.get(k)) and k not in rec["flags"]:
            rec["flags"].append(k)


def absorb_row_json(rec, d):
    """data 列の JSON 1件ぶん: 最上位と events[] の要素を読む。"""
    absorb_event_dict(rec, d)
    events = d.get("events")
    if isinstance(events, list):
        for e in events:
            if not isinstance(e, dict):
                continue
            # 要素に別の受注番号が書いてあることがある（配置時のコピーで持ち越されたもの）。それは飛ばす
            e_key, _ = norm_order(e.get("order_number") or e.get("orderNumber") or "")
            if e_key and e_key != rec["key"]:
                rec["flags"].append(f"events:{e.get('order_number') or e.get('orderNumber')}")
                continue
            absorb_event_dict(rec, e)


def rows_from_event_table(cur, table, keys):
    """timeline_processes / waiting_list: 明細は data 列の JSON に入っている。
    MIS の CSV から取り込んだ項目（仕上加工名など）は data.events[] の中にしか無いことがある。"""
    cols = existing_columns(cur, table)
    if not cols or "data" not in cols:
        print(f"  {table}: テーブルまたは data 列が無いので飛ばします")
        return []
    ocol = '"orderNumber"' if "orderNumber" in cols else ("order_number" if "order_number" in cols else None)
    if ocol is None:
        print(f"  {table}: 受注番号の列が無いので飛ばします")
        return []
    extra = [c for c in ("updated_at", "created_at") if c in cols]
    sel = ", ".join(["id", ocol, "data"] + extra)
    cur.execute(
        f"SELECT {sel} FROM {table} WHERE {SQL_NORM.format(col=ocol)} = ANY(%s) ORDER BY id",
        (keys,),
    )
    out = []
    for row in cur.fetchall():
        rec = new_record(table, row[0], row[1],
                         row[3 + extra.index("updated_at")] if "updated_at" in extra else "",
                         row[3 + extra.index("created_at")] if "created_at" in extra else "")
        absorb_row_json(rec, parse_json(row[2]))
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
        f"SELECT {sel} FROM {table} WHERE {SQL_NORM.format(col=ocol)} = ANY(%s) ORDER BY id",
        (keys,),
    )
    out = []
    for row in cur.fetchall():
        d = dict(zip(wanted, row[2:]))
        rec = new_record(table, row[0], row[1], d.get("updated_at"), d.get("created_at"))
        rec["part"] = clean(d.get("part_name"))
        add_unique(rec["sizes"], d.get("finish_size"))
        rec["work_department"] = clean(d.get("work_department"))
        rec["outsourcing"] = truthy(d.get("outsourcing"))
        add_unique(rec["contents"], d.get("processing_content"), NOT_A_CONTENT)
        comp = clean(d.get("outsourcing_company"), NOT_A_COMPANY)
        add_unique(rec["companies"], comp, NOT_A_COMPANY)
        add_paper(rec, d.get("paper_type"), d.get("standard_size"), d.get("paper_weight"))
        if table == "processing_works":
            # 内作加工の作業一覧。classification は 折加工1／折加工2／中綴じ加工 などの機械区分
            rec["internal"] = True
            add_unique(rec["contents"], d.get("classification"), NOT_A_CONTENT)
        if table == "outsourcing_list":
            rec["outsourcing"] = True
        cont = clean(d.get("processing_content"), NOT_A_CONTENT)
        if comp and (comp, cont) not in rec["company_contents"]:
            rec["company_contents"].append((comp, cont))
        out.append(rec)
    print(f"  {table}: {len(out)} 行")
    return out


# ------------------------------------------------------------
# 仕上りサイズが A3 以下か（オンデマンド機に載るか）
# ------------------------------------------------------------
A3_MM = (297, 420)
# 規格名 → 仕上り mm（短辺, 長辺）
SIZE_MM = {
    "a0": (841, 1189), "a1": (594, 841), "a2": (420, 594), "a3": (297, 420), "a4": (210, 297), "a5": (148, 210),
    "a6": (105, 148), "a7": (74, 105), "a8": (52, 74),
    "b0": (1030, 1456), "b1": (728, 1030), "b2": (515, 728), "b3": (364, 515), "b4": (257, 364), "b5": (182, 257),
    "b6": (128, 182), "b7": (91, 128), "b8": (64, 91),
}
# 名前だけで小さいと分かるもの
SMALL_WORDS = ("ハガキ", "はがき", "葉書", "名刺", "カード", "長3", "長4", "長40", "角2", "角3", "角形", "洋形", "洋長",
               "封筒", "ショップカード", "ポストカード", "cd", "dvd", "しおり", "チケット", "シール", "ラベル", "a4以下", "b5以下")
# 名前だけで A3 より大きいと分かるもの
LARGE_WORDS = ("a全", "b全", "菊全", "菊判", "四六", "菊半", "a倍", "b倍", "ポスター", "a3ノビ", "a3のび", "a3+", "sra3", "a3伸")


def size_mm(text):
    """'A4 297×210' / '297*210' / '29.7×21cm' から (短辺, 長辺) mm を取り出す。無ければ None。"""
    # 「a4 297×210」の 4 を数字に巻き込まないよう、直前が英数字でない所から読む
    m = re.search(r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(?:mm|㎜)?\s*[x×*＊]\s*(\d+(?:\.\d+)?)\s*(mm|㎜|cm|㎝)?", text)
    if not m:
        return None
    a, b = float(m.group(1)), float(m.group(2))
    if m.group(3) in ("cm", "㎝") or (a < 50 and b < 50):
        a, b = a * 10, b * 10
    return (min(a, b), max(a, b))


def judge_a3(size_text):
    """1つの仕上りサイズ表記を '○'（A3以下）/ '×'（A3より大きい）/ '?'（分からない）にする。"""
    t0 = unicodedata.normalize("NFKC", size_text).lower().strip()
    t = re.sub(r"\s+", "", t0)
    if not t:
        return "?"
    mm = size_mm(t0)
    if mm:
        return "○" if mm[0] <= A3_MM[0] + 3 and mm[1] <= A3_MM[1] + 3 else "×"
    for w in LARGE_WORDS:
        if w.lower() in t:
            return "×"
    for w in SMALL_WORDS:
        if w.lower() in t:
            return "○"
    m = re.search(r"(?<![a-z0-9])([ab])(\d{1,2})(?![0-9])", t)
    if m:
        key = m.group(1) + m.group(2)
        if key in SIZE_MM:
            mm = SIZE_MM[key]
            return "○" if mm[0] <= A3_MM[0] and mm[1] <= A3_MM[1] else "×"
    return "?"


def judge_a3_all(sizes):
    """受注のサイズ一覧（部品ぶん）をまとめて判定。1つでも A3 より大きければ ×、全部不明なら「不明」。"""
    if not sizes:
        return "不明"
    marks = [judge_a3(x) for x in sizes]
    if "×" in marks:
        return "×"
    if "○" in marks:
        return "○"
    return "不明"


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
        for v in r["sizes"]:
            add_unique(s["sizes"], v)
        for pt, ps, pw in r["papers"]:
            add_unique(s["paper_types"], pt)
            add_unique(s["paper_sizes"], ps)
            add_unique(s["paper_weights"], pw)
        for c in r["contents"]:
            add_unique(s["contents"], c, NOT_A_CONTENT)
        for c in r["companies"]:
            add_unique(s["companies"], c, NOT_A_COMPANY)
        add_unique(s["depts"], r["work_department"])
        s["internal"] = s["internal"] or r["internal"] or (r["work_department"] in INTERNAL_DEPARTMENTS)
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


MAIN_HEADER = ["受注番号", "仕上りサイズ", "A3以下", "加工内容", "内外作区分", "委託先名", "用紙銘柄", "用紙規格", "斤量"]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--orders", required=True, help="受注番号一覧 CSV（1列目 or 「受注番号」列）")
    ap.add_argument("--out", default="仕上りサイズ_koutei取得結果.csv", help="出力 CSV")
    ap.add_argument("--a3-only", action="store_true", help="本体 CSV も A3 以下（判定○）の受注番号だけにする")
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
        ["finish_size", "outsourcing_company", "processing_content", "part_name", "updated_at", "created_at"],
    )
    records += rows_from_flat_table(
        cur, "delivery_schedule", keys,
        ["finish_size", "work_department", "outsourcing", "outsourcing_company", "processing_content",
         "part_name", "paper_type", "standard_size", "paper_weight", "updated_at", "created_at"],
    )
    records += rows_from_flat_table(
        cur, "processing_works", keys,
        ["classification", "processing_content", "updated_at", "created_at"],
    )
    conn.close()

    summary = summarize(records)

    # 本体 CSV: 入力の並びのまま、無いものは空欄
    main_rows = []
    a3_rows = []
    hit = 0
    a3_count = {"○": 0, "×": 0, "不明": 0}
    for key, raw in orders.items():
        s = summary.get(key)
        if s:
            hit += 1
            a3 = judge_a3_all(s["sizes"])
            row = [raw, " / ".join(s["sizes"]), a3, " / ".join(s["contents"]), classify(s), " / ".join(s["companies"]),
                   " / ".join(s["paper_types"]), " / ".join(s["paper_sizes"]), " / ".join(s["paper_weights"])]
        else:
            a3 = "不明"
            row = [raw, "", a3] + [""] * (len(MAIN_HEADER) - 3)
        a3_count[a3] += 1
        if a3 == "○":
            a3_rows.append(row)
        if not args.a3_only or a3 == "○":
            main_rows.append(row)
    write_csv(args.out, MAIN_HEADER, main_rows)
    base, ext = os.path.splitext(args.out)
    a3_path = f"{base}_A3以下{ext or '.csv'}"
    write_csv(a3_path, MAIN_HEADER, a3_rows)

    # 詳細 CSV: 行単位（検算用）
    detail_path = f"{base}_詳細{ext or '.csv'}"
    pos = {k: i for i, k in enumerate(keys)}
    detail_rows = []
    for r in sorted(records, key=lambda r: (pos.get(r["key"], 10**9), r["table"], str(r["row_id"]))):
        detail_rows.append([
            orders.get(r["key"], r["key"]), r["order_raw"], r["suffix"], r["table"], r["row_id"], r["part"],
            " / ".join(r["sizes"]), " / ".join(r["contents"]),
            "1" if r["internal"] else "", "1" if r["outsourcing"] else "", r["work_department"],
            " / ".join(r["companies"]),
            " / ".join(f"{c}：{t}" if c and t else (c or t) for c, t in r["company_contents"]),
            " / ".join(" ".join(x for x in t if x) for t in r["papers"]),
            ",".join(r["flags"]), str(r["created_at"] or ""), str(r["updated_at"] or ""),
        ])
    write_csv(
        detail_path,
        ["受注番号", "DB上の受注番号", "枝", "テーブル", "行ID", "部品", "仕上りサイズ", "加工内容", "内作フラグ", "外注フラグ",
         "加工所(work_department)", "委託先名", "委託先ごとの加工内容", "用紙（銘柄 規格 斤量）", "返し/続き", "作成日時", "更新日時"],
        detail_rows,
    )

    with_size = sum(1 for k in keys if summary.get(k) and summary[k]["sizes"])
    with_proc = sum(1 for k in keys if summary.get(k) and summary[k]["contents"])
    with_paper = sum(1 for k in keys if summary.get(k) and summary[k]["paper_types"])
    only_b1 = sum(1 for k in keys if summary.get(k) and summary[k]["sizes"] == ["B1"])
    print()
    print(f"出力: {args.out}" + ("（A3 以下だけ）" if args.a3_only else ""))
    print(f"      {a3_path}（A3 以下だけ）")
    print(f"      {detail_path}")
    print(f"A3 以下の判定: ○ {a3_count['○']} 件、× {a3_count['×']} 件、不明 {a3_count['不明']} 件"
          f"（× は 1 部品でも A3 より大きいもの。不明は仕上りサイズが無いか読めないもの）")
    print(f"該当あり {hit} / {len(keys)} 件（仕上りサイズあり {with_size} 件、加工内容あり {with_proc} 件、"
          f"用紙銘柄あり {with_paper} 件、該当なし {len(keys) - hit} 件）")
    if only_b1:
        print(f"※ 仕上りサイズが「B1」だけの案件が {only_b1} 件あります。MIS からの一括取込で入った既定値の疑いがあるので、"
              f"詳細 CSV の作成日時（同じ時刻に大量に入っていないか）と MIS 側の値を確かめてください。")


if __name__ == "__main__":
    main()
