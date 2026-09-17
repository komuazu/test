#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全社印刷実績の一覧（稼動日報から作った Excel）を母集団にして、通し数 3,000 以下の案件に
帳票 PDF（製造指示書など）から読んだ 仕上りサイズ・加工・用紙 を付け、月別シートの Excel 資料にする。

    python build_final.py --base 2026年9月_全社_印刷実績.xlsx --textdir <pdftext フォルダ> \
                          --out 平版印刷_オンデマンド移行検討_3000通し以下_資料.xlsx [--max-pass 3000]

* 母集団の列: 営業部, 管理番号, ｸﾗｲｱﾝﾄ名, 品名, 今年の動向, 無しの場合の代替対策, 対策通し数, 営業担当ｺｰﾄﾞ, 印刷日, 色数, 通し数
  （見出し行が途中に何度も入っているので、管理番号が数字の行だけを明細として読む）
* 月は印刷日の月。同じ管理番号が複数行ある月はそのまま複数行出す（稼動日報の実績どおり）
* 帳票が無い管理番号は、帳票の列が空のまま並ぶ
"""

import argparse
import collections
import datetime as dt
import os
import re
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_from_pdf import load_docs, merge  # noqa: E402
from export_finish_size import norm_order  # noqa: E402

FONT_NAME = "Meiryo"
FONT = Font(name=FONT_NAME, size=9)
FONT_BOLD = Font(name=FONT_NAME, size=9, bold=True)
FONT_TITLE = Font(name=FONT_NAME, size=14, bold=True, color="1F4E78")
FONT_NOTE = Font(name=FONT_NAME, size=8, color="595959")
FONT_HEADER = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
FONT_LINK = Font(name=FONT_NAME, size=9, color="0563C1", underline="single")
FILL_HEADER = PatternFill("solid", fgColor="1F4E78")
FILL_HEADER_K = PatternFill("solid", fgColor="2E75B6")
FILL_BAND = PatternFill("solid", fgColor="F2F2F2")
FILL_A3 = {"○": PatternFill("solid", fgColor="C6EFCE"), "×": PatternFill("solid", fgColor="F8CBAD"),
           "不明": PatternFill("solid", fgColor="FFF2CC")}
FILL_NONE = PatternFill("solid", fgColor="E7E6E6")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
ALIGN_WRAP = Alignment(vertical="top", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="top")

BASE_COLS = [("営業部", 13), ("管理番号", 10), ("ｸﾗｲｱﾝﾄ名", 20), ("品名", 26), ("営業担当\nｺｰﾄﾞ", 7), ("印刷日", 11), ("色数", 6), ("通し数", 8)]
DOC_COLS = [("仕上りサイズ", 14), ("A3以下", 7), ("加工内容", 26), ("内外作", 8), ("加工所／委託先", 14),
            ("用紙銘柄", 18), ("用紙規格", 10), ("連量", 8), ("総頁数", 6), ("受注数量", 8), ("備考案", 34),
            ("元の帳票", 12), ("帳票\nリンク", 7), ("加工欄の原文", 40)]
COLUMNS = [("No", 5)] + BASE_COLS + DOC_COLS
N_BASE = 1 + len(BASE_COLS)


def toi(v):
    try:
        return int(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return None


def read_base(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = []
    for r in ws.iter_rows(min_row=1, values_only=True):
        if r is None or len(r) < 11:
            continue
        key, _ = norm_order(r[1])
        if not key or str(r[1]).strip() == "管理番号":
            continue
        rows.append({
            "key": key, "営業部": r[0] or "", "管理番号": str(r[1]).strip(), "ｸﾗｲｱﾝﾄ名": r[2] or "", "品名": r[3] or "",
            "営業担当ｺｰﾄﾞ": r[7] or "", "印刷日": str(r[8] or "")[:10], "色数": r[9] or "", "通し数": toi(r[10]),
        })
    return rows


def remark(d):
    """備考欄の文面の案。例: 外注（松岡製本）：ミシン(筋) ／ 内作（第二工場）：中綴じ12P"""
    if not d:
        return ""
    kind, site, cont = d["内外作"], d["加工所"], d["加工内容"]
    if not (kind or site or cont):
        return ""
    head = kind or ""
    if site:
        head = f"{head}（{site}）" if head else f"（{site}）"
    return f"{head}：{cont}" if head and cont else (head or cont)


def style_header(ws, hrow, n_base, n_all):
    for c in range(1, n_all + 1):
        cell = ws.cell(hrow, c)
        cell.font, cell.alignment, cell.border = FONT_HEADER, ALIGN_CENTER, BORDER
        cell.fill = FILL_HEADER if c <= n_base else FILL_HEADER_K
    ws.row_dimensions[hrow].height = 30


def write_month(wb, title, rows, docs, note):
    ws = wb.create_sheet(title)
    ws["A1"] = f"平版印刷 → オンデマンド移行検討　{title}　通し数 3,000 以下"
    ws["A1"].font = FONT_TITLE
    ws["A2"] = note
    ws["A2"].font = FONT_NOTE
    hrow = 4
    for i, (h, w) in enumerate(COLUMNS, 1):
        ws.cell(hrow, i, h)
        ws.column_dimensions[get_column_letter(i)].width = w
    style_header(ws, hrow, N_BASE, len(COLUMNS))
    r = hrow
    for n, x in enumerate(rows, 1):
        r += 1
        d = docs.get(x["key"])
        vals = [n, x["営業部"], x["管理番号"], x["ｸﾗｲｱﾝﾄ名"], x["品名"], x["営業担当ｺｰﾄﾞ"], x["印刷日"], x["色数"], x["通し数"]]
        if d:
            vals += [d["仕上りサイズ"], d["A3以下"], d["加工内容"], d["内外作"], d["加工所"], d["用紙銘柄"], d["用紙規格"], d["連量"],
                     d["総頁数"], d["受注数量"], remark(d), "・".join(d["帳票"]), "開く", d["加工原文"]]
        else:
            vals += ["", "帳票なし"] + [""] * (len(DOC_COLS) - 2)
        band = n % 2 == 0
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v)
            c.font, c.border, c.alignment = FONT, BORDER, ALIGN_WRAP
            if band:
                c.fill = FILL_BAND
        a3 = ws.cell(r, N_BASE + 2)
        if a3.value in FILL_A3:
            a3.fill, a3.alignment = FILL_A3[a3.value], ALIGN_CENTER
        elif a3.value == "帳票なし":
            a3.fill, a3.alignment, a3.font = FILL_NONE, ALIGN_CENTER, FONT_NOTE
        if d:
            link = ws.cell(r, N_BASE + len(DOC_COLS) - 1)
            link.hyperlink, link.font = d["url"], FONT_LINK
        for i in (9, N_BASE + 9, N_BASE + 10):
            ws.cell(r, i).number_format = "#,##0"
            ws.cell(r, i).alignment = ALIGN_RIGHT
    ws.freeze_panes = ws.cell(hrow + 1, 4)
    ws.auto_filter.ref = f"A{hrow}:{get_column_letter(len(COLUMNS))}{max(r, hrow + 1)}"
    ws.print_title_rows = f"{hrow}:{hrow}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddFooter.center.text = "&P / &N"
    return hrow + 1, r


def write_summary(wb, months, ranges, stats, base_name):
    ws = wb.create_sheet("集計", 0)
    ws["A1"] = "平版印刷 → オンデマンド移行検討（通し数 3,000 以下）　月別の集計"
    ws["A1"].font = FONT_TITLE
    ws["A2"] = f"作成 {dt.date.today().isoformat()}　母集団: {base_name}　件数は各月シートを COUNTIF で数えている（月シートを直せば追従する）"
    ws["A2"].font = FONT_NOTE
    heads = ["月", "件数", "帳票あり", "帳票なし", "仕上りサイズあり", "A3以下 ○", "A3以下 ×", "A3以下 不明",
             "内作", "外注", "内作・外注", "加工内容あり", "用紙銘柄あり"]
    hrow = 4
    for i, h in enumerate(heads, 1):
        ws.cell(hrow, i, h)
    style_header(ws, hrow, len(heads), len(heads))
    L = get_column_letter
    c_size, c_a3, c_cont, c_kind, c_paper = L(N_BASE + 1), L(N_BASE + 2), L(N_BASE + 3), L(N_BASE + 4), L(N_BASE + 6)
    r = hrow
    for mth in months:
        r += 1
        first, last = ranges[mth]
        q = f"'{mth}'"
        def rng(col):
            return f"{q}!{col}{first}:{col}{last}"
        vals = [mth, f'=COUNTA({rng("C")})',
                f'=COUNTA({rng("C")})-COUNTIF({rng(c_a3)},"帳票なし")', f'=COUNTIF({rng(c_a3)},"帳票なし")',
                f'=COUNTA({rng(c_size)})',
                f'=COUNTIF({rng(c_a3)},"○")', f'=COUNTIF({rng(c_a3)},"×")', f'=COUNTIF({rng(c_a3)},"不明")',
                f'=COUNTIF({rng(c_kind)},"内作")', f'=COUNTIF({rng(c_kind)},"外注")', f'=COUNTIF({rng(c_kind)},"内作・外注")',
                f'=COUNTA({rng(c_cont)})', f'=COUNTA({rng(c_paper)})']
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v)
            c.font, c.border = FONT, BORDER
    r += 1
    ws.cell(r, 1, "合計").font = FONT_BOLD
    ws.cell(r, 1).border = BORDER
    for i in range(2, len(heads) + 1):
        col = L(i)
        c = ws.cell(r, i, f"=SUM({col}{hrow + 1}:{col}{r - 1})")
        c.font, c.border = FONT_BOLD, BORDER
    r += 2
    ws.cell(r, 1, "凡例").font = FONT_BOLD
    legend = [("○", FILL_A3["○"], "仕上りサイズが A3 以下（オンデマンド機に載る大きさ）"),
              ("×", FILL_A3["×"], "1 部品でも A3 より大きい"),
              ("不明", FILL_A3["不明"], "仕上りサイズが「規格外」だけで実寸が無い"),
              ("帳票なし", FILL_NONE, "ドライブに製造指示書などの PDF が無く、帳票の列が埋められなかった")]
    for label, fill, text in legend:
        r += 1
        c = ws.cell(r, 1, label)
        c.fill, c.font, c.border, c.alignment = fill, FONT, BORDER, ALIGN_CENTER
        ws.cell(r, 2, text).font = FONT
    r += 2
    notes = ["この資料の作り方と注意",
             f"・母集団は「{base_name}」の明細のうち通し数 3,000 以下の {stats['rows']} 行（管理番号 {stats['keys']} 件）。月は印刷日の月",
             f"・帳票の列は Google ドライブ「作業依頼書」「委託業務依頼書」の PDF（製造指示書 {stats['seizo']}、印刷指示書 {stats['insatsu']}、外注委託依頼書 {stats['itaku']} 通）の文字を読んで管理番号で付けた",
             "・加工内容は帳票の加工欄から機械的に拾った言葉。迷うときは「加工欄の原文」列と「帳票リンク」で元を見る",
             "・備考案は 外注（委託先）：加工内容 ／ 内作（加工所）：加工内容 の形。元の Excel の備考欄に貼るときの案",
             "・製造指示書の得意先名・製品名の切れ目、用紙の銘柄の切れ目は機械的に推定している",
             "・同じ管理番号が複数行ある月は、稼動日報の実績どおり複数行のまま。帳票の列は同じ値が入る"]
    for i, t in enumerate(notes):
        ws.cell(r + i, 1, t).font = FONT_BOLD if i == 0 else FONT_NOTE
    for i, w in enumerate([10, 8, 9, 9, 13, 10, 10, 11, 8, 8, 11, 12, 12], 1):
        ws.column_dimensions[L(i)].width = w
    ws.freeze_panes = ws.cell(hrow + 1, 2)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--base", required=True, help="全社印刷実績の Excel")
    ap.add_argument("--textdir", required=True, help="帳票 PDF の文字（pdftext）フォルダ")
    ap.add_argument("--out", default="平版印刷_オンデマンド移行検討_3000通し以下_資料.xlsx")
    ap.add_argument("--max-pass", type=int, default=3000)
    ap.add_argument("--csv", help="同じ内容の CSV も出す")
    args = ap.parse_args()

    base = read_base(args.base)
    small = [x for x in base if x["通し数"] is not None and x["通し数"] <= args.max_pass]
    print(f"母集団: 明細 {len(base)} 行 → 通し数 {args.max_pass:,} 以下 {len(small)} 行（管理番号 {len({x['key'] for x in small})} 件）")

    all_docs = load_docs(args.textdir)
    docs = {r["key"]: r for r in merge(all_docs)}
    stats = {"rows": len(small), "keys": len({x["key"] for x in small}),
             "seizo": sum(d["doc"] == "製造指示書" for d in all_docs), "insatsu": sum(d["doc"] == "印刷指示書" for d in all_docs),
             "itaku": sum(d["doc"] == "外注委託依頼書" for d in all_docs)}
    hit = sum(1 for x in small if x["key"] in docs)
    print(f"帳票: {len(all_docs)} 通・受注 {len(docs)} 件 → 母集団と一致 {hit} / {len(small)} 行")

    by_month = collections.defaultdict(list)
    for x in small:
        m = x["印刷日"][:7].replace("/", "-") or "月不明"
        by_month[m].append(x)
    months = sorted(by_month)
    for m in months:
        by_month[m].sort(key=lambda x: (x["印刷日"], x["営業部"], x["管理番号"]))

    wb = Workbook()
    wb.remove(wb.active)
    ranges = {}
    note = ("左の紺色の列は稼動日報（全社印刷実績）、右の青色の列は製造指示書などの帳票 PDF から。"
            "A3以下: 緑=○／橙=×／黄=不明／灰=帳票なし。加工内容は機械的に拾った言葉なので、迷うときは右端の原文とリンク先で確かめる")
    for m in months:
        ranges[m] = write_month(wb, m, by_month[m], docs, note)
        print(f"  {m}: {len(by_month[m])} 行、帳票あり {sum(1 for x in by_month[m] if x['key'] in docs)}")
    write_summary(wb, months, ranges, stats, os.path.basename(args.base))
    wb.save(args.out)
    print(f"出力: {args.out}")

    if args.csv:
        import csv
        with open(args.csv, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, lineterminator="\r\n")
            w.writerow(["月"] + [h.replace("\n", "") for h, _ in BASE_COLS] + [h.replace("\n", "") for h, _ in DOC_COLS if h != "帳票\nリンク"] + ["リンク"])
            for m in months:
                for x in by_month[m]:
                    d = docs.get(x["key"])
                    row = [m, x["営業部"], x["管理番号"], x["ｸﾗｲｱﾝﾄ名"], x["品名"], x["営業担当ｺｰﾄﾞ"], x["印刷日"], x["色数"], x["通し数"]]
                    if d:
                        row += [d["仕上りサイズ"], d["A3以下"], d["加工内容"], d["内外作"], d["加工所"], d["用紙銘柄"], d["用紙規格"], d["連量"],
                                d["総頁数"], d["受注数量"], remark(d), "・".join(d["帳票"]), d["加工原文"], d["url"]]
                    else:
                        row += ["", "帳票なし"] + [""] * (len(DOC_COLS) - 2)
                    w.writerow(row)
        print(f"      {args.csv}")


if __name__ == "__main__":
    main()
