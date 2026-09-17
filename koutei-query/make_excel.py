#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
export_finish_size.py の結果 CSV を、元の Excel（平版印刷_オンデマンド移行検討_3000通し以下一覧.xlsx）の
月別シートに流し込み、資料として見られる形に整えて別名で保存する。

    python make_excel.py --base 平版印刷_オンデマンド移行検討_3000通し以下一覧.xlsx ^
                         --result 仕上りサイズ_koutei取得結果.csv

* 元の Excel は書き換えない。`<元の名前>_koutei反映.xlsx` を新しく作る
* 月別シート（シート名に「月」が入り、「受注番号」または「管理番号」の見出しがある）を対象にする。
  その他のシート（サマリーなど）はそのまま残す
* 「仕上りサイズ」「備考」は **空いている欄だけ** 埋める（手で記入済みの 46 件は触らない）。
  埋めた欄は水色で塗る
* 右端に koutei の列（A3以下・加工内容・内外作区分・委託先名・用紙銘柄・用紙規格・斤量・koutei仕上りサイズ）を足す
* 見出し・罫線・列幅・ウィンドウ枠の固定・オートフィルタを整え、A3以下の判定を色で分ける
* 末尾に「集計（koutei）」シートを足す。件数は COUNTIF の式で、月別シートを直せば追従する

--base を省くと、結果 CSV だけから「一覧」シートの Excel を作る（月別には分かれない）。
"""

import argparse
import csv
import datetime as dt
import os
import re
import sys

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.exit("openpyxl が入っていません。 pip install openpyxl で入れてください。")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from export_finish_size import norm_order  # noqa: E402

# 結果 CSV の列（export_finish_size.py の MAIN_HEADER と同じ）
RESULT_COLS = ["受注番号", "仕上りサイズ", "A3以下", "加工内容", "内外作区分", "委託先名", "用紙銘柄", "用紙規格", "斤量"]
# 月別シートの右端に足す列（結果 CSV の列名 → 見出し）
APPEND_COLS = [
    ("A3以下", "A3以下\n(koutei判定)"),
    ("加工内容", "加工内容\n(koutei)"),
    ("内外作区分", "内外作\n(koutei)"),
    ("委託先名", "委託先名\n(koutei)"),
    ("用紙銘柄", "用紙銘柄\n(koutei)"),
    ("用紙規格", "用紙規格\n(koutei)"),
    ("斤量", "斤量\n(koutei)"),
    ("仕上りサイズ", "仕上りサイズ\n(koutei)"),
]

FONT_NAME = "Meiryo"
FONT = Font(name=FONT_NAME, size=9)
FONT_BOLD = Font(name=FONT_NAME, size=9, bold=True)
FONT_TITLE = Font(name=FONT_NAME, size=14, bold=True, color="1F4E78")
FONT_NOTE = Font(name=FONT_NAME, size=8, color="595959")
FONT_HEADER = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
FILL_HEADER = PatternFill("solid", fgColor="1F4E78")
FILL_HEADER_K = PatternFill("solid", fgColor="2E75B6")   # koutei から足した列の見出し
FILL_BAND = PatternFill("solid", fgColor="F2F2F2")
FILL_FILLED = PatternFill("solid", fgColor="DDEBF7")     # ツールが埋めた欄
FILL_A3 = {"○": PatternFill("solid", fgColor="C6EFCE"), "×": PatternFill("solid", fgColor="F8CBAD"),
           "不明": PatternFill("solid", fgColor="FFF2CC")}
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
ALIGN_WRAP = Alignment(vertical="top", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


# ------------------------------------------------------------
# 結果 CSV
# ------------------------------------------------------------
def read_result(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp932"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        sys.exit(f"{path} の文字コードが読めません")
    rows = list(csv.DictReader(text.splitlines()))
    missing = [c for c in RESULT_COLS if rows and c not in rows[0]]
    if missing:
        sys.exit(f"結果 CSV に列がありません: {missing}（export_finish_size.py の出力を渡してください）")
    result = {}
    for r in rows:
        k, _ = norm_order(r["受注番号"])
        if k and k not in result:
            result[k] = r
    return result


def remark_draft(r):
    """備考欄の文面の案。例: 外注（松岡製本）：抜き・ポケット貼り・24P中綴じ ／ 内作（第二工場）：折加工"""
    kind, comp, cont = r.get("内外作区分", ""), r.get("委託先名", ""), r.get("加工内容", "")
    if not (kind or comp or cont):
        return ""
    head = kind or ("外注" if comp else "")
    if comp:
        head = f"{head}（{comp}）" if head else f"（{comp}）"
    return f"{head}：{cont}" if head and cont else (head or cont)


# ------------------------------------------------------------
# 月別シートの見出しを探す
# ------------------------------------------------------------
def find_header(ws):
    """先頭 15 行から「受注番号／管理番号」の見出し行を探す。(行番号, {見出し: 列番号}) を返す。"""
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 15)):
        for cell in row:
            v = str(cell.value) if cell.value is not None else ""
            if "受注番号" in v or "管理番号" in v:
                headers = {}
                for c in row:
                    if c.value is not None and str(c.value).strip():
                        headers[str(c.value).strip()] = c.column
                return cell.row, headers
    return None, {}


def pick(headers, *words):
    for h, col in headers.items():
        if all(w in h for w in words):
            return col
    return None


def is_month_sheet(ws):
    return bool(re.search(r"\d+\s*月|月", ws.title)) or bool(re.match(r"^\d{4}[-_/]?\d{1,2}$", ws.title))


# ------------------------------------------------------------
# 1 シートぶんの流し込み
# ------------------------------------------------------------
def fill_sheet(ws, result, stats):
    hrow, headers = find_header(ws)
    if not hrow:
        return False
    col_order = pick(headers, "受注番号") or pick(headers, "管理番号")
    col_size = pick(headers, "仕上")
    col_remark = pick(headers, "備考")
    last_col = max(headers.values())
    last_row = ws.max_row
    while last_row > hrow and all(ws.cell(last_row, c).value in (None, "") for c in range(1, last_col + 1)):
        last_row -= 1

    # 足す列（既に足してあれば同じ列を使う）
    append_at = {}
    for key, title in APPEND_COLS:
        existing = pick(headers, title.split("\n")[0], "koutei")
        if existing:
            append_at[key] = existing
        else:
            last_col += 1
            append_at[key] = last_col
            ws.cell(hrow, last_col, title)
    if col_size is None:
        last_col += 1
        col_size = last_col
        ws.cell(hrow, col_size, "仕上りサイズ")
    if col_remark is None:
        last_col += 1
        col_remark = last_col
        ws.cell(hrow, col_remark, "備考")

    st = stats.setdefault(ws.title, {"rows": 0, "hit": 0, "size_filled": 0, "remark_filled": 0, "hrow": hrow,
                                     "first": hrow + 1, "last": last_row, "col_a3": append_at["A3以下"],
                                     "col_kind": append_at["内外作区分"], "col_size": col_size})
    for r in range(hrow + 1, last_row + 1):
        onum = ws.cell(r, col_order).value
        key, _ = norm_order(onum)
        if not key:
            continue
        st["rows"] += 1
        rec = result.get(key)
        if not rec:
            continue
        st["hit"] += 1
        if ws.cell(r, col_size).value in (None, "") and rec["仕上りサイズ"]:
            ws.cell(r, col_size, rec["仕上りサイズ"]).fill = FILL_FILLED
            st["size_filled"] += 1
        draft = remark_draft(rec)
        if ws.cell(r, col_remark).value in (None, "") and draft:
            ws.cell(r, col_remark, draft).fill = FILL_FILLED
            st["remark_filled"] += 1
        for key2, col in append_at.items():
            ws.cell(r, col, rec.get(key2, ""))
    st["last_col"] = last_col
    return True


# ------------------------------------------------------------
# 見た目
# ------------------------------------------------------------
def style_sheet(ws, st, title_text):
    hrow, last_row, last_col = st["hrow"], st["last"], st["last_col"]
    # 見出し
    for c in range(1, last_col + 1):
        cell = ws.cell(hrow, c)
        is_k = "koutei" in str(cell.value or "")
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER_K if is_k else FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER
    ws.row_dimensions[hrow].height = 30
    # 明細
    for r in range(hrow + 1, last_row + 1):
        band = (r - hrow) % 2 == 0
        for c in range(1, last_col + 1):
            cell = ws.cell(r, c)
            cell.font = FONT
            cell.border = BORDER
            if cell.alignment is None or not cell.alignment.wrap_text:
                cell.alignment = ALIGN_WRAP
            if band and cell.fill.fgColor.rgb in (None, "00000000") and cell.fill.fill_type is None:
                cell.fill = FILL_BAND
        a3 = ws.cell(r, st["col_a3"])
        if a3.value in FILL_A3:
            a3.fill = FILL_A3[a3.value]
            a3.alignment = ALIGN_CENTER
    # 列幅（中身の長さから。全角を 2 と数える）
    for c in range(1, last_col + 1):
        letter = get_column_letter(c)
        if ws.column_dimensions[letter].width and c <= st.get("orig_last_col", 0):
            continue  # 元の Excel で決めてある幅は尊重する
        longest = 4
        for r in range(hrow, min(last_row, hrow + 300) + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            for line in str(v).split("\n"):
                longest = max(longest, sum(2 if ord(ch) > 0x7F else 1 for ch in line))
        ws.column_dimensions[letter].width = min(max(longest + 2, 6), 40)
    # 固定・フィルタ・印刷
    ws.freeze_panes = ws.cell(hrow + 1, 2)
    ws.auto_filter.ref = f"A{hrow}:{get_column_letter(last_col)}{last_row}"
    ws.print_title_rows = f"{hrow}:{hrow}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddFooter.center.text = "&P / &N"
    # 表題（見出しの上に空きがあれば）
    if hrow >= 2 and all(ws.cell(1, c).value in (None, "") for c in range(1, last_col + 1)):
        ws.cell(1, 1, title_text).font = FONT_TITLE
        if hrow >= 3 and all(ws.cell(2, c).value in (None, "") for c in range(1, last_col + 1)):
            ws.cell(2, 1, "水色の欄は koutei-kanr30 から自動で埋めたもの（手入力の欄は触っていない）。"
                          "A3以下: 緑=○（A3以下）／橙=×（A3より大きい）／黄=不明").font = FONT_NOTE


# ------------------------------------------------------------
# 集計シート
# ------------------------------------------------------------
def add_summary(wb, stats, result_path, base_path):
    name = "集計（koutei）"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws.cell(1, 1, "平版印刷 → オンデマンド移行検討　koutei-kanr30 反映の集計").font = FONT_TITLE
    ws.cell(2, 1, f"作成: {dt.date.today().isoformat()}　元: {os.path.basename(base_path or '')}　"
                  f"取得結果: {os.path.basename(result_path)}").font = FONT_NOTE
    ws.cell(3, 1, "件数は各月シートを COUNTIF で数えている（月シートを直せば追従する）").font = FONT_NOTE
    heads = ["月（シート）", "行数", "koutei該当", "仕上りサイズを埋めた", "備考を埋めた",
             "A3以下 ○", "A3以下 ×", "A3以下 不明", "内作", "外注", "内作・外注"]
    hrow = 5
    for i, h in enumerate(heads, 1):
        c = ws.cell(hrow, i, h)
        c.font, c.fill, c.alignment, c.border = FONT_HEADER, FILL_HEADER, ALIGN_CENTER, BORDER
    r = hrow + 1
    for title, st in stats.items():
        q = f"'{title.replace(chr(39), chr(39) * 2)}'"
        a3 = f"{q}!{get_column_letter(st['col_a3'])}{st['first']}:{get_column_letter(st['col_a3'])}{st['last']}"
        kd = f"{q}!{get_column_letter(st['col_kind'])}{st['first']}:{get_column_letter(st['col_kind'])}{st['last']}"
        vals = [title, st["rows"], st["hit"], st["size_filled"], st["remark_filled"],
                f'=COUNTIF({a3},"○")', f'=COUNTIF({a3},"×")', f'=COUNTIF({a3},"不明")',
                f'=COUNTIF({kd},"内作")', f'=COUNTIF({kd},"外注")', f'=COUNTIF({kd},"内作・外注")']
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v)
            c.font, c.border = FONT, BORDER
        r += 1
    if stats:
        ws.cell(r, 1, "合計").font = FONT_BOLD
        for i in range(2, len(heads) + 1):
            col = get_column_letter(i)
            c = ws.cell(r, i, f"=SUM({col}{hrow + 1}:{col}{r - 1})")
            c.font, c.border = FONT_BOLD, BORDER
        ws.cell(r, 1).border = BORDER
    r += 2
    ws.cell(r, 1, "凡例").font = FONT_BOLD
    legend = [("水色の欄", FILL_FILLED, "koutei-kanr30 の値で自動で埋めた欄（元から入っていた欄は触っていない）"),
              ("○", FILL_A3["○"], "仕上りサイズが A3 以下（オンデマンド機に載る大きさ）"),
              ("×", FILL_A3["×"], "1 部品でも A3 より大きい"),
              ("不明", FILL_A3["不明"], "仕上りサイズが無い、または「規格外」だけで実寸が無い")]
    for label, fill, text in legend:
        r += 1
        c = ws.cell(r, 1, label)
        c.fill, c.font, c.border, c.alignment = fill, FONT, BORDER, ALIGN_CENTER
        ws.cell(r, 2, text).font = FONT
    r += 2
    notes = ["注意",
             "・「行数」「koutei該当」「埋めた」の 3 列はこのファイルを作った時点の数（式ではない）",
             "・仕上りサイズが「B1」だけの案件は基幹システムの一括取込の既定値の疑いがある。取得結果の _詳細.csv で作成日時を確かめる",
             "・加工所・加工日程・台数は koutei-kanr30 に無い。外注委託依頼書を出した案件と内作加工を登録した案件だけ加工内容が入る",
             "・委託先名の表記ゆれ（㈱松岡製本／松岡製本（株））はそのまま。名寄せはしていない"]
    for i, t in enumerate(notes):
        ws.cell(r + i, 1, t).font = FONT_BOLD if i == 0 else FONT_NOTE
    for i, w in enumerate([22, 8, 11, 16, 11, 10, 10, 11, 8, 8, 11], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(hrow + 1, 2)


# ------------------------------------------------------------
# base 無し: 結果 CSV だけから一覧を作る
# ------------------------------------------------------------
def build_from_result(result):
    wb = Workbook()
    ws = wb.active
    ws.title = "一覧"
    hrow = 4
    for i, h in enumerate(RESULT_COLS, 1):
        ws.cell(hrow, i, h)
    r = hrow
    for rec in result.values():
        r += 1
        for i, h in enumerate(RESULT_COLS, 1):
            ws.cell(r, i, rec.get(h, ""))
    st = {"rows": len(result), "hit": len(result), "size_filled": 0, "remark_filled": 0, "hrow": hrow, "first": hrow + 1,
          "last": r, "col_a3": 3, "col_kind": 5, "col_size": 2, "last_col": len(RESULT_COLS)}
    return wb, {"一覧": st}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--result", required=True, help="export_finish_size.py の出力 CSV")
    ap.add_argument("--base", help="元の Excel（月別シート入り）。省くと一覧だけの Excel を作る")
    ap.add_argument("--out", help="出力ファイル名。省くと <元の名前>_koutei反映.xlsx")
    ap.add_argument("--title", default="平版印刷 → オンデマンド移行検討（通し数 3,000 以下）", help="各シートの表題")
    args = ap.parse_args()

    result = read_result(args.result)
    print(f"取得結果: {len(result)} 件（{args.result}）")

    if args.base:
        if not os.path.exists(args.base):
            sys.exit(f"元の Excel が見つかりません: {args.base}")
        out = args.out or f"{os.path.splitext(args.base)[0]}_koutei反映.xlsx"
        if os.path.abspath(out) == os.path.abspath(args.base):
            sys.exit("出力先が元の Excel と同じです。元のファイルは書き換えません")
        wb = load_workbook(args.base)
        stats = {}
        for ws in wb.worksheets:
            if not is_month_sheet(ws):
                continue
            hrow, headers = find_header(ws)
            orig_last_col = max(headers.values()) if headers else 0
            if fill_sheet(ws, result, stats):
                stats[ws.title]["orig_last_col"] = orig_last_col
                style_sheet(ws, stats[ws.title], f"{args.title}　{ws.title}")
                s = stats[ws.title]
                print(f"  {ws.title}: {s['rows']} 行、koutei 該当 {s['hit']}、仕上りサイズ埋め {s['size_filled']}、備考埋め {s['remark_filled']}")
            else:
                print(f"  {ws.title}: 受注番号の見出しが見つからないので飛ばします")
        if not stats:
            sys.exit("月別シートが 1 つも見つかりませんでした（シート名に「月」、見出しに「受注番号」か「管理番号」が要ります）")
    else:
        out = args.out or f"{os.path.splitext(args.result)[0]}.xlsx"
        wb, stats = build_from_result(result)
        style_sheet(wb["一覧"], stats["一覧"], args.title)

    add_summary(wb, stats, args.result, args.base)
    wb.save(out)
    print(f"出力: {out}")


if __name__ == "__main__":
    main()
