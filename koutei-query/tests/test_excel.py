#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_excel.py の動作確認。DB は使わない。

元 Excel のダミー（月別シート 2 枚＋サマリー、46 件相当の記入済み行を含む）と結果 CSV を作り、
流し込んだ結果を openpyxl で読み直して確かめる。LibreOffice があれば式の再計算も確かめる。

    python tests/test_excel.py
"""
import csv
import glob
import json
import os
import subprocess
import sys
import tempfile

from openpyxl import Workbook, load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "..", "make_excel.py")

RESULT_ROWS = [
    # 受注番号, 仕上りサイズ, A3以下, 加工内容, 内外作区分, 委託先名, 用紙銘柄, 用紙規格, 斤量
    ["8726258", "A4 297×210", "○", "中綴じ12P / ミシン(筋)", "内作・外注", "八王子紙工", "A2マット", "A全判", "86.5 / 110"],
    ["8726259", "B2 728×515", "×", "折り", "", "", "オーロラコート", "菊全判", "93.5"],
    ["8726260", "A4", "○", "抜き・ポケット貼り・24P中綴じ", "外注", "松岡製本", "", "", ""],
    ["8726261", "", "不明", "", "", "", "", "", ""],
    ["8726262", "B5", "○", "折加工（二つ折り） / 二つ折り / 折加工1", "内作", "", "", "", ""],
]


def make_base(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "サマリー"
    ws["A1"] = "サマリー（元のまま残ること）"
    ws["A2"] = "=COUNTA('2026年4月'!B:B)"
    heads = ["No", "受注番号（管理番号）", "得意先", "品名", "通し数", "仕上りサイズ", "備考"]
    data = {
        "2026年4月": [
            [1, 8726258, "㈱テスト", "ﾊﾟﾝﾌﾚｯﾄ", 2400, "", ""],
            [2, 8726259, "㈱テスト", "ﾎﾟｽﾀｰ", 800, "B2", "手入力済み（触らない）"],   # 記入済み
            [3, 8726261, "㈱テスト", "該当なし", 500, "", ""],
        ],
        "2026年5月": [
            [1, "08726260", "㈱テスト", "会社案内", 1200, "", ""],                 # ゼロ埋め表記
            [2, 8726262.0, "㈱テスト", "ﾁﾗｼ", 3000, "", ""],                     # 小数表記
            [3, 8799999, "㈱テスト", "koutei に無い", 100, "", ""],
        ],
    }
    for title, rows in data.items():
        s = wb.create_sheet(title)
        s["A1"] = f"{title} 平版印刷実績"
        for i, h in enumerate(heads, 1):
            s.cell(3, i, h)
        for r, row in enumerate(rows, 4):
            for i, v in enumerate(row, 1):
                s.cell(r, i, v)
    wb.save(path)


def main():
    with tempfile.TemporaryDirectory() as td:
        base = os.path.join(td, "平版印刷_オンデマンド移行検討_3000通し以下一覧.xlsx")
        make_base(base)
        res = os.path.join(td, "仕上りサイズ_koutei取得結果.csv")
        with open(res, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, lineterminator="\r\n")
            w.writerow(["受注番号", "仕上りサイズ", "A3以下", "加工内容", "内外作区分", "委託先名", "用紙銘柄", "用紙規格", "斤量"])
            w.writerows(RESULT_ROWS)
        base_before = open(base, "rb").read()

        r = subprocess.run([sys.executable, SCRIPT, "--base", base, "--result", res], capture_output=True, text=True)
        print(r.stdout)
        if r.returncode:
            print(r.stderr)
            sys.exit("make_excel.py が失敗")
        assert open(base, "rb").read() == base_before, "元の Excel が書き換わっている"
        out = os.path.join(td, "平版印刷_オンデマンド移行検討_3000通し以下一覧_koutei反映.xlsx")
        assert os.path.exists(out), "出力ファイルが無い"

        wb = load_workbook(out)
        assert wb.sheetnames == ["サマリー", "2026年4月", "2026年5月", "集計（koutei）"], wb.sheetnames
        assert wb["サマリー"]["A2"].value == "=COUNTA('2026年4月'!B:B)", "元のサマリーの式が変わった"

        s4 = wb["2026年4月"]
        heads = [c.value for c in s4[3]]
        assert heads[:7] == ["No", "受注番号（管理番号）", "得意先", "品名", "通し数", "仕上りサイズ", "備考"], heads
        assert "A3以下\n(koutei判定)" in heads and "仕上りサイズ\n(koutei)" in heads, heads
        col = {h: i + 1 for i, h in enumerate(heads)}
        # 空欄は埋まる、記入済みは残る
        assert s4.cell(4, col["仕上りサイズ"]).value == "A4 297×210"
        assert s4.cell(4, col["備考"]).value == "内作・外注（八王子紙工）：中綴じ12P / ミシン(筋)", s4.cell(4, col["備考"]).value
        assert s4.cell(4, col["仕上りサイズ"]).fill.fgColor.rgb.endswith("DDEBF7"), "埋めた欄が水色でない"
        assert s4.cell(5, col["仕上りサイズ"]).value == "B2" and s4.cell(5, col["備考"]).value == "手入力済み（触らない）"
        assert not s4.cell(5, col["仕上りサイズ"]).fill.fgColor.rgb.endswith("DDEBF7"), "記入済みの欄を塗っている"
        assert s4.cell(5, col["A3以下\n(koutei判定)"]).value == "×"
        assert s4.cell(6, col["仕上りサイズ"]).value in (None, "") and s4.cell(6, col["A3以下\n(koutei判定)"]).value == "不明"
        assert s4.freeze_panes == "B4" and s4.auto_filter.ref.startswith("A3:"), (s4.freeze_panes, s4.auto_filter.ref)
        assert s4["A1"].value == "2026年4月 平版印刷実績", "元の表題が消えた"

        s5 = wb["2026年5月"]
        assert s5.cell(4, col["仕上りサイズ"]).value == "A4" and s5.cell(4, col["備考"]).value == "外注（松岡製本）：抜き・ポケット貼り・24P中綴じ"
        assert s5.cell(5, col["仕上りサイズ"]).value == "B5" and s5.cell(5, col["備考"]).value == "内作：折加工（二つ折り） / 二つ折り / 折加工1"
        assert s5.cell(6, col["仕上りサイズ"]).value in (None, "") and s5.cell(6, col["A3以下\n(koutei判定)"]).value in (None, "")

        sm = wb["集計（koutei）"]
        rows = {sm.cell(r, 1).value: [sm.cell(r, c).value for c in range(2, 12)] for r in range(6, 9)}
        assert rows["2026年4月"][:4] == [3, 3, 1, 1], rows["2026年4月"]
        assert rows["2026年5月"][:4] == [3, 2, 2, 2], rows["2026年5月"]
        assert rows["2026年4月"][4] == "=COUNTIF('2026年4月'!H4:H6,\"○\")", rows["2026年4月"][4]
        assert rows["合計"][0] == "=SUM(B6:B7)", rows["合計"]
        print("OK: 流し込み・記入済み保護・列追加・集計シート")

        # LibreOffice があれば式を再計算して確かめる
        recalc = glob.glob("/root/.claude/skills/synced/*/xlsx/scripts/recalc.py")
        if recalc:
            r = subprocess.run([sys.executable, recalc[0], out, "240"], capture_output=True, text=True)
            try:
                info = json.loads(r.stdout)
            except ValueError:
                info = {"error": r.stdout + r.stderr}
            print("recalc:", info)
            assert info.get("status") == "success" and info.get("total_errors") == 0, info
            wb2 = load_workbook(out, data_only=True)
            sm2 = wb2["集計（koutei）"]
            vals = {sm2.cell(r, 1).value: [sm2.cell(r, c).value for c in range(6, 12)] for r in range(6, 9)}
            assert vals["2026年4月"] == [1, 1, 1, 0, 0, 1], vals["2026年4月"]
            assert vals["2026年5月"] == [2, 0, 0, 1, 1, 0], vals["2026年5月"]
            assert vals["合計"] == [3, 1, 1, 1, 1, 1], vals["合計"]
            print("OK: 集計の式が正しく計算される")
        else:
            print("recalc.py が無いので式の再計算は飛ばした")

        # base 無し
        r = subprocess.run([sys.executable, SCRIPT, "--result", res, "--out", os.path.join(td, "one.xlsx")],
                           capture_output=True, text=True)
        assert r.returncode == 0, r.stderr
        wb3 = load_workbook(os.path.join(td, "one.xlsx"))
        assert wb3.sheetnames == ["一覧", "集計（koutei）"] and wb3["一覧"].max_row == 4 + len(RESULT_ROWS)
        print("OK: base 無しの一覧")
    print("\nALL OK")


if __name__ == "__main__":
    main()
