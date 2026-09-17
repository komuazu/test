#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
基幹システム（MIS）の帳票 PDF から取り出した文字を読んで、受注ごとの項目にする。

対象の帳票（Google ドライブ「作業依頼書」「委託業務依頼書」フォルダ）:
  * 【製造指示書】      … 受注番号・得意先・製品名・受注数量・仕上りサイズ・総頁数・色数・用紙・通し数・加工・出荷日
  * 印 刷 指 示 書      … koutei-kanr30 が出す簡易版（仕上サイズ・受注数量・用紙・通し数）
  * 外注委託依頼書      … koutei-kanr30 が出す（委託先・加工内容・仕上がりサイズ・用紙）

PDF の文字は表の枠が無くなって 1 行に潰れているので、見出し語を目印に切り出す。
取れなかった項目は空にし、加工の欄は原文も残す（判断は人が最後にする）。
"""

import re
import unicodedata

DATE = r"\d{4}/\d{1,2}/\d{1,2}"
SIZE_MM = r"\d{2,4}(?:\.\d+)?[×x*]\d{2,4}(?:\.\d+)?"
# 用紙規格（全判の呼び名）
SHEET_SIZES = r"(?:A全判|B全判|菊全判|四六判|菊半裁|A半裁|B半裁|四六半裁|菊四裁|A倍判|B倍判|ハトロン判|新聞判|規格外|[A-Za-z0-9]*判)"
# 加工作業として拾う言葉（長いものを先に）
PROCESS_WORDS = [
    r"大巻[三四]つ折り?", r"小巻[三四]つ折り?", r"巻き?[三四]つ折り?", r"外[三四]つ折り?", r"観音折り?", r"蛇腹折り?", r"DM折り?",
    r"[二三四]つ折り?", r"十字折り?", r"クロス折り?", r"Z折り?", r"折り?加工", r"折り",
    r"中綴じ\d*[Pp]?", r"無線綴じ\d*[Pp]?", r"あじろ綴じ\d*[Pp]?", r"PUR綴じ?\d*[Pp]?", r"平綴じ", r"針金綴じ", r"糸かがり", r"上製本", r"製本",
    r"平化粧断裁", r"化粧断裁", r"断裁", r"断ち", r"ｽﾘｯﾀｰ|スリッター|スリット",
    r"抜きミシン", r"ミシン\(?[^\s)]*\)?", r"筋押し", r"筋入れ", r"スジ入れ", r"筋", r"型抜き", r"抜き", r"角丸", r"穴[あ開]け", r"穴",
    r"PP(?:加工|貼り)?", r"マットPP", r"グロスPP", r"箔押し", r"エンボス", r"ニス", r"ラミネート", r"ラミ",
    r"丁合", r"封入", r"帯掛け", r"帯", r"糊付け", r"糊", r"ポケット貼り", r"貼り", r"シール", r"封筒",
    r"その他外注", r"その他",
]
PROCESS_RE = re.compile("|".join(PROCESS_WORDS))
# 加工所として拾う言葉（社内）
INTERNAL_SITES = ("第二工場", "本社", "POP課", "工務課", "第一工場", "社内")


def z2h(s):
    return unicodedata.normalize("NFKC", s)


def toks(s):
    return s.split()


def is_int(t):
    return bool(re.fullmatch(r"\d{1,3}(?:,\d{3})*|\d+", t))


def to_int(t):
    try:
        return int(t.replace(",", ""))
    except (ValueError, AttributeError):
        return None


def order_from_title(title):
    m = re.match(r"^\D{0,12}?(\d{7,9})", z2h(title or ""))
    return m.group(1) if m else ""


def product_from_title(title):
    """ファイル名の「08728296　障害者採用パンフレット.pdf」から品名部分を取る。"""
    t = re.sub(r"\.pdf$", "", title or "", flags=re.I)
    t = re.sub(r"^(再|外注委託依頼書|印刷指示書)\s*", "", t)
    t = re.sub(r"^\D{0,12}?\d{7,9}\s*[.．・_\-　]*\s*", "", t)
    t = re.sub(r"[_＿][^_＿]*(様|株式会社|㈱|㈲|有限会社)[^_＿]*$", "", t)  # 委託依頼書の「_委託先名」
    t = re.sub(r"[【\[（(][^】\])）]*(指示書|指示|面付|指示書\)|複製\d*)[^】\])）]*[】\])）]", "", t)
    return t.strip(" 　・_-.")


# ------------------------------------------------------------
# 製造指示書
# ------------------------------------------------------------
PRINT_HEADER_WORDS = {"部品", "内・外作", "印刷工場", "印刷機種", "オフ輪折", "色数", "ｽﾘｯﾀｰ", "台数", "全判面付", "通し数",
                      "印刷予備", "加工予備", "コメント(部品)", "備考(色名)"}


def parse_seizo(text, title=""):
    """製造指示書。通常の PDF は表の行ごとに 1 行、Drive が Markdown 風に返したものはセルごとに 1 行になるので、
    全部を 1 行につないでから見出し語で切り出す。"""
    d = {"doc": "製造指示書"}
    t = text.replace("\r", "").replace("　", " ")
    flat = re.sub(r"\s+", " ", t)
    m = re.search(r"発行日 (\S+ \S+) 【製造指示書】", flat)
    d["発行日"] = m.group(1) if m else ""

    # 受注番号 得意先 製品名 所属 担当
    m = re.search(r"金額区分 確定区分 (\d{8}) (.+?) (\d{6}) (\S+) ", flat)
    if m:
        d["受注番号"] = m.group(1)
        seg = m.group(2)
        d["所属"] = m.group(4)
        prod = product_from_title(title)
        if prod and seg.endswith(prod):
            d["得意先"] = seg[: -len(prod)].strip()
            d["製品名"] = prod
        else:
            d["得意先"] = seg
            d["製品名"] = prod or seg
            d["得意先・製品名(原文)"] = seg
    else:
        d["受注番号"] = order_from_title(title)
        d["得意先"] = ""
        d["製品名"] = product_from_title(title)

    # 受注数量 実内見本数 実外見本数 仕上りサイズ 総頁数
    m = re.search(r"前回受注番号 (.+?) 受注日 入稿日 下版日", flat)
    if m:
        tk = toks(m.group(1))
        # 先頭に見出し語（受注数量 …色数）が混ざる並びもあるので、最初の数字から
        while tk and not is_int(tk[0]):
            tk.pop(0)
        if len(tk) >= 4 and is_int(tk[0]) and is_int(tk[1]) and is_int(tk[2]):
            d["受注数量"] = to_int(tk[0])
            d["実内見本数"] = to_int(tk[1])
            d["実外見本数"] = to_int(tk[2])
            rest = tk[3:]
            if len(rest) >= 2 and re.fullmatch(r"\d+", rest[-1]):
                d["総頁数"] = to_int(rest[-1])
                rest = rest[:-1]
            d["仕上りサイズ"] = " ".join(rest)
    # 色数と日付（受注日 入稿日 下版日 … 部品 版種 の間にある）
    m = re.search(r"受注日 入稿日 下版日 (.+?) 部品 版種", flat)
    if m:
        tk = toks(m.group(1))
        dates = [x for x in tk if re.fullmatch(DATE, x)]
        colors = [x for x in tk if re.fullmatch(r"\d+\+\d+", x)]
        prev = [x for x in tk if re.fullmatch(r"\d{8}", x)]
        if colors:
            d["色数"] = colors[0]
        if prev and prev[0] != d.get("受注番号"):
            d["前回受注番号"] = prev[0]
        if dates:
            d["受注日"] = dates[0]
        if len(dates) >= 2:
            d["入稿日"] = dates[1]
        d["_dates"] = dates

    # 出荷日・納品日: 加工欄より後ろにある最初の「日付 日付」の組。無ければ受注日欄の後ろ 2 つ
    pos = flat.find("部品 内・外作 加工所")
    tail = flat[pos:] if pos >= 0 else flat
    m = re.search(rf"({DATE}) ({DATE})", tail)
    if m:
        d["出荷日"], d["納品日"] = m.group(1), m.group(2)
    elif len(d.get("_dates", [])) >= 4:
        d["出荷日"], d["納品日"] = d["_dates"][-2], d["_dates"][-1]
    d.pop("_dates", None)

    # 用紙
    papers = []
    seg = re.search(r"加工予備枚数 備考 (.*?) 部品 内・外作 印刷工場", flat)
    if seg:
        s = seg.group(1)
        for m in re.finditer(rf"(?:([^\d\s,]+\d) )?(当方|先方|支給) (.+?) (\S*判|規格外) ({SIZE_MM})(?: (\d+\.\d+|\d+))?", s):
            part, side, brand, size, mm, ream = m.groups()
            grain = ""
            brand_toks = toks(brand)
            while brand_toks and re.fullmatch(r"平|巻|[YT]目?|[\d.]+k|\S*判|規格外", brand_toks[-1]):
                x = brand_toks.pop()
                if re.fullmatch(r"[YT]目?", x):
                    grain = x[0]
            papers.append({"部品": part or "", "当先": side, "銘柄": " ".join(brand_toks), "規格": size,
                           "寸法": mm, "連量": ream or "", "目": grain})
        # Markdown 風では連量が後ろにまとめて出るので、無い行に順に当てる
        missing = [p for p in papers if not p["連量"]]
        if missing:
            after = s[s.rfind(papers[-1]["寸法"]) + len(papers[-1]["寸法"]):] if papers else ""
            reams = re.findall(r"(?<![\d,])(\d{2,3}\.\d)(?![\d])", after)
            for p, r in zip(missing, reams):
                p["連量"] = r
    d["用紙"] = papers

    # 印刷（通し数）: 色数「4＋4」の後ろの 5 つの数字 = 台数 面付 通し数 印刷予備 加工予備
    prints = []
    seg = re.search(r"コメント\(部品\) (.*?) 部品 内・外作 加工所", flat)
    if seg:
        tk = [x for x in toks(seg.group(1)) if x not in PRINT_HEADER_WORDS]
        i = 0
        while i < len(tk):
            if re.fullmatch(r"\d+[＋+]\d+", tk[i]):
                nums = []
                j = i + 1
                while j < len(tk) and len(nums) < 5:
                    if is_int(tk[j]):
                        nums.append(to_int(tk[j]))
                    elif nums:
                        break
                    j += 1
                before = tk[max(0, i - 4):i]
                machine = next((x for x in reversed(before) if re.fullmatch(r"(菊全|A全|B全|菊半|A半|4/6全|四六全|四六|オンデマンド|POD|\S*UV\S*|\S*機)", x)), "")
                inout = next((x for x in before if x in ("内作", "外作")), "")
                part = next((x for x in before if re.fullmatch(r"[^\d\s,]+\d", x) and x != machine), "")
                row = {"部品": part, "内外": inout, "機種": machine, "色数": tk[i]}
                for k, v in zip(["台数", "面付", "通し数", "印刷予備", "加工予備"], nums):
                    row[k] = v
                prints.append(row)
                i = j
            else:
                i += 1
    d["印刷"] = prints
    d["通し数合計"] = sum(p.get("通し数") or 0 for p in prints) if prints else None

    # 加工
    proc = {"原文": "", "内外": "", "加工所": [], "作業": [], "加工日程": []}
    seg = re.search(r"部品 内・外作 加工所 加工日程 加工作業(.*?)(?:区分 部数 納品先名|時間 実外 備考|ジャパンプリント株式会社)", flat)
    if seg:
        s = seg.group(1)
        for w in ("台数 仕上サイズ", "展開サイズ 数量", "台数", "仕上サイズ", "展開サイズ", "数量", "区分 部数 納品先名 住所1 出荷日 納品日"):
            s = s.replace(w, " ")
        s = re.sub(r"\s+", " ", s).strip()
        proc["原文"] = s
        tk = toks(s)
        for i, x in enumerate(tk):
            if x in ("外作", "内作", "外注"):
                nxt = tk[i + 1] if i + 1 < len(tk) else ""
                if nxt and not re.fullmatch(DATE, nxt) and not is_int(nxt) and not PROCESS_RE.fullmatch(nxt) \
                        and not re.fullmatch(r"[^\d\s,]+\d", nxt) and nxt not in ("梱包", "納品", "配送"):
                    proc["加工所"].append(nxt)
            elif x in INTERNAL_SITES:
                proc["加工所"].append(x)
            elif re.fullmatch(DATE, x):
                proc["加工日程"].append(x)
        proc["加工所"] = list(dict.fromkeys(proc["加工所"]))
        s2 = s
        for site in proc["加工所"]:
            s2 = s2.replace(site, " ")
        for m in PROCESS_RE.finditer(s2):
            w = m.group(0)
            if w not in proc["作業"] and w != "その他":
                proc["作業"].append(w)
        ext = [x for x in proc["加工所"] if x not in INTERNAL_SITES]
        inn = [x for x in proc["加工所"] if x in INTERNAL_SITES]
        if "外作" in tk or "外注" in tk or ext or "その他外注" in proc["作業"]:
            proc["内外"] = "外注"
        if "内作" in tk or inn:
            proc["内外"] = "内作・外注" if proc["内外"] == "外注" else "内作"
    d["加工"] = proc
    return d


# ------------------------------------------------------------
# 印刷指示書（koutei）
# ------------------------------------------------------------
def parse_insatsu(text, title=""):
    d = {"doc": "印刷指示書"}
    t = text.replace("\r", "")
    m = re.search(r"受注番号 (\d{8})", t)
    d["受注番号"] = m.group(1) if m else order_from_title(title)
    m = re.search(r"品名 (.+?) 得意先 (.+?) 担当/所属", t)
    if m:
        d["製品名"], d["得意先"] = m.group(1).strip(), m.group(2).strip()
    m = re.search(r"仕上サイズ (.+?) 受注数量 ([\d,]+)", t)
    if m:
        d["仕上りサイズ"], d["受注数量"] = m.group(1).replace("\\*", "×").replace("*", "×").strip(), to_int(m.group(2))
    m = re.search(r"表色数 (\d+) 裏色数 (\d+)", t)
    if m:
        d["色数"] = f"{m.group(1)}+{m.group(2)}"
    m = re.search(r"印刷機種 (\S+)", t)
    d["機種"] = m.group(1) if m else ""
    m = re.search(r"用紙品質 (.+?) 厚さ\(斤量\) (\S+) 用紙規格 (\S+)(?: 紙目 (\S+))?", t)
    papers = []
    if m:
        papers.append({"部品": "", "当先": "", "銘柄": m.group(1).strip(), "規格": m.group(3), "寸法": "",
                       "連量": m.group(2).replace("kg", ""), "目": (m.group(4) or "")[:1]})
    d["用紙"] = papers
    m = re.search(r"合計 ([\d,]+) ([\d,]+) ([\d,]+) ([\d,]+)", t)
    d["通し数合計"] = to_int(m.group(1)) if m else None
    d["印刷"] = []
    m = re.search(r"最終納品日 (\S+) 発送日 (\S+)", t)
    if m:
        d["納品日"] = m.group(1).replace("-", "/")
        d["出荷日"] = m.group(2).replace("-", "/")
    d["加工"] = {"原文": "", "内外": "", "加工所": [], "作業": [], "加工日程": []}
    return d


# ------------------------------------------------------------
# 外注委託依頼書（koutei）
# ------------------------------------------------------------
def parse_itaku(text, title=""):
    d = {"doc": "外注委託依頼書"}
    t = text.replace("\r", "")
    m = re.search(r"【 委託先 】\s*\n\s*(.+?)\s*\n", t)
    d["委託先"] = m.group(1).strip() if m else ""
    m = re.search(r"受注番号 (\d{7,9})", t)
    d["受注番号"] = m.group(1) if m else order_from_title(title)
    m = re.search(r"客先名 (.+)", t)
    d["得意先"] = m.group(1).strip() if m else ""
    m = re.search(r"品名 (.+)", t)
    d["製品名"] = m.group(1).strip() if m else product_from_title(title)
    m = re.search(r"数量 ([\d,]+)", t)
    d["受注数量"] = to_int(m.group(1)) if m else None
    m = re.search(r"仕上がりサイズ (.+)", t)
    d["仕上りサイズ"] = m.group(1).replace("\\*", "×").replace("*", "×").strip() if m else ""
    papers = []
    for m in re.finditer(r"用紙（([^）]*)） (.+)", t):
        tk = toks(m.group(2))
        # 例: OKマットポスト 菊全判 104 Y
        brand, size, ream, grain = tk[0] if tk else "", "", "", ""
        for x in tk[1:]:
            if re.fullmatch(SHEET_SIZES, x):
                size = x
            elif re.fullmatch(r"[\d.]+k?g?", x):
                ream = x
            elif re.fullmatch(r"[YT]目?", x):
                grain = x[0]
            else:
                brand += " " + x
        papers.append({"部品": m.group(1), "当先": "", "銘柄": brand.strip(), "規格": size, "寸法": "", "連量": ream, "目": grain})
    d["用紙"] = papers
    m = re.search(r"\n加工内容 (.+)", t)
    content = m.group(1).strip() if m else ""
    m = re.search(r"備考【加工内容】\s*\n\s*(.+)", t)
    if m and m.group(1).strip() and m.group(1).strip() != content:
        content = f"{content} / {m.group(1).strip()}" if content else m.group(1).strip()
    m = re.search(r"外注納品日 (\S+)", t)
    d["外注納品日"] = m.group(1) if m else ""
    d["加工"] = {"原文": content, "内外": "外注", "加工所": [d["委託先"]] if d["委託先"] else [], "作業": [content] if content else [], "加工日程": []}
    d["印刷"] = []
    d["通し数合計"] = None
    return d


def parse_any(text, title=""):
    if not text:
        return None
    if "【製造指示書】" in text:
        return parse_seizo(text, title)
    if text.lstrip().startswith("外注委託依頼書"):
        return parse_itaku(text, title)
    if re.match(r"\s*印\s*刷\s*指\s*示\s*書", text):
        return parse_insatsu(text, title)
    return {"doc": "その他", "受注番号": order_from_title(title), "製品名": product_from_title(title), "用紙": [], "印刷": [],
            "加工": {"原文": "", "内外": "", "加工所": [], "作業": [], "加工日程": []}, "通し数合計": None}
