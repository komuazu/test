# -*- coding: utf-8 -*-
"""Excel ファイル(.xlsx)を作る

画面の表を、そのままExcelで開ける形で書き出すための小さな書き出し器。
追加のライブラリは使わない（標準の zipfile だけで .xlsx を組み立てる）。

  ・1行目を見出しとして色と罫線を付け、上で固定してオートフィルタを付ける
  ・数値は数値のまま入れて 3桁区切りで表示する（文字列にしない）
  ・列幅は中身の長さに合わせる（全角は2文字ぶんで数える）
  ・印刷の設定（A4横・横1ページ・見出し行の繰り返し・余白）まで入れる

印刷まわりについて（2026-09）
    以前は印刷の設定を1つも書いていなかったため、Excelが自前の既定で
    組版することになり、開けるのに印刷でつまずく、という状態だった。
    いまは Excel が自分で保存するときと同じ要素（sheetPr / printOptions /
    pageMargins / pageSetup / 印刷タイトル）をそろえて書いている。
    - 見出し行は各ページの先頭で繰り返す（印刷タイトル = 1行目）
    - 繰り返す行が印刷範囲の中にあると1ページ目で二重に出るので、
      印刷範囲は2行目から
    - フォントは日本語グリフを持つ Meiryo。欧文専用フォントや、この環境に
      無いフォントを指定すると、印刷時の差し替えでExcelが不安定になる
    - ヘッダーの中のフォント指定は &"名前,字体" と字体まで書く。字体を省くと
      Excel はここを解釈できない。ヘッダーは開いたときではなく印刷するときに
      読まれるので、開けるのに印刷だけ失敗する形になる
"""
import datetime as _dt
import math
import re
import zipfile
from io import BytesIO

from xl_theme import THEME1

NOW = _dt.datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')
MAXW, MINW = 48, 8
FONT = 'Meiryo'
# ヘッダー／フッターの中でフォントを指定する書き方は &"フォント名,字体" で、
# 字体まで書くのが決まり（Excel 自身もそう書く）。字体を省くと Excel は
# ここを解釈できず、開けるのに印刷でつまずく。ヘッダーは開いたときではなく
# 印刷するときに読まれるので、印刷だけが失敗する形になる。
FONT_HF = FONT + ',Regular'

# XML 1.0 で使えない制御文字（これが混ざると Excel はファイルを開けない）
CTRL = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]')
# シート名に使えない文字（区切りの / \ は - に、ほかは落とす）
SLASH = re.compile(r"[/\\]")
BADSHEET = re.compile(r"[\[\]:*?]")


def esc(s):
    s = CTRL.sub('', str(s))
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;'))


def col_name(i):
    """0 → A、25 → Z、26 → AA"""
    s = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def width_of(v):
    """全角を2文字ぶんとして数えた表示幅"""
    n = 0
    for ch in str(v):
        n += 2 if ord(ch) > 0x2000 else 1
    return n


def is_num(v):
    """数値として書けるか（nan・inf は Excel が読めないので文字として書く）"""
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return False
    return not isinstance(v, float) or math.isfinite(v)


def sheet_title(name):
    """Excel が受け付けるシート名にする（使えない文字と長さの制限）"""
    s = CTRL.sub('', str(name))
    s = BADSHEET.sub('', SLASH.sub('-', s)).strip().strip("'")[:31]
    return s or '集計'


def quoted(name):
    """数式・定義名の中で使うシート名  例: 'と'を含む名前 → '''を含む名前'"""
    return "'%s'" % esc(name).replace("'", "''")


def sheet_xml(rows, ncol):
    out = []
    for r, row in enumerate(rows, start=1):
        cells = []
        for c in range(ncol):
            v = row[c] if c < len(row) else None
            if v is None or v == '':
                continue
            ref = '%s%d' % (col_name(c), r)
            if r == 1:                                   # 見出し
                cells.append('<c r="%s" s="1" t="inlineStr"><is><t>%s</t></is></c>'
                             % (ref, esc(v)))
            elif is_num(v):
                cells.append('<c r="%s" s="3"><v>%s</v></c>' % (ref, v))
            else:
                cells.append('<c r="%s" s="2" t="inlineStr"><is><t>%s</t></is></c>'
                             % (ref, esc(v)))
        out.append('<row r="%d">%s</row>' % (r, ''.join(cells)))
    return ''.join(out)


def header_text(s):
    """ヘッダーに入れる文字列を用意する

    ヘッダーの中では & が命令の合図（&P はページ番号など）なので、
    文字としての & は && と重ねて書く。そのうえで XML として書き出す。
    """
    return esc(str(s).replace('&', '&&'))


def build(rows, sheet_name='集計'):
    """表（1行目が見出し）から .xlsx のバイト列を作る"""
    rows = [list(r) for r in rows] or [['']]
    ncol = max(len(r) for r in rows)
    nrow = len(rows)
    name = sheet_title(sheet_name)
    lastcol = col_name(ncol - 1)
    last = '%s%d' % (lastcol, nrow)

    widths = []
    for c in range(ncol):
        w = max(width_of(r[c]) for r in rows if c < len(r) and r[c] is not None) \
            if any(c < len(r) and r[c] is not None for r in rows) else MINW
        widths.append(min(max(w + 2, MINW), MAXW))
    cols = ''.join('<col min="%d" max="%d" width="%d" customWidth="1"/>'
                   % (i + 1, i + 1, w) for i, w in enumerate(widths))

    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
        '<dimension ref="A1:%s"/>'
        '<sheetViews><sheetView workbookViewId="0" tabSelected="1">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
        '</sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols>%s</cols><sheetData>%s</sheetData>'
        '<autoFilter ref="A1:%s"/>'
        '<printOptions horizontalCentered="1"/>'
        '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6"'
        ' header="0.3" footer="0.3"/>'
        '<pageSetup paperSize="9" orientation="landscape"'
        ' fitToWidth="1" fitToHeight="0"/>'
        '<headerFooter><oddHeader>&amp;L&amp;"%s"&amp;10 %s'
        '&amp;R&amp;"%s"&amp;9 &amp;P / &amp;N</oddHeader></headerFooter>'
        '</worksheet>'
        % (last, cols, sheet_xml(rows, ncol), last,
           FONT_HF, header_text(name), FONT_HF))

    # 印刷タイトル（見出し行の繰り返し）と印刷範囲。
    # 繰り返す1行目が印刷範囲に入っていると1ページ目で二重に印刷されるため、
    # 印刷範囲は2行目から。明細が無いときは範囲を作らない。
    q = quoted(name)
    names = ['<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'
             '%s!$A$1:$%s$%d</definedName>' % (q, lastcol, nrow)]
    if nrow > 1:
        names.append('<definedName name="_xlnm.Print_Titles" localSheetId="0">'
                     '%s!$1:$1</definedName>' % q)
        names.append('<definedName name="_xlnm.Print_Area" localSheetId="0">'
                     '%s!$A$2:$%s$%d</definedName>' % (q, lastcol, nrow))

    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<workbookPr/>'
        '<bookViews><workbookView activeTab="0"/></bookViews>'
        '<sheets><sheet name="%s" sheetId="1" r:id="rId1"/></sheets>'
        '<definedNames>%s</definedNames>'
        '<calcPr calcId="124519"/>'
        '</workbook>' % (esc(name), ''.join(names)))

    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="%s"/><family val="3"/><charset val="128"/></font>'
        '<font><b/><sz val="11"/><name val="%s"/><family val="3"/>'
        '<charset val="128"/></font>'
        '</fonts>'
        '<fills count="3">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid">'
        '<fgColor rgb="FFE7E9EC"/><bgColor indexed="64"/></patternFill></fill>'
        '</fills>'
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
        '<border><left style="thin"><color rgb="FF9BA1A8"/></left>'
        '<right style="thin"><color rgb="FF9BA1A8"/></right>'
        '<top style="thin"><color rgb="FF9BA1A8"/></top>'
        '<bottom style="thin"><color rgb="FF9BA1A8"/></bottom><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
        '</cellStyleXfs>'
        '<cellXfs count="4">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1"'
        ' applyFill="1" applyBorder="1" applyAlignment="1">'
        '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"'
        ' applyAlignment="1"><alignment vertical="center"/></xf>'
        '<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"'
        ' applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '<dxfs count="0"/>'
        '<tableStyles count="0" defaultTableStyle="TableStyleMedium9"'
        ' defaultPivotStyle="PivotStyleLight16"/>'
        '</styleSheet>' % (FONT, FONT))

    buf = BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
                   'content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxml'
                   'formats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.'
                   'openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/worksheets/sheet1.xml" ContentType='
                   '"application/vnd.openxmlformats-officedocument.spreadsheetml.'
                   'worksheet+xml"/>'
                   '<Override PartName="/xl/styles.xml" ContentType="application/vnd.'
                   'openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                   '<Override PartName="/xl/theme/theme1.xml" ContentType="application/'
                   'vnd.openxmlformats-officedocument.theme+xml"/>'
                   '<Override PartName="/docProps/core.xml" ContentType="application/'
                   'vnd.openxmlformats-package.core-properties+xml"/>'
                   '<Override PartName="/docProps/app.xml" ContentType="application/'
                   'vnd.openxmlformats-officedocument.extended-properties+xml"/>'
                   '</Types>')
        z.writestr('_rels/.rels',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                   '2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                   'officeDocument/2006/relationships/officeDocument"'
                   ' Target="xl/workbook.xml"/>'
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
                   'package/2006/relationships/metadata/core-properties"'
                   ' Target="docProps/core.xml"/>'
                   '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/'
                   'officeDocument/2006/relationships/extended-properties"'
                   ' Target="docProps/app.xml"/>'
                   '</Relationships>')
        z.writestr('xl/_rels/workbook.xml.rels',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                   '2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                   'officeDocument/2006/relationships/worksheet"'
                   ' Target="worksheets/sheet1.xml"/>'
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
                   'officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                   '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/'
                   'officeDocument/2006/relationships/theme"'
                   ' Target="theme/theme1.xml"/>'
                   '</Relationships>')
        # Excel が保存するときと同じ部品をそろえる。テーマ（配色とフォントの
        # 定義）が無いと、開けても印刷（＝描画）でつまずくことがある。
        z.writestr('docProps/core.xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<cp:coreProperties'
                   ' xmlns:cp="http://schemas.openxmlformats.org/package/2006/'
                   'metadata/core-properties"'
                   ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
                   ' xmlns:dcterms="http://purl.org/dc/terms/"'
                   ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                   '<dc:title>%s</dc:title>'
                   '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
                   '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
                   '</cp:coreProperties>' % (esc(name), NOW, NOW))
        z.writestr('docProps/app.xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Properties xmlns="http://schemas.openxmlformats.org/'
                   'officeDocument/2006/extended-properties"'
                   ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/'
                   '2006/docPropsVTypes">'
                   '<Application>Microsoft Excel</Application>'
                   '</Properties>')
        z.writestr('xl/theme/theme1.xml', THEME1)
        z.writestr('xl/workbook.xml', workbook)
        z.writestr('xl/styles.xml', styles)
        z.writestr('xl/worksheets/sheet1.xml', sheet)
    return buf.getvalue()
