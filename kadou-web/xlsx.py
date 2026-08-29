# -*- coding: utf-8 -*-
"""Excel ファイル(.xlsx)を作る

画面の表を、そのままExcelで開ける形で書き出すための小さな書き出し器。
追加のライブラリは使わない（標準の zipfile だけで .xlsx を組み立てる）。

  ・1行目を見出しとして色と罫線を付け、上で固定してオートフィルタを付ける
  ・数値は数値のまま入れて 3桁区切りで表示する（文字列にしない）
  ・列幅は中身の長さに合わせる（全角は2文字ぶんで数える）
"""
import zipfile
from io import BytesIO

MAXW, MINW = 48, 8


def esc(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
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
    return isinstance(v, (int, float)) and not isinstance(v, bool)


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


def build(rows, sheet_name='集計'):
    """表（1行目が見出し）から .xlsx のバイト列を作る"""
    rows = [list(r) for r in rows] or [['']]
    ncol = max(len(r) for r in rows)
    nrow = len(rows)
    last = '%s%d' % (col_name(ncol - 1), nrow)

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
        '<dimension ref="A1:%s"/>'
        '<sheetViews><sheetView workbookViewId="0" tabSelected="1">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '</sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols>%s</cols><sheetData>%s</sheetData>'
        '<autoFilter ref="A1:%s"/>'
        '</worksheet>' % (last, cols, sheet_xml(rows, ncol), last))

    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="Yu Gothic UI"/></font>'
        '<font><b/><sz val="11"/><name val="Yu Gothic UI"/></font>'
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
        '</styleSheet>')

    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="%s" sheetId="1" r:id="rId1"/></sheets></workbook>'
        % esc(sheet_name[:31] or '集計'))

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
                   '</Types>')
        z.writestr('_rels/.rels',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                   '2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                   'officeDocument/2006/relationships/officeDocument"'
                   ' Target="xl/workbook.xml"/></Relationships>')
        z.writestr('xl/_rels/workbook.xml.rels',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                   '2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                   'officeDocument/2006/relationships/worksheet"'
                   ' Target="worksheets/sheet1.xml"/>'
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
                   'officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                   '</Relationships>')
        z.writestr('xl/workbook.xml', workbook)
        z.writestr('xl/styles.xml', styles)
        z.writestr('xl/worksheets/sheet1.xml', sheet)
    return buf.getvalue()
