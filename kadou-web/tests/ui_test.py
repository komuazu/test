# -*- coding: utf-8 -*-
"""画面の動作確認（Playwright）: 表示・タブ切替・記入欄の保存・合計の再計算・CSV出力"""
import json
import os
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

# ポートは環境変数 KADOU_PORT で変えられる（既定 8765）
URL = 'http://127.0.0.1:%s/' % (os.environ.get('KADOU_PORT') or 8765)
SHOT = Path(sys.argv[1] if len(sys.argv) > 1 else '.')
CHROMIUM = '/opt/pw-browsers/chromium'      # 用意されていればそれを使う
ok, ng = [], []


def check(cond, msg):
    (ok if cond else ng).append(msg)
    print(('  OK  ' if cond else '  NG  ') + msg)


with sync_playwright() as pw:
    b = (pw.chromium.launch(executable_path=CHROMIUM) if Path(CHROMIUM).exists()
         else pw.chromium.launch())
    pg = b.new_page(viewport={'width': 1500, 'height': 1000})
    errors = []
    pg.on('pageerror', lambda e: errors.append(str(e)))
    pg.on('console', lambda m: errors.append(m.text) if m.type == 'error' else None)

    pg.goto(URL)
    pg.wait_for_selector('#matrix tbody tr')

    # ── 年間サマリー ──
    check(pg.locator('#sum tbody tr').count() == 6, '年間集計が5部署+合計行')
    check(pg.locator('#sum tbody tr.total td').first.inner_text().replace(',', '')
          == '1874777', '年間集計の合計が 1,874,777 通し（その他を含む）')
    check(pg.locator('#chart text.lab').count() == 12, '棒グラフが12か月分')
    check(pg.locator('#chart rect.col').count() > 0, '積み上げの棒が出る')
    check(pg.locator('#chart line.grid').count() >= 2, '目盛りの横線が引かれる')
    check(pg.locator('#chart text.tick').count() >= 3, '左に通し数の目盛りが出る')
    check(pg.locator('#chart line.conn').count() > 0,
          '積み上げの段をつなぐ区分線が引かれる  → %d本'
          % pg.locator('#chart line.conn').count())
    check(pg.locator('#chart polyline.rate').count() == 0,
          '積み上げには折れ線を出さない')
    check(pg.locator('#chart text.val').count() > 0, '棒の上に合計が出る')

    # ── 前年との比較（2026年と2025年） ──
    check(pg.locator('#cmpSum tbody tr').count() == 6, '前年との比較が5部署+合計行')
    hdr = pg.locator('#cmpSum thead').inner_text()
    check('前年' in hdr and '増減' in hdr and '前年比' in hdr,
          '前年・増減・前年比の列がある')
    check('2025年' in pg.locator('#cmpNote').inner_text(),
          '比べている年が見出しに出る  → ' + pg.locator('#cmpNote').inner_text())

    # ── グラフを前年比較に切り替える ──
    pg.locator('#pillChart button[data-c=yoy]').click()
    pg.wait_for_timeout(300)
    ms = pg.locator('#chart text.lab').count()
    check(ms and pg.locator('#chart rect.col').count() > ms,
          '月ごとに今年と前年の積み上げが並ぶ  → %d月' % ms)
    check(pg.locator('#chart pattern').count() > 0,
          '前年の棒が斜線になる')
    check(pg.locator('#chart line.conn').count() > 0,
          '今年と前年の段をつなぐ区分線が引かれる  → %d本'
          % pg.locator('#chart line.conn').count())
    check(pg.locator('#chart polyline.rate').count() == 0
          and pg.locator('#chart text.tick2').count() == 0,
          '前年比較のグラフにも折れ線を出さない')
    yl = pg.locator('#chart text.ylab').all_text_contents()
    check(yl[:2] == ['26年', '25年'] and len(yl) == ms * 2,
          '棒の下に年が出る（%d月ぶん×2本）  → %s' % (ms, yl[:4]))
    lg = pg.locator('#legend').inner_text().replace('\n', ' ')
    check('2026' in lg and '2025' in lg and '斜線' in lg,
          '凡例に今年（塗り）と前年（斜線）が出る  → ' + lg)
    check(pg.locator('#selChartDept').is_visible(), '比べる部署を選べる')
    pg.select_option('#selChartDept', '本社営業部')
    pg.wait_for_timeout(300)
    check('本社営業部' in pg.locator('#chartNote').inner_text(),
          '選んだ部署が見出しに出る  → ' + pg.locator('#chartNote').inner_text())
    check(0 < pg.locator('#chart rect.col').count() <= ms * 2,
          '1部署を選ぶと今年と前年の2本だけになる')
    pg.screenshot(path=str(SHOT / '09_グラフ前年比較.png'), full_page=True)
    pg.locator('#pillChart button[data-c=stack]').click()
    pg.wait_for_timeout(300)
    check(pg.locator('#chart rect.col').count() > 0
          and pg.locator('#chart line.conn').count() > 0, '積み上げに戻せる')
    check(pg.locator('#matrix tbody').count() == 6, 'マトリクスが5部署+合計のかたまり')
    kinds = pg.locator('#matrix tbody').first.locator('th.kh').all_inner_texts()
    check(kinds == ['本年（通し数）', '前年', '増減', '前年比', '件数'],
          '1部署が本年・前年・増減・前年比・件数の5段  → %s' % kinds)
    total = pg.locator('#matrix tbody.total tr').first.locator('td.yr').inner_text()
    check('1,874,777' in total, '年間合計 1,874,777 通し  → ' + total.replace('\n', ' '))
    pre = pg.locator('#matrix tbody.total tr').nth(1).locator('td').first.inner_text()
    check(pre == '95,000', '前年の実数が行として出る（合計 1月）  → ' + pre)
    tip = pg.locator('#matrix tbody.total tr').nth(3).locator('td').first.get_attribute('title')
    check('本年' in tip and '前年' in tip,
          '前年比のマスに本年と前年の数字が出る  → ' + tip)
    pg.screenshot(path=str(SHOT / '01_年間サマリー.png'), full_page=True)

    # ── 印刷とPDF ──
    check(pg.locator('#btnPrintTop').is_visible() and pg.locator('#btnPdf').is_visible(),
          '印刷とPDFで保存のボタンがある')
    pg.evaluate('setPrintTitle()')
    check('年間サマリー' in pg.locator('#printTitle').inner_text(),
          '印刷用の見出しが入る  → ' + pg.locator('#printTitle').inner_text())
    pg.emulate_media(media='print')
    check(not pg.locator('header').is_visible()
          and not pg.locator('nav.tabs').is_visible(),
          '印刷ではヘッダーとタブを出さない')
    check(pg.locator('#sum').is_visible() and pg.locator('#printTitle').is_visible(),
          '印刷でも表と見出しは出る')
    pg.emulate_media(media='screen')

    # ── 表を Excel で保存できる ──
    with pg.expect_download(timeout=30000) as dl:
        pg.locator('#btnCsvYear').click()
    got = dl.value
    check(got.suggested_filename.endswith('.xlsx'),
          'Excelファイルとして保存できる  → ' + got.suggested_filename)

    check(pg.locator('#matrix a').count() == 0, 'マトリクスの数字はリンクにしない')

    # ── 月別明細へ ──
    pg.locator('nav.tabs button[data-view=month]').click()
    pg.wait_for_selector('#detail tbody tr')
    check(pg.locator('section#v-month.on').count() == 1, 'タブで月別明細へ移動')

    # ── 1月・全社 ──
    pg.select_option('#selMonth', '1')
    pg.locator('#pillDept button', has_text='全社').click()
    pg.wait_for_selector('#detail tbody tr')
    rows = pg.locator('#detail tbody tr:not(.total):not(.diff)')
    check(rows.count() == 6, '1月・全社は6件に統合  → %d件' % rows.count())
    check('枝番2' in pg.locator('#detail tbody').inner_text(), '枝番統合のバッジが出る')
    check(pg.locator('#detail td.multi').count() == 2, '複数日にまたがる行の日付が黄色（2件）')
    hdr = pg.locator('#detail thead').inner_text().replace('\n', ' / ')
    check('今年の動向' in hdr and '無しの場合の代替対策' in hdr and '対策通し数' in hdr, '記入欄3列がある')
    check('171,777' in pg.locator('#detail tbody tr.total').inner_text(),
          '1月合計 171,777 通し')

    # ── 記入欄に入力 → 合計・差引・充足率が再計算され保存される ──
    first = pg.locator('#detail tbody tr:not(.total):not(.diff)').first
    first.locator('textarea[data-f=trend]').fill('継続（前年並み）')
    first.locator('input[data-f=tsu]').fill('12000')
    first.locator('input[data-f=tsu]').blur()
    pg.wait_for_timeout(900)
    check(pg.locator('#sumPlan').inner_text() == '12,000', '対策通し数の合計が 12,000')
    check(pg.locator('#sumDiff').inner_text() == '159,777',
          '差引が 171,777 − 12,000 = 159,777')
    check(pg.locator('#sumRate').inner_text() == '7.0%', '充足率 7.0%')
    pg.screenshot(path=str(SHOT / '02_月別明細.png'), full_page=True)

    memo = json.loads(pg.evaluate("fetch('/api/memo').then(r=>r.text())"))
    check(any(v.get('trend') == '継続（前年並み）' for v in memo.values()), 'memo.json に保存された')

    # ── 再読込しても入力が残る ──
    pg.reload()
    pg.wait_for_selector('#matrix tbody tr')
    pg.locator('nav.tabs button[data-view=month]').click()
    pg.select_option('#selMonth', '1')
    pg.locator('#pillDept button', has_text='全社').click()
    pg.wait_for_selector('#detail tbody tr')
    kept = pg.locator('#detail tbody tr:not(.total):not(.diff)').first
    check(kept.locator('textarea[data-f=trend]').input_value() == '継続（前年並み）'
          and kept.locator('input[data-f=tsu]').input_value() == '12000', '再読込後も入力が残る')
    check(pg.locator('#sumPlan').inner_text() == '12,000', '再読込後も合計に反映される')

    # ── 部署の分け方が変わっても記入欄は見失わない ──
    moved = pg.evaluate(
        "(function(){ var k = Object.keys(MEMO)[0], p = k.split('|');"
        " var n = [p[0], p[1], 'ちがう部署', p[3]].join('|');"
        " var body = {}; body[k] = null; body[n] = MEMO[k];"
        " return fetch('/api/memo', {method:'POST',"
        " headers:{'Content-Type':'application/json'},"
        " body: JSON.stringify(body)}).then(function(r){return r.ok;}); })()")
    check(moved, '記入欄の部署名を差し替えられた（確かめのため）')
    pg.reload()
    pg.wait_for_selector('#matrix tbody tr')
    pg.locator('nav.tabs button[data-view=month]').click()
    pg.select_option('#selMonth', '1')
    pg.locator('#pillDept button', has_text='全社').click()
    pg.wait_for_selector('#detail tbody tr')
    kept2 = pg.locator('#detail tbody tr:not(.total):not(.diff)').first
    check(kept2.locator('textarea[data-f=trend]').input_value() == '継続（前年並み）',
          '部署の分け方が変わっても記入欄は残る（管理番号で探し直す）')

    # ── 絞り込み ──
    pg.fill('#q', 'アルファ')
    pg.wait_for_timeout(200)
    check(pg.locator('#detail tbody tr:not(.total):not(.diff)').count() == 1, '検索で1件に絞られる')
    pg.fill('#q', '')

    # ── 得意先別 ──
    pg.locator('nav.tabs button[data-view=client]').click()
    pg.wait_for_selector('#clients tbody tr')
    check(pg.locator('#clients tbody tr').count() > 3, '得意先別が表示される')
    ct = pg.locator('#clients tbody').inner_text()
    check('表記2' in ct, '「㈱アルファ商事」と「（株）アルファ商事」を同一得意先として集計')
    check('1,874,777' in pg.locator('#clients tbody tr.total').inner_text(),
          '得意先別の合計が年間合計と一致')
    pg.screenshot(path=str(SHOT / '03_得意先別.png'), full_page=True)

    # ── 得意先別: 月を選ぶと その月のクライアント別になる ──
    pg.select_option('#selMonthC', '1')
    pg.wait_for_timeout(200)
    crows = pg.locator('#clients tbody tr:not(.total)')
    check(crows.count() == 6, '1月の得意先別が6社  → %d社' % crows.count())
    check('品名（管理番号 / 通し数）' in pg.locator('#clients thead').inner_text(),
          '月を選ぶと品名・管理番号の内訳が出る')
    check('171,777' in pg.locator('#clients tbody tr.total').inner_text(),
          '1月の得意先別合計が 171,777 で月合計と一致')
    pg.screenshot(path=str(SHOT / '06_得意先別_月別.png'), full_page=True)
    pg.select_option('#selMonthC', '0')

    # ── 稼働率（有効時間 ÷ 22時間×稼働日数） ──
    pg.locator('nav.tabs button[data-view=oper]').click()
    pg.wait_for_selector('#operSum tbody tr')
    head = pg.locator('#operSum thead').inner_text()
    check('1月' in head and '年間' in head, '機械別 月間稼働率に月と年間の列がある')
    names = pg.locator('#operSum tbody th.rh').all_text_contents()
    check(names == ['稼働日数', 'A全UV', '小森1号機', '菊全UV', '全機械'],
          '稼働日数の行と機械の行が出る  → %s' % names)
    work = pg.locator('#operSum tbody tr').first.locator('td').all_text_contents()
    # 2026年: 1月は元日と成人の日、8月は山の日が休み（土日も除く）
    check(work[0] == '20日' and work[1] == '18日' and work[7] == '20日',
          '祝日と土日を除いた稼働日数が出る（1月20日・2月18日・8月20日）  → %s' % work[:8])

    pg.select_option('#selMonthO', '8')
    pg.wait_for_timeout(200)
    check('稼働 20日 × 22時間 = 440時間' in pg.locator('#operDayNote').inner_text(),
          '分母の内訳が出る  → ' + pg.locator('#operDayNote').inner_text())
    days = pg.locator('#operDay tbody tr')
    check(days.count() == 32, '8月は31日ぶん＋月合計の32行  → %d行' % days.count())
    check(days.nth(10).inner_text().find('山の日') >= 0,
          '祝日は名前で出る（8/11 山の日）')
    off = pg.locator('#operDay tbody tr.off').count()
    check(off == 11, '土日と祝日は灰色にする（8月は10日＋山の日）  → %d日' % off)
    # ダミー日報は 8/9(日) に 8.0h 入れてある。休みの日でも表には出すが分母には入れない
    check('36.4%' in days.nth(8).inner_text(),
          '休みの日に動いた分も出す（8.0h ÷ 22h = 36.4%）')
    check('1.8%' in pg.locator('#operDay tbody tr.total').inner_text(),
          '月合計は 8.0h ÷ 440h = 1.8%')
    pg.locator('#pillOper button[data-o=hours]').click()
    pg.wait_for_timeout(200)
    check('8.00h' in days.nth(8).inner_text(), '有効時間に切り替えられる')
    pg.locator('#pillOper button[data-o=rate]').click()
    pg.screenshot(path=str(SHOT / '10_稼働率.png'), full_page=True)
    with pg.expect_download(timeout=30000) as dl:
        pg.locator('#btnCsvOper').click()
    check(dl.value.suggested_filename.endswith('.xlsx'),
          '稼働率をExcelで保存できる  → ' + dl.value.suggested_filename)
    with pg.expect_download(timeout=30000) as dl2:
        pg.locator('#btnCsvOperDay').click()
    check(dl2.value.suggested_filename.endswith('.xlsx'),
          '日ごとの稼働率をExcelで保存できる  → ' + dl2.value.suggested_filename)

    # ── 損紙率・予備率 ──
    #   ダミー日報は  実印刷 = 通し枚数 ／ 基準予備 = 通しの10% ／ やれ = 通しの5%
    #   → 出庫 = 通し×1.1、損紙率 = 5/110 = 4.55%、予備率 = 10/100 = 10.00%
    pg.locator('nav.tabs button[data-view=waste]').click()
    pg.wait_for_selector('#wasteSum tbody tr')
    names = pg.locator('#wasteSum tbody th.rh').all_text_contents()
    check(names == ['A全UV', '小森1号機', '菊全UV', '全機械'],
          '機械の行と全機械の行が出る  → %s' % names)
    tot = pg.locator('#wasteSum tbody tr.total td').all_text_contents()
    check(all(x == '4.55%' for x in tot),
          'どの月も 損紙率 5/110 = 4.55%%  → %s' % tot[:4])
    note = pg.locator('#wasteNote').inner_text()
    check('4.55%' in note and '18.85%' not in note and '10.00%' in note,
          '見出しに年間の損紙率と予備率が出る  → ' + note)
    tip = pg.locator('#wasteSum tbody tr.total td').first.get_attribute('title')
    check(all(k in tip for k in ('出庫', '実印刷', '基準予備', 'やれ', '通し')),
          'マスにもとの枚数が全部出る  → ' + tip)
    pg.locator('#pillWaste button[data-w=spare]').click()
    pg.wait_for_timeout(200)
    check(pg.locator('#wasteSum tbody tr.total td').first.inner_text() == '10.00%',
          '予備率に切り替えられる（基準予備 ÷ 実印刷 = 10.00%）')
    pg.locator('#pillWaste button[data-w=out]').click()
    pg.wait_for_timeout(200)
    check(pg.locator('#wasteSum tbody tr.total td').first.inner_text() == '188,954',
          '出庫枚数に切り替えられる（1月 188,954枚）')
    pg.locator('#pillWaste button[data-w=yare]').click()
    pg.wait_for_timeout(200)
    check(pg.locator('#wasteSum tbody tr.total td').first.inner_text() == '8,588',
          'やれ枚数に切り替えられる（1月 8,588枚）')
    pg.locator('#pillWaste button[data-w=rate]').click()

    # 日ごとの損紙率（日付順・機械順）
    pg.select_option('#selMonthWD', '1')
    pg.wait_for_timeout(200)
    wd = pg.locator('#wasteDay tbody tr')
    check(wd.count() == 32, '1月は31日ぶん＋月合計の32行  → %d行' % wd.count())
    dh = pg.locator('#wasteDay thead').inner_text()
    check('A全UV' in dh and '小森1号機' in dh and '全機械' in dh,
          '機械が列に並ぶ  → ' + dh.replace('\n', ' '))
    # ダミー日報は 1/6 に A全UV で 出庫46,200・やれ2,100 = 4.55%
    check(wd.nth(5).locator('td').nth(2).inner_text() == '4.55%',
          '1月6日 A全UV が 4.55%')
    check(wd.nth(3).locator('td').nth(2).inner_text() == '－'
          and 'off' in (wd.nth(3).get_attribute('class') or ''),
          '刷っていない日は灰色にして「－」にする')
    check('4.55%' in pg.locator('#wasteDay tbody tr.total').inner_text(),
          '月合計も 4.55%')
    pg.locator('#pillWasteDay button[data-d=spare]').click()
    pg.wait_for_timeout(200)
    check(wd.nth(5).locator('td').nth(2).inner_text() == '10.00%',
          '日ごとも予備率に切り替えられる（1/6 A全UV 10.00%）')
    pg.locator('#pillWasteDay button[data-d=rate]').click()
    pg.locator('#pillWasteDay button[data-d=yare]').click()
    pg.wait_for_timeout(200)
    check(wd.nth(5).locator('td').nth(2).inner_text() == '2,100',
          'やれ枚数に切り替えられる（1/6 A全UV 2,100枚）')
    pg.locator('#pillWasteDay button[data-d=out]').click()
    pg.wait_for_timeout(200)
    check(wd.nth(5).locator('td').nth(2).inner_text() == '46,200',
          '出庫枚数にも切り替えられる（1/6 A全UV 46,200枚）')
    pg.locator('#pillWasteDay button[data-d=rate]').click()
    with pg.expect_download(timeout=30000) as dl5:
        pg.locator('#btnCsvWasteDay').click()
    check(dl5.value.suggested_filename.endswith('.xlsx'),
          '日ごとの損紙率をExcelで保存できる  → ' + dl5.value.suggested_filename)

    wr = pg.locator('#wasteRank tbody tr')
    check(wr.count() == 70, '損紙の多い案件が69件＋合計行  → %d行' % wr.count())
    yares = [int(wr.nth(i).locator('td').nth(8).inner_text().replace(',', ''))
             for i in range(5)]
    check(yares == sorted(yares, reverse=True),
          'やれ枚数の多い順に並ぶ  → %s' % yares)
    check(wr.first.locator('td').nth(7).inner_text() == '85,554',
          '案件ごとの出庫枚数が出る  → ' + wr.first.locator('td').nth(7).inner_text())
    pg.locator('#pillWasteSort button[data-s=rate]').click()
    pg.wait_for_timeout(200)
    check(pg.locator('#wasteRank tbody tr').first.locator('td').nth(9).inner_text() == '4.55%',
          '損紙率の高い順にも並べ替えられる')
    pg.locator('#pillWasteSort button[data-s=yare]').click()
    pg.select_option('#selMonthW', '1')
    pg.wait_for_timeout(200)
    check('1月' in pg.locator('#wasteRankNote').inner_text(),
          '月を選べる  → ' + pg.locator('#wasteRankNote').inner_text())
    pg.select_option('#selMonthW', '0')
    pg.screenshot(path=str(SHOT / '11_損紙率.png'), full_page=True)
    with pg.expect_download(timeout=30000) as dl3:
        pg.locator('#btnCsvWaste').click()
    check(dl3.value.suggested_filename.endswith('.xlsx'),
          '損紙率をExcelで保存できる  → ' + dl3.value.suggested_filename)
    with pg.expect_download(timeout=30000) as dl4:
        pg.locator('#btnCsvWasteRank').click()
    check(dl4.value.suggested_filename.endswith('.xlsx'),
          '損紙の多い案件をExcelで保存できる  → ' + dl4.value.suggested_filename)

    # ── 日報明細: 日報作成に使った内容がそのまま出る ──
    pg.locator('nav.tabs button[data-view=raw]').click()
    pg.select_option('#selMonthR', '1')
    pg.wait_for_selector('#raw tbody tr')
    rr = pg.locator('#raw tbody tr')
    check(rr.count() == 8, '1月の日報明細が8行（統合前の生の行）  → %d行' % rr.count())
    rhdr = pg.locator('#raw thead').inner_text()
    check('通し枚数' in rhdr and '表版数' in rhdr and '裏版数' in rhdr, '日報の列がそのまま出る')
    check('通し枚数（集計）' in rhdr, '2か所に出る同名列は「（集計）」で区別される')
    check(pg.locator('#dailyT tbody tr').count() > 0, '日次集計ブロック（有効時間など）も見られる')
    check('1,110' not in pg.locator('#raw tbody').inner_text()
          and '1110' in pg.locator('#raw tbody').inner_text(),
          '営業担当ｺｰﾄﾞに桁区切りが付かない')
    check(pg.locator('#dailyT thead').inner_text().count('日付') == 1, '日次集計の日付列が重複しない')
    check('有効時間' in pg.locator('#dailyT tbody').inner_text(), '日次集計の項目名が出る')
    pg.screenshot(path=str(SHOT / '07_日報明細.png'), full_page=True)

    # 機械での絞り込み
    pg.select_option('#selMachineR', '新26・A全UV稼動日報')
    pg.wait_for_timeout(200)
    check(pg.locator('#raw tbody tr').count() == 5,
          '機械で絞り込める（A全UVは1月5行）  → %d行' % pg.locator('#raw tbody tr').count())
    pg.select_option('#selMachineR', '全機械')

    # ── 月別明細から日報の元の行を開く ──
    pg.locator('nav.tabs button[data-view=month]').click()
    pg.select_option('#selMonth', '1')
    pg.wait_for_selector('#detail tbody tr')
    tgt = pg.locator('#detail tbody tr', has_text='8632175').first
    tgt.locator('a.drill').click()
    pg.wait_for_selector('#detail tr.sub')
    sub = pg.locator('#detail tr.sub table.sub tbody tr')
    check(sub.count() == 2, '枝番2件の内訳（日報の元の行）が開く  → %d行' % sub.count())
    st = pg.locator('#detail tr.sub').inner_text()
    check('8632175-1-2' in st and '8,000' in st, '内訳に枝番の管理番号と通し枚数が出る')
    pg.screenshot(path=str(SHOT / '08_内訳ドリルダウン.png'), full_page=True)
    tgt.locator('a.drill').click()
    check(pg.locator('#detail tr.sub').count() == 0, 'もう一度押すと閉じる')

    # ── 元ファイル・検証 ──
    pg.locator('nav.tabs button[data-view=src]').click()
    pg.wait_for_selector('#files tbody tr')
    check(pg.locator('#files tbody tr').count() == 3, '読み込んだファイルが3件')
    # ファイル別の明細行の合計が、検証タブの明細行合計と一致すること
    fr = [int(pg.locator('#files tbody tr').nth(i).locator('td').last.inner_text().replace(',', ''))
          for i in range(3)]
    sr = [int(pg.locator('#stats tbody tr').nth(i).locator('td').nth(2).inner_text().replace(',', ''))
          for i in range(pg.locator('#stats tbody tr').count())]
    check(sum(fr) == sum(sr), 'ファイル別の明細行の合計が検証の合計と一致  → %d / %d'
          % (sum(fr), sum(sr)))
    check('不一致' not in pg.locator('#stats tbody').inner_text(), '検証: 統合前後の通し数がすべて一致')
    sp = pg.locator('#srcPath').inner_text()
    check('読込元フォルダ' in sp, '読込元フォルダが一覧で出る')
    check('/no/such/folder' in sp and '読み飛ばし' in sp,
          '無かったフォルダは警告ではなく元ファイルタブに出る')
    check('/no/such/folder' not in pg.locator('#warns').inner_text(),
          '無かったフォルダで注意バナーを出さない')
    pg.screenshot(path=str(SHOT / '04_元ファイル.png'), full_page=True)

    # ── Excel 出力 ──
    pg.locator('nav.tabs button[data-view=month]').click()
    with pg.expect_download() as d:
        pg.click('#btnCsvMonth')
    name = d.value.suggested_filename
    check(name.endswith('.xlsx') and '印刷実績' in name,
          '月別明細を Excel で保存できる → ' + name)

    # ── 年の切替（複数年フォルダ） ──
    pg.locator('nav.tabs button[data-view=year]').click()      # マトリクスを見える状態に
    yopts = pg.locator('#selYear option')
    if yopts.count() > 1:
        years = [int(yopts.nth(i).get_attribute('value')) for i in range(yopts.count())]
        check(years == sorted(years), '年が古い順に並ぶ  → %s' % years)
        check(pg.locator('#selYear').input_value() == str(years[-1]),
              '最初は一番新しい年が選ばれる')
        pg.select_option('#selYear', str(years[0]))
        pg.wait_for_function('document.querySelector("#ttl").textContent.indexOf("%d年") === 0'
                             % years[0])
        pg.wait_for_selector('#matrix tbody tr')
        check('%d年' % years[0] in pg.locator('#ttl').inner_text(),
              '%d年に切り替わる' % years[0])
        t0 = pg.locator('#matrix tbody.total tr').first.locator('td.yr').inner_text()
        pg.screenshot(path=str(SHOT / '09_年切替.png'), full_page=True)
        pg.select_option('#selYear', str(years[-1]))
        pg.wait_for_function('document.querySelector("#ttl").textContent.indexOf("%d年") === 0'
                             % years[-1])
        t1 = pg.locator('#matrix tbody.total tr').first.locator('td.yr').inner_text()
        check(t0 != t1, '年ごとに中身が入れ替わる  → %s / %s'
              % (t0.split(chr(10))[0], t1.split(chr(10))[0]))
    else:
        check(True, '年が1つだけのフォルダ（年の切替は非表示）')

    # ── 設定ダイアログ ──
    pg.click('#btnSetting')
    pg.wait_for_timeout(300)
    check('sample' in pg.locator('#inSrc').input_value(), '設定に現在のフォルダが入る')
    check('今回読み込んだフォルダ' in pg.locator('#altNote').inner_text(),
          '実際に読み込んだフォルダが分かる')
    pg.click('#dlgCancel')

    check(not errors, 'JavaScript エラーなし' + ('' if not errors else ': ' + '; '.join(errors[:3])))
    b.close()

print('\n合格 %d / 不合格 %d' % (len(ok), len(ng)))
sys.exit(1 if ng else 0)
