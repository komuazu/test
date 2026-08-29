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
    check(pg.locator('#chart .bar').count() == 12, '棒グラフが12か月分')

    # ── 前年との比較（2026年と2025年） ──
    check(pg.locator('#cmpSum tbody tr').count() == 6, '前年との比較が5部署+合計行')
    hdr = pg.locator('#cmpSum thead').inner_text()
    check('前年' in hdr and '増減' in hdr and '前年比' in hdr,
          '前年・増減・前年比の列がある')
    check('2025年' in pg.locator('#cmpNote').inner_text(),
          '比べている年が見出しに出る  → ' + pg.locator('#cmpNote').inner_text())
    check(pg.locator('#cmpMonth tbody tr').count() == 6, '月別前年比が5部署+合計行')
    tip = pg.locator('#cmpMonth tbody tr.total td').first.get_attribute('title')
    check('今年' in tip and '前年' in tip,
          '月別前年比のマスに今年と前年の数字が出る  → ' + tip)

    # ── グラフを前年比較に切り替える ──
    pg.locator('#pillChart button[data-c=yoy]').click()
    pg.wait_for_timeout(300)
    bars = pg.locator('#chart .bar').count()
    check(bars and pg.locator('#chart .pair .col').count() == bars * 2,
          'グラフが今年と前年の2本組になる  → %d組' % bars)
    check('2025年' in pg.locator('#legend').inner_text(),
          '凡例が今年と前年になる  → ' + pg.locator('#legend').inner_text().replace('\n', ' '))
    check(pg.locator('#selChartDept').is_visible(), '比べる部署を選べる')
    pg.select_option('#selChartDept', '本社営業部')
    pg.wait_for_timeout(300)
    check('本社営業部' in pg.locator('#chartNote').inner_text(),
          '選んだ部署が見出しに出る  → ' + pg.locator('#chartNote').inner_text())
    pg.screenshot(path=str(SHOT / '09_グラフ前年比較.png'), full_page=True)
    pg.locator('#pillChart button[data-c=stack]').click()
    pg.wait_for_timeout(300)
    check(pg.locator('#chart .seg').count() > 0, '積み上げに戻せる')
    check(pg.locator('#matrix tbody tr').count() == 6, 'マトリクスが5部署+合計行')
    total = pg.locator('#matrix tbody tr.total td').last.inner_text()
    check('1,874,777' in total, '年間合計 1,874,777 通し  → ' + total.replace('\n', ' '))
    pg.screenshot(path=str(SHOT / '01_年間サマリー.png'), full_page=True)

    # ── マトリクスの数字から月別明細へジャンプ ──
    pg.locator('#matrix a.jump').first.click()
    pg.wait_for_selector('#detail tbody tr')
    check(pg.locator('section#v-month.on').count() == 1, 'マトリクスのリンクで月別明細へ移動')

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

    # ── CSV 出力 ──
    pg.locator('nav.tabs button[data-view=month]').click()
    with pg.expect_download() as d:
        pg.click('#btnCsvMonth')
    name = d.value.suggested_filename
    check(name.endswith('.csv') and '印刷実績' in name, 'CSV出力できる → ' + name)

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
        t0 = pg.locator('#matrix tbody tr.total td').last.inner_text()
        pg.screenshot(path=str(SHOT / '09_年切替.png'), full_page=True)
        pg.select_option('#selYear', str(years[-1]))
        pg.wait_for_function('document.querySelector("#ttl").textContent.indexOf("%d年") === 0'
                             % years[-1])
        t1 = pg.locator('#matrix tbody tr.total td').last.inner_text()
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
