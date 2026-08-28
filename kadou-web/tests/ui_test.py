# -*- coding: utf-8 -*-
"""画面の動作確認（Playwright）: 表示・タブ切替・記入欄の保存・合計の再計算・CSV出力"""
import json
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

URL = 'http://127.0.0.1:8765/'
SHOT = Path(sys.argv[1] if len(sys.argv) > 1 else '.')
ok, ng = [], []


def check(cond, msg):
    (ok if cond else ng).append(msg)
    print(('  OK  ' if cond else '  NG  ') + msg)


with sync_playwright() as pw:
    b = pw.chromium.launch(executable_path='/opt/pw-browsers/chromium')
    pg = b.new_page(viewport={'width': 1500, 'height': 1000})
    errors = []
    pg.on('pageerror', lambda e: errors.append(str(e)))
    pg.on('console', lambda m: errors.append(m.text) if m.type == 'error' else None)

    pg.goto(URL)
    pg.wait_for_selector('#matrix tbody tr')

    # ── 年間サマリー ──
    check(pg.locator('.kpi').count() == 5, 'KPIカードが5枚（合計+3営業部）')
    check(pg.locator('#chart .bar').count() == 12, '棒グラフが12か月分')
    check(pg.locator('#matrix tbody tr').count() == 4, 'マトリクスが3営業部+合計行')
    total = pg.locator('#matrix tbody tr.total td').last.inner_text()
    check('1,797,000' in total, '年間合計 1,797,000 通し  → ' + total.replace('\n', ' '))
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
    check(rows.count() == 5, '1月・全社は5件に統合  → %d件' % rows.count())
    check('枝番2' in pg.locator('#detail tbody').inner_text(), '枝番統合のバッジが出る')
    check(pg.locator('#detail td.multi').count() == 2, '複数日にまたがる行の日付が黄色（2件）')
    hdr = pg.locator('#detail thead').inner_text().replace('\n', ' / ')
    check('今年の動向' in hdr and '無しの場合の代替対策' in hdr and '対策通し数' in hdr, '記入欄3列がある')
    check('94,000' in pg.locator('#detail tbody tr.total').inner_text(), '1月合計 94,000 通し')

    # ── 記入欄に入力 → 合計・差引・充足率が再計算され保存される ──
    first = pg.locator('#detail tbody tr:not(.total):not(.diff)').first
    first.locator('textarea[data-f=trend]').fill('継続（前年並み）')
    first.locator('input[data-f=tsu]').fill('12000')
    first.locator('input[data-f=tsu]').blur()
    pg.wait_for_timeout(900)
    check(pg.locator('#sumPlan').inner_text() == '12,000', '対策通し数の合計が 12,000')
    check(pg.locator('#sumDiff').inner_text() == '82,000', '差引が 94,000 − 12,000 = 82,000')
    check(pg.locator('#sumRate').inner_text() == '12.8%', '充足率 12.8%')
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
    check('1,797,000' in pg.locator('#clients tbody tr.total').inner_text(),
          '得意先別の合計が年間合計と一致')
    pg.screenshot(path=str(SHOT / '03_得意先別.png'), full_page=True)

    # ── 元ファイル・検証 ──
    pg.locator('nav.tabs button[data-view=src]').click()
    pg.wait_for_selector('#files tbody tr')
    check(pg.locator('#files tbody tr').count() == 3, '読み込んだファイルが3件')
    check('不一致' not in pg.locator('#stats tbody').inner_text(), '検証: 統合前後の通し数がすべて一致')
    pg.screenshot(path=str(SHOT / '04_元ファイル.png'), full_page=True)

    # ── CSV 出力 ──
    pg.locator('nav.tabs button[data-view=month]').click()
    with pg.expect_download() as d:
        pg.click('#btnCsvMonth')
    name = d.value.suggested_filename
    check(name.endswith('.csv') and '印刷実績' in name, 'CSV出力できる → ' + name)

    check(not errors, 'JavaScript エラーなし' + ('' if not errors else ': ' + '; '.join(errors[:3])))
    b.close()

print('\n合格 %d / 不合格 %d' % (len(ok), len(ng)))
sys.exit(1 if ng else 0)
