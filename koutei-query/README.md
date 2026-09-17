# koutei-query — 受注番号から「仕上りサイズ・加工内容・内外作・委託先・用紙」を引く

平版印刷 → オンデマンド移行検討で、`平版印刷_オンデマンド移行検討_3000通し以下一覧.xlsx` の
「仕上りサイズ」「備考」を埋めるための道具。koutei-kanr30（社内工程管理システム）の
**本番 DB（丸本PC）を読むだけ**で、DB もコードも書き換えない。

```
koutei-query/
├── export_finish_size.py     … 本体（丸本PC で動かす）
├── query_finish_size.sql     … 同じことを psql だけでやる版（検算用）
├── tests/fixture.sql         … 動作確認用のダミーデータ（本番に流さない）
├── tests/run_test.py         … 動作確認（Python 版と SQL 版の両方）
└── tests/compare_on_dump.py  … 本番ダンプで Python 版と SQL 版を全件突き合わせる
```

---

## 1. 先に知っておくこと（調べて分かったこと）

### 製造指示書・印刷指示書の PDF は koutei-kanr30 が作ったものではない

Google ドライブの `作業依頼書` フォルダにある「【製造指示書】」PDF（例: 08728296 障害者採用パンフレット）は
**基幹の受注・生産管理システム（MIS）の帳票**で、`部品／内・外作／加工所／加工日程／加工作業／台数／仕上サイズ／展開サイズ`
の表もそちらのもの。koutei-kanr30 のコード（`komuazu/koutei`）には、この表を持つテーブルも、
「製造指示書」「印刷指示書」を出力する処理も存在しない（あるのは「外注委託依頼書」だけ）。

koutei-kanr30 が持っているのは、MIS の CSV から取り込んだ項目と、画面で手入力した項目に限られる。

| 欲しい項目 | koutei-kanr30 での在りか | 元 |
|---|---|---|
| 仕上りサイズ | `timeline_processes.data` / `waiting_list.data` の JSON `finish_size`（最上位と `events[]` の要素）、`outsourcing_list.finish_size`、`delivery_schedule.finish_size` | MIS CSV の「製品仕上サイズ名」（`calendar_ui.html` の取込）または手入力 |
| 加工内容 | JSON `finish_processing_name` / `finish_processing` / `finishProcessing` / `processing_content` / `finish_process`（最上位と `events[]`）、`outsourcing_items[].processing_content`、`outsourcing_list.processing_content`、`delivery_schedule.processing_content`、`processing_works.processing_content` / `classification` | MIS CSV の「仕上加工名」または手入力 |
| 内作／外注 | JSON `internal_work` / `outsourcing` / `work_department`、`processing_works` に行があれば内作、`outsourcing_list` に行があれば外注 | 画面のチェック |
| 委託先名 | JSON `outsourcing_company` / `outsource_name` / `outsourcing_items[].company`、`outsourcing_list.outsourcing_company`、`delivery_schedule.outsourcing_company` | MIS CSV の「外注先名」または手入力 |
| 用紙銘柄・規格・斤量 | JSON `paper_type` / `standard_size` / `paper_weight`、部品ごとの `order_paper_info[]`・`papers[]` の同名キー | 画面の用紙選択（`paper_master` から） |

**MIS から取り込んだ「仕上加工名」は `data.events[]` の中にしか残らない**（`calendar_ui.html` 5295〜5309 行で最上位へ
写しているのは `finish_size` と `outsource_name` だけ）。2025-12 のダンプでは、最上位の `finish_process` は 0 行、
`events[].finish_process` に値があるのは 45 受注だった。スクリプトは最上位と `events[]` の両方を読む。

つまり **加工所（第二工場など）・加工日程・台数は「外注委託依頼書」を出した案件と、内作加工を登録した案件にしか入っていない。**
MIS の製造指示書と同じ粒度がどうしても要る項目は、MIS 側から取る必要がある。

### 「仕上りサイズ B1」は MIS 一括取込の既定値の疑いがある

2025-12 のダンプでは、`waiting_list` の `finish_size = 'B1'` が 874 行あり、**全部が同じ分（2025-12-12 09:52）の
一括取込**で、受注数量「1」が 768 行、仕上加工名は「-」だった。手入力側（タイムライン）の仕上りサイズは A4 が最多で
B1 はほぼ無い。MIS の「製品仕上サイズ名」が未入力のとき B1 になっていないか、MIS 側で確かめること。
スクリプトは B1 だけの案件の件数を末尾に出す。詳細 CSV の「作成日時」で一括取込かどうか見分けられる。

### 受注番号には枝が付いていることがある

koutei には `08701248（表紙）`、`08704276②`、`08704276③（複製1）`、`08702636複製1` のように、部品や複製で
受注番号の後ろに文字が付いた行がある（ダンプで 283 行・87 受注。うち 8 受注は素の番号の行が無い）。
**先頭の数字だけで突き合わせ**、枝は詳細 CSV の「枝」列に残す。本体 CSV では部品ぶんの値を ` / ` でつなぐ。
※ ② のような丸数字は NFKC 正規化すると「2」になって別の受注番号に化けるので、正規化は数字部分の全角→半角だけにしている。

### 本番 DB は丸本PC の中からしか見えない

`web_app/.env.marumoto` の接続先は `localhost:5432`、DB 名は `koutei_kanri_dev`（本番機でもこの名前）。
PostgreSQL は丸本PC の自分自身にしか口を開けていないので、**このスクリプトは丸本PC の上で動かす。**
クラウド側の Claude Code からは届かない（2026-09-17 に確認済み）。

### 受注番号一覧 CSV は同梱されていなかった

`仕上りサイズ未確認_受注番号一覧.csv` はこのリポジトリにも Google ドライブにも無かった。
手元のものを丸本PC に置いて `--orders` で渡す。

---

## 2. 丸本PC での動かし方

前提: koutei-kanr30 が動いている PC（`C:\Users\116544\Desktop\UPDATE17` に本体がある）。
Python と `psycopg2` は koutei 本体が使っているものがそのまま使える。

```bat
cd /d <このフォルダ>
python export_finish_size.py --orders 仕上りサイズ未確認_受注番号一覧.csv ^
    --env C:\Users\116544\Desktop\UPDATE17\web_app\.env
```

* `--env` には koutei が使っている `.env`（`DATABASE_HOST` などが書いてあるファイル）を渡す。
  省略すると、カレント／`web_app/` の `.env` → `.env.marumoto` の順に探す
* パスワードは `.env` の `DATABASE_PASSWORD` を使う。無ければその場で聞く。**画面にもファイルにも出さない**
* 接続は `default_transaction_read_only=on` で開く。うっかり書き込み文を流しても DB 側で拒否される
  （`tests/run_test.py` で DELETE が拒否されることを確かめている）
* 受注番号一覧は 1 列目、または「受注番号」「管理番号」を含む見出しの列を読む。
  UTF-8 でも CP932（Excel の CSV 保存）でもよい。数字で始まらない値（空欄・「0」など）は飛ばして件数を出す

### 受注番号のゆれ

| 表記 | 例 | 扱い |
|---|---|---|
| PDF・koutei | `08726258` | 前ゼロを落として `8726258` |
| Excel | `8726258` / `8726258.0` | `.0` と前ゼロを落として `8726258` |
| 全角・空白入り | `８７２６２５８`、`8726 258` | 数字を半角にし、空白を除く |
| 枝付き（koutei 側） | `08726258（表紙）`、`08726258②` | 先頭の数字 `8726258` で突き合わせ、枝は詳細 CSV に残す |

DB 側も同じ規則で比べるので、どちらの表記で入っていても拾う。

---

## 3. 出力

### `仕上りサイズ_koutei取得結果.csv`（Excel にそのまま貼る用）

```
受注番号, 仕上りサイズ, 加工内容, 内外作区分, 委託先名, 用紙銘柄, 用紙規格, 斤量
```

* 依頼メモの 5 列に、用紙の 3 列を後から足した。要らなければ Excel 側で読み飛ばせばよい
* 入力 CSV の並びのまま、**受注番号は入力の表記のまま**返す（XLOOKUP のキーにそのまま使える）
* 該当が無い受注番号は空欄
* 同じ受注番号に複数行（本体／表紙などの部品、返し・続きイベント、枝付き）があるときは、
  重複を除いて ` / ` でつなぐ。並びは読んだ順（テーブルの順 → 行 ID 順）で固定
* 内外作区分は `内作` / `外注` / `内作・外注`（両方ある）/ 空欄（判定材料なし）
  * 内作 = 内作チェックあり、内作加工の作業一覧に行あり、または加工所が 第二工場／本社／POP課
    （「営業」「その他」「-」は内作と判定しない）
  * 外注 = 外注チェックあり、外注依頼書あり、または委託先名あり（「未手配」は委託先と数えない）
* 「-」「未設定」「未定」「なし」は値なし扱い。`processing_works` の区分「その他」は加工内容に出さない
* 委託先名の表記ゆれ（`㈱松岡製本`／`松岡製本（株）`／`松岡製本`）はそのまま別名で並ぶ。名寄せはしない
* UTF-8（BOM 付き）、CRLF。Excel でそのまま開ける

### `仕上りサイズ_koutei取得結果_詳細.csv`（検算用）

どのテーブルのどの行から何を取ったかを 1 行ずつ。「枝」「部品」「作成日時」があるので、
部品ごとの値の切り分けと、B1 の一括取込の見分けに使う。`委託先ごとの加工内容` は
`八王子紙工：ミシン(筋)` の形で、備考欄の文面を作るときに使える。

---

## 4. 動作確認

本番に触らずに確かめるには、空の PostgreSQL に `tests/fixture.sql` を流して走らせる。

```bash
python tests/run_test.py --dsn "host=localhost port=5432 dbname=koutei_test user=koutei_dev"
```

* DB 名に `test` が入っていないと止まる（本番に fixture を流さないための安全弁）
* fixture の列定義は本番ダンプ（koutei リポジトリ `web_app/data/pg_backup.sql`、2025-12-15）の
  `CREATE TABLE` を写している（`delivery_schedule.outsourcing` が TEXT、など）
* 8 件の受注番号（ゼロ埋め・小数・全角・枝付き・該当なし・壊れた JSON・「-」「未手配」を含む）で 8 列を突き合わせ、
  SQL 版も同じ結果になることを見る

本番ダンプを別のテスト DB に読み込んで、Python 版と SQL 版を全件（2,164 受注）突き合わせるには:

```bash
psql -d koutei_dump_test -f <koutei>/web_app/data/pg_backup.sql
python tests/compare_on_dump.py --dsn "host=localhost port=5432 dbname=koutei_dump_test user=koutei_dev"
```

2026-09-17 の結果: 差 0 件。内訳は 仕上りサイズあり 2,099／加工内容あり 449／用紙銘柄あり 917／B1 だけ 892。

### SQL 版（`query_finish_size.sql`）の注意

* `pg_input_is_valid()` を使うので PostgreSQL 16 以降（本番は 18）
* 5 テーブルが全部ある前提。無いテーブルがあるとエラーになる（Python 版は飛ばす）
* 一時テーブル `target` を作ってから `SET default_transaction_read_only = on` の順（ファイル先頭の手順どおり）
* ` / ` でつないだ値の並びは文字コード順で、Python 版（読んだ順）と違う。中身は同じ

---

## 5. 決めごと

* **本番 DB に書かない。** 接続は読み取り専用で開く。`fixture.sql` は `DROP TABLE` を含むので本番には流さない
* **koutei に無いものは無いと言う。** 加工所・加工日程・台数は外注委託依頼書／内作加工を登録した案件にしか無い。
  空欄が多い項目は「入力されていない」のであって、スクリプトの取りこぼしとは限らない。
  疑うときは `_詳細.csv` で行単位に戻って見る
* **受注番号で機械的につなぐだけ。** 得意先名や品名での名寄せはしない
* **直したら `run_test.py` と `compare_on_dump.py` を通し直す。** Python 版と SQL 版がずれたまま
  「検算用」と言わない
