# Gmail 添付ファイル → Google ドライブ 自動保存

特定のラベル／件名の Gmail に届いた添付ファイルを、Google ドライブの指定フォルダーへ
自動で保存する Google Apps Script です。10 分おきに自動実行されます。

## セットアップ手順

### 1. スクリプトを作成
1. https://script.google.com を開き、「新しいプロジェクト」を作成
2. 既定の `Code.gs` の中身を消して、[`SaveAttachments.gs`](./SaveAttachments.gs) の内容を貼り付け

### 2. 設定を書き換える（ファイル冒頭の「設定」部分）
| 変数 | 内容 | 例 |
| --- | --- | --- |
| `SEARCH_QUERY` | 対象メールの検索条件 | `'label:請求書 has:attachment'` |
| `DEST_FOLDER_ID` | 保存先フォルダーの ID | URL の `folders/` の後ろの文字列 |
| `PROCESSED_LABEL` | 処理済みの目印ラベル名 | `'Drive保存済み'` |

**フォルダー ID の調べ方**: ドライブで保存先フォルダーを開き、URL
`https://drive.google.com/drive/folders/`**`1AbCdEf...`** の太字部分をコピー。

**ラベルでの絞り込み**: Gmail 側でフィルタを作り、対象メールに自動でラベルが
付くようにしておくと確実です（例: 特定の差出人 → ラベル「請求書」）。

### 3. 動作確認（手動実行）
1. 関数一覧から `saveGmailAttachmentsToDrive` を選び「実行」
2. 初回は Google の認可画面が出るので許可（Gmail 読み取り・ドライブ書き込み）
3. ドライブのフォルダーに添付ファイルが入れば成功

### 4. 自動化をオンにする
1. 関数一覧から `createTrigger` を選び「実行」
2. これで 10 分おきに自動実行されます

## 仕組み・注意点
- 処理済みのメールには `PROCESSED_LABEL` のラベルが付き、**重複保存されません**。
- 既定では同名ファイルがあればスキップします（`SKIP_IF_EXISTS`）。
- `ORGANIZE_BY_MONTH = true` にすると受信月ごと（例 `2026-06`）のサブフォルダーに整理します。
- 実行間隔を変えたい場合は `createTrigger` 内の `everyMinutes(10)` を調整してください。
- スクリプトは**過去メールにもさかのぼって**適用されます。直近分だけにしたい場合は
  `SEARCH_QUERY` に `newer_than:7d` などを追加してください。
