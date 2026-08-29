# 稼動日報 印刷実績ビューア

印刷機の稼動日報（.xls）を読んで、**年別・月別・営業部別（本社・東京・池袋）**の印刷実績を
ブラウザで見るためのアプリです。

アプリ本体は **[`kadou-web/`](kadou-web/)** にあります。使い方・集計ルールは
[kadou-web/README.md](kadou-web/README.md) をご覧ください。

```
kadou-web/起動.bat    ← アプリを開く
プッシュ.bat           ← 変更を GitHub に送る（このフォルダ）
```

---

## PCでの始め方

### 1. 取ってくる

デスクトップに **`セットアップ.bat`** を置いてダブルクリックすると、
`arunasiweb` フォルダにこのリポジトリを取ってきます。
2回目以降は、同じファイルをダブルクリックすると最新に更新します。

`セットアップ.bat` はこのリポジトリに入れてあります。まだ何も無いPCでは、
GitHub の画面でこのファイルを開き、右上の **Download raw file** で
デスクトップに保存してください。すでに取得済みのPCなら
`arunasiweb\セットアップ.bat` をデスクトップにコピーしても構いません。

手で行う場合は次のとおりです。

```bat
cd C:\Users\<ユーザー名>\Desktop
git clone -b claude/factory-report-folder-check-b7a9kx https://github.com/komuazu/test.git arunasiweb
```

### 2. アプリを開く

`arunasiweb\kadou-web\起動.bat` をダブルクリックします。
ブラウザが開いて、稼動日報の集計が見られます。

### 3. 変更を送る（push）

`arunasiweb\プッシュ.bat` をダブルクリックすると、変更をまとめて GitHub に送ります。

* 変更内容の一言を聞かれます（そのまま Enter でも可）
* ネットワークが不安定なときは、待ち時間を延ばしながら4回まで再試行します
* 初回はブラウザで GitHub のログインを求められます

日本語のファイル名が開けない環境では `push.bat` / `setup.bat` を使ってください（中身は同じです）。

---

## 初回だけ必要なもの

| | 確認方法 | 入手先 |
|---|---|---|
| Git | `git --version` | https://git-scm.com/download/win |
| Python | `python --version` | https://www.python.org/downloads/windows/ |

Python はインストール時に **「Add python.exe to PATH」にチェック**を入れてください。

コミットが初めてのPCでは、名前とメールの設定が要ります。

```bat
git config --global user.name  "あなたの名前"
git config --global user.email "あなたのメール"
```

push には `komuazu/test` への書き込み権限が必要です。

---

## 各PCに残るもの（GitHubには送りません）

| ファイル | 中身 |
|---|---|
| `kadou-web/config.json` | 稼動日報フォルダの場所・対象年 |
| `kadou-web/memo.json` | 画面で入力した「今年の動向 / 代替対策 / 対策通し数」 |
| `kadou-web/web/years.json`, `data_<年>.json` | 起動のたびに作り直されるデータ |

これらは `.gitignore` に入れてあるので、push しても他のPCの設定を上書きしません。

**元の稼動日報（.xls）は読み取りのみで、一切変更しません。**
