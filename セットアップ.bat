@echo off
chcp 932 >nul
setlocal
cd /d "%~dp0"

set REPO=https://github.com/komuazu/test.git
set BRANCH=claude/factory-report-folder-check-b7a9kx
set DIR=arunasiweb
set TARGET=%CD%\%DIR%

rem ── このファイルがすでに取得済みのフォルダの中にあるとき ────────
rem   arunasiweb の中で実行すると arunasiweb\arunasiweb と入れ子になって
rem   しまうため、そのときは入れ子を作らず、このフォルダ自身を更新する。
if exist "%~dp0.git" if exist "%~dp0kadou-web" (
  set DIR=.
  set TARGET=%CD%
  echo このフォルダ自身が取得済みです。ここを最新にします。
  echo.
)

echo ============================================================
echo   稼動日報ビューア  セットアップ
echo     %REPO%
echo     ブランチ: %BRANCH%
echo     取得先  : %TARGET%
echo ============================================================
echo.

rem ── git があるか ──────────────────────────────
where git >nul 2>&1
if errorlevel 1 (
  echo [エラー] git が見つかりません。
  echo   https://git-scm.com/download/win からインストールしてください。
  echo   （インストールは既定のままで大丈夫です）
  echo.
  pause
  exit /b 1
)

rem ── すでに clone 済みなら最新を取得するだけ ────
if exist "%DIR%\.git" (
  echo すでにセットアップ済みです。最新を取得します。
  pushd "%DIR%"
  rem まだ送っていない変更があるときは、上書きせずに止める
  git diff-index --quiet HEAD --
  if errorlevel 1 goto :dirty
  git fetch origin %BRANCH%
  if errorlevel 1 goto :fetchfail
  git checkout %BRANCH%
  git pull origin %BRANCH%
  popd
  goto :done
)

rem ── 中身のあるフォルダがあれば退避してから clone ──
if exist "%DIR%" (
  echo %DIR% がすでにあります。念のため %DIR%_backup に退避します。
  if exist "%DIR%_backup" (
    echo [エラー] %DIR%_backup もすでにあります。
    echo   手で名前を変えるか消してから、もう一度実行してください。
    echo.
    pause
    exit /b 1
  )
  ren "%DIR%" "%DIR%_backup"
  if errorlevel 1 (
    echo [エラー] 名前を変えられませんでした。フォルダを開いていないか確認してください。
    echo.
    pause
    exit /b 1
  )
)

echo clone しています…
git clone -b %BRANCH% %REPO% "%DIR%"
if errorlevel 1 goto :fetchfail

:done
echo.
echo ============================================================
echo   完了しました
echo.
echo   アプリを開く   : %TARGET%\kadou-web\起動.bat をダブルクリック
echo   変更を送る     : %TARGET%\プッシュ.bat をダブルクリック
echo   最新を取り込む : このファイルをもう一度ダブルクリック
echo ============================================================
echo.
pause
exit /b 0

:dirty
popd 2>nul
echo.
echo [中止] このフォルダには、まだ GitHub に送っていない変更があります。
echo   上書きしてしまわないよう、更新を止めました。
echo   プッシュ.bat で変更を送ってから、もう一度お試しください。
echo.
pause
exit /b 1

:fetchfail
popd 2>nul
echo.
echo [エラー] GitHub に接続できませんでした。
echo   ・社内ネットワークの制限がないか
echo   ・初回はブラウザで GitHub のログインを求められます
echo   ・komuazu/test を見る権限があるか
echo.
pause
exit /b 1
