@echo off
chcp 932 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ============================================================
echo   変更を GitHub に push します
echo ============================================================
echo.

rem ── git があるか ──────────────────────────────
where git >nul 2>&1
if errorlevel 1 (
  echo [エラー] git が見つかりません。
  echo   https://git-scm.com/download/win からインストールしてください。
  echo.
  pause
  exit /b 1
)

rem ── ここが git の作業フォルダか ────────────────
git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
  echo [エラー] このフォルダは git の作業フォルダではありません。
  echo   セットアップ.bat で clone したフォルダの中で実行してください。
  echo.
  pause
  exit /b 1
)

for /f "delims=" %%i in ('git rev-parse --abbrev-ref HEAD') do set BR=%%i
echo 現在のブランチ: !BR!
echo.

rem ── 変更があるか ──────────────────────────────
git add -A
git diff --cached --quiet
if not errorlevel 1 (
  echo コミットする変更はありません。
  echo 未 push のコミットがあれば push します。
  goto :push
)

echo ── 変更内容 ────────────────────────────────
git diff --cached --stat
echo.
set "MSG="
set /p MSG=変更内容を一言で（そのまま Enter で「更新」）: 
if "!MSG!"=="" set MSG=更新
git commit -m "!MSG!"
if errorlevel 1 (
  echo [エラー] コミットできませんでした。
  echo   初回は名前とメールの設定が要ります:
  echo     git config --global user.name  "あなたの名前"
  echo     git config --global user.email "あなたのメール"
  echo.
  pause
  exit /b 1
)
echo.

:push
rem ── push（ネットワークが不安定なとき用に4回まで再試行）──
set TRY=0
set WAIT=2
:retry
set /a TRY+=1
echo push しています（%TRY% 回目）…
git push -u origin "!BR!"
if not errorlevel 1 goto :ok
if %TRY% geq 5 (
  echo.
  echo [エラー] push できませんでした。
  echo   ・初回はブラウザで GitHub のログインを求められます
  echo   ・komuazu/test への書き込み権限が要ります
  echo.
  pause
  exit /b 1
)
echo   失敗しました。%WAIT% 秒待って再試行します…
timeout /t %WAIT% /nobreak >nul
set /a WAIT*=2
goto :retry

:ok
echo.
echo ============================================================
echo   push 完了
echo   https://github.com/komuazu/test/tree/!BR!
echo ============================================================
echo.
pause
