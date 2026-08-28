@echo off
chcp 932 >nul
setlocal
cd /d "%~dp0"

echo ============================================================
echo   稼動日報 印刷実績ビューア
echo ============================================================
echo.

rem ── Python を探す ────────────────────────────────
set PY=
where py >nul 2>&1 && set PY=py -3
if not defined PY (where python >nul 2>&1 && set PY=python)
if not defined PY (
  echo [エラー] Python が見つかりません。
  echo   https://www.python.org/downloads/windows/ からインストールし、
  echo   インストール時に "Add python.exe to PATH" にチェックを入れてください。
  echo.
  pause
  exit /b 1
)

rem ── 必要なライブラリ（初回のみ導入） ──────────────
%PY% -c "import xlrd" >nul 2>&1
if errorlevel 1 (
  echo 初回セットアップ: 必要なライブラリを入れています…
  %PY% -m pip install --quiet --user xlrd
  if errorlevel 1 (
    echo [エラー] xlrd を入れられませんでした。社内ネットワークの制限をご確認ください。
    pause
    exit /b 1
  )
  echo 完了しました。
  echo.
)

rem ── 起動 ─────────────────────────────────────────
%PY% server.py %*
echo.
pause
