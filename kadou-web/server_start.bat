@echo off
chcp 932 >nul
setlocal
cd /d "%~dp0"

echo ============================================================
echo   稼動日報 印刷実績ビューア  ＜みんなで使う＞
echo ============================================================
echo.

if not exist "password.txt" (
  echo [準備] password.txt がありません。
  echo   このフォルダに password.txt を作り、1行めに合い言葉を書いてください。
  echo   ほかの端末から開くときに、この合い言葉を入れてもらいます。
  echo.
  pause
  exit /b 1
)

rem ── 使う Python を探す ────────────────────────────────
set PY=
where py >nul 2>&1 && set PY=py -3
if not defined PY (where python >nul 2>&1 && set PY=python)
if not defined PY (
  echo [エラー] Python が見つかりません。
  echo   https://www.python.org/downloads/windows/ からインストールし、
  echo   インストール時に "Add python.exe to PATH" にチェックしてください。
  echo.
  pause
  exit /b 1
)

%PY% -c "import xlrd" >nul 2>&1
if errorlevel 1 (
  echo 初回セットアップ: 必要なライブラリを入れています…
  %PY% -m pip install --quiet --user xlrd
  if errorlevel 1 (
    echo [エラー] xlrd を入れられませんでした。
    pause
    exit /b 1
  )
)

rem ── ほかの端末からも届くように起動する ────────────────
echo この画面を閉じると、ほかの端末から見られなくなります。
echo このPCのブラウザも開きます（このPCからは合い言葉なしで使えます）。
echo 終了するには Ctrl+C を押してください。
echo.
%PY% server.py --host 0.0.0.0 %*
echo.
pause
