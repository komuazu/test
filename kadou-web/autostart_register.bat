@echo off
chcp 932 >nul
setlocal
cd /d "%~dp0"

echo ============================================================
echo   稼動日報 印刷実績ビューア  自動起動の登録
echo ============================================================
echo.
echo   このPCにログオンしたとき、黒い画面を出さずにサーバーを開始します。
echo   ほかの端末からは、これまでどおり合い言葉を入れて見られます。
echo.

if not exist "password.txt" (
  echo [準備] password.txt がありません。
  echo   このフォルダに password.txt を作り、1行めに合い言葉を書いてください。
  echo.
  pause
  exit /b 1
)

rem ── 画面を出さない Python を探す ──────────────────────
set PYW=
where pythonw >nul 2>&1 && for /f "delims=" %%i in ('where pythonw') do set PYW=%%i
if not defined PYW (
  echo [エラー] pythonw.exe が見つかりません。
  echo   Python をインストールし直すか、サーバー起動.bat を使ってください。
  echo.
  pause
  exit /b 1
)

schtasks /create /tn "稼動日報ビューア" /f /sc onlogon /rl highest ^
  /tr ""%PYW%" "%CD%\server.py" --host 0.0.0.0 --no-browser"
if errorlevel 1 (
  echo.
  echo [エラー] 登録できませんでした。
  echo   このファイルを右クリックし「管理者として実行」でもう一度お試しください。
  echo.
  pause
  exit /b 1
)

echo.
echo 登録しました。次にこのPCにログオンしたときから自動で開始します。
echo いますぐ開始するには、続けて Enter を押してください。
pause >nul
schtasks /run /tn "稼動日報ビューア" >nul 2>&1
echo 開始しました。
echo   止めるとき      : 自動起動を解除.bat
echo   動いているか確認: 自動起動の状態.bat
echo.
pause
