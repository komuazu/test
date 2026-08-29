@echo off
chcp 932 >nul
setlocal

echo ============================================================
echo   稼動日報 印刷実績ビューア  自動起動の状態
echo ============================================================
echo.
schtasks /query /tn "稼動日報ビューア" /fo list 2>nul
if errorlevel 1 (
  echo   自動起動は登録されていません。
  echo   登録するには 自動起動を登録.bat を実行してください。
)
echo.
echo ── いま開けるアドレス ──────────────────────────
ipconfig | findstr /c:"IPv4"
echo   上のアドレスに :8765 を付けて、ほかの端末のブラウザで開きます。
echo   例: http://192.168.1.20:8765/
echo.
pause
