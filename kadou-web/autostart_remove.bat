@echo off
chcp 932 >nul
setlocal

echo 稼動日報ビューアの自動起動を解除し、いま動いているものを止めます。
echo.
schtasks /end /tn "稼動日報ビューア" >nul 2>&1
schtasks /delete /tn "稼動日報ビューア" /f
if errorlevel 1 (
  echo   登録されていませんでした。
)
echo.
echo 解除しました。ほかの端末からは見られなくなります。
echo.
pause
