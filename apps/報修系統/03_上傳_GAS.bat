@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo ========================================
echo   報修系統：將本機程式上傳到 GAS
echo ========================================
echo.

where clasp >nul 2>nul
if errorlevel 1 (
  echo [錯誤] 找不到 clasp。
  echo 請先安裝：npm install -g @google/clasp
  echo.
  pause
  exit /b 1
)

if not exist ".clasp.json" (
  echo [錯誤] 找不到 .clasp.json
  pause
  exit /b 1
)

findstr /C:"PASTE_SCRIPT_ID_HERE" ".clasp.json" >nul
if not errorlevel 1 (
  echo [尚未設定] 請先把 .clasp.json 的 PASTE_SCRIPT_ID_HERE
  echo 改成這個報修系統 GAS 專案的 Script ID。
  echo.
  pause
  exit /b 1
)

if not exist "src\Code.gs" (
  echo [錯誤] src\Code.gs 不存在，為避免誤上傳已停止。
  pause
  exit /b 1
)

echo 注意：本動作會把 src 內的程式上傳到 Google Apps Script。
echo 建議先執行 02_下載_GAS.bat 確認本機與 GAS 版本。
echo.
choice /C YN /M "確定要上傳嗎"
if errorlevel 2 (
  echo 已取消上傳。
  pause
  exit /b 0
)

echo.
echo 正在上傳...
call clasp push

if errorlevel 1 (
  echo.
  echo [失敗] 上傳未完成，請查看上方錯誤訊息。
) else (
  echo.
  echo [完成] src 程式已上傳到 GAS。
  echo 若有 Web App 功能變更，請再到 GAS 執行「部署 → 管理部署作業 → 新增版本」。
)

echo.
pause
