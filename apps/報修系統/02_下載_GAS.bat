@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo ========================================
echo   報修系統：從 GAS 下載最新程式
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

if not exist "src" mkdir "src"

echo 正在從 Google Apps Script 下載...
call clasp pull

if errorlevel 1 (
  echo.
  echo [失敗] 下載未完成，請查看上方錯誤訊息。
) else (
  echo.
  echo [完成] GAS 最新程式已下載到 src 資料夾。
)

echo.
pause
