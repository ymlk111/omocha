@echo off
rem ==========================================================================
rem  Split-ExcelRowsToSheets.ps1 をダブルクリックから実行するためのバッチ
rem  対話形式でパラメータを入力します。
rem ==========================================================================
chcp 65001 > nul
setlocal

set "PS_SCRIPT=%~dp0Split-ExcelRowsToSheets.ps1"

echo ============================================
echo  Excel行分割スクリプト
echo ============================================
echo.

set /p FILE_PATH=対象Excelファイルのフルパス: 
set /p SHEET_NAME=対象シート名: 
set /p START_ROW=データ開始行(数値): 

set /p PASTE_ROW=貼り付け先行(空Enterで1): 
if "%PASTE_ROW%"=="" set PASTE_ROW=1

echo.
echo 貼り付けモード:
echo   1) ValuesOnly  (値のみ・既定)
echo   2) AllContents (値+書式)
echo   3) Formulas    (数式保持)
set /p MODE_NUM=選択(空Enterで1): 
if "%MODE_NUM%"=="" set MODE_NUM=1
if "%MODE_NUM%"=="1" set PASTE_MODE=ValuesOnly
if "%MODE_NUM%"=="2" set PASTE_MODE=AllContents
if "%MODE_NUM%"=="3" set PASTE_MODE=Formulas

set /p OUTPUT_PATH=出力先パス(空Enterで元ファイルを上書き): 

echo.
echo ============================================
echo  実行します...
echo ============================================
echo.

if "%OUTPUT_PATH%"=="" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" ^
        -FilePath "%FILE_PATH%" ^
        -SheetName "%SHEET_NAME%" ^
        -StartRow %START_ROW% ^
        -PasteRow %PASTE_ROW% ^
        -PasteMode %PASTE_MODE%
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" ^
        -FilePath "%FILE_PATH%" ^
        -SheetName "%SHEET_NAME%" ^
        -StartRow %START_ROW% ^
        -PasteRow %PASTE_ROW% ^
        -PasteMode %PASTE_MODE% ^
        -OutputPath "%OUTPUT_PATH%"
)

echo.
pause
endlocal
