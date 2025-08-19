@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem --- 文字コード注意: このファイルは UTF-8 (BOMなし) or ANSI、改行は CRLF で保存 ---
rem --- UNCパスでも安全にスクリプトフォルダへ移動（pushdはUNCを一時ドライブにマップ）---
pushd "%~dp0" || (echo [ERROR] pushd failed & exit /b 1)

rem 出力先・ターゲットなど必要ならここで設定
set "OUTDIR=%CD%\Output"
if not exist "%OUTDIR%" mkdir "%OUTDIR%"

rem PowerShell スクリプト呼び出し（例）
powershell -NoProfile -ExecutionPolicy Bypass ^
  -File ".\Generate-TeamsNet-Report.ps1" ^
  -CsvPath "%LOCALAPPDATA%\TeamsNet\teams_net_quality.csv" ^
  -TargetsCsv ".\targets.csv" ^
  -Output "%OUTDIR%\TeamsNet-Report.xlsx" ^
  -BucketMinutes 5 -ThresholdMs 100 2^>^&1

set "ERR=%ERRORLEVEL%"

if not "%ERR%"=="0" (
  echo [ERROR] PowerShell script failed. ERRORLEVEL=%ERR%
  popd
  exit /b %ERR%
)

echo [OK] Report generated: "%OUTDIR%\TeamsNet-Report.xlsx"
popd
exit /b 0