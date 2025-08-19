@echo off
setlocal EnableExtensions EnableDelayedExpansion

pushd "%~dp0" || (echo [ERROR] pushd failed & exit /b 1)

set "PS=Generate-TeamsNet-Report.ps1"
set "CSV=%LOCALAPPDATA%\TeamsNet\teams_net_quality.csv"
set "TARGETS=targets.csv"
set "FLOORS=floors.csv"
set "OUTDIR=%CD%\Output"
if not exist "%OUTDIR%" mkdir "%OUTDIR%"
set "OUT=%OUTDIR%\TeamsNet-Report.xlsx"
set "FLOORFILE=%CD%\%FLOORS%"

if exist "%FLOORFILE%" (
  echo [INFO] フロアマップを使用します: "%FLOORFILE%"
  powershell -NoProfile -ExecutionPolicy Bypass ^
    -File ".\%PS%" ^
    -CsvPath "%CSV%" ^
    -TargetsCsv ".\%TARGETS%" ^
    -Output "%OUT%" ^
    -BucketMinutes 5 -ThresholdMs 100 ^
    -FloorMap "%FLOORFILE%" 2^>^&1
) else (
  echo [INFO] floors.csv が無いためフロア色分けはスキップします
  powershell -NoProfile -ExecutionPolicy Bypass ^
    -File ".\%PS%" ^
    -CsvPath "%CSV%" ^
    -TargetsCsv ".\%TARGETS%" ^
    -Output "%OUT%" ^
    -BucketMinutes 5 -ThresholdMs 100 2^>^&1
)

set "ERR=%ERRORLEVEL%"
if not "%ERR%"=="0" (
  echo [ERROR] PowerShell script failed. ERRORLEVEL=%ERR%
  popd & exit /b %ERR%
)
echo [OK] Report generated: "%OUT%"
popd
exit /b 0