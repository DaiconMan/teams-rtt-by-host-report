@echo off
setlocal

rem === このbatと同じ場所にあるPS1/target.txtを使う ===
set "BASE=%~dp0"
set "PS1=%BASE%Generate-TeamsNet-RTT-ByHost.ps1"
if not exist "%PS1%" (
  echo ERROR: not found: "%PS1%"
  exit /b 1
)

rem === 出力先: BASE\Output\yyyyMMdd_HHmmss\TeamsNet-RTT-ByHost.xlsx ===
set "OUTROOT=%BASE%Output"
if not exist "%OUTROOT%" mkdir "%OUTROOT%"

for /f "usebackq delims=" %%T in (`powershell -NoProfile -Command "$([DateTime]::Now.ToString('yyyyMMdd_HHmmss'))"`) do set "TS=%%T"
set "OUTDIR=%OUTROOT%\%TS%"
mkdir "%OUTDIR%" >nul 2>nul
set "OUT=%OUTDIR%\TeamsNet-RTT-ByHost.xlsx"
set "LOG=%OUTDIR%\run.log"

rem ルートの target.txt があれば自動使用
set "TARGET=%BASE%target.txt"
set "TARGETARG="
if exist "%TARGET%" set "TARGETARG=-TargetsFile ""%TARGET%"""

rem 使う PowerShell を決定（Windows PowerShell → PowerShell 7 の順）
set "PSCMD="
where powershell >nul 2>nul && set "PSCMD=powershell -NoProfile -ExecutionPolicy Bypass"
if not defined PSCMD (
  where pwsh >nul 2>nul && set "PSCMD=pwsh -NoProfile -ExecutionPolicy Bypass"
)
if not defined PSCMD (
  echo ERROR: PowerShell not found. Install Windows PowerShell or PowerShell 7.
  exit /b 1
)

echo BASE   = "%BASE%"
echo OUTDIR = "%OUTDIR%"
echo OUT    = "%OUT%"
echo PS1    = "%PS1%"
echo TARGET = "%TARGET%"
echo Running: %PSCMD% -File "%PS1%" -Output "%OUT%" %TARGETARG%
echo ------------------------------------------------------------ > "%LOG%"
echo Command: %PSCMD% -File "%PS1%" -Output "%OUT%" %TARGETARG% >> "%LOG%"
echo ------------------------------------------------------------ >> "%LOG%"

%PSCMD% -File "%PS1%" -Output "%OUT%" %TARGETARG% 1>>"%LOG%" 2>&1
if errorlevel 1 goto :fail

if exist "%OUT%" (
  echo Done. Output: "%OUT%"
  start "" "%OUT%" 2>nul
  exit /b 0
) else (
  echo ERROR: Script reported success but file not found.
  echo ERROR: Expected output: "%OUT%"
  echo See log: "%LOG%"
  exit /b 2
)

:fail
echo.
echo === Failed with exit code %errorlevel% ===
echo See log: "%LOG%"
echo - よくある原因:
echo   1) teams_net_quality.csv が %%LOCALAPPDATA%%\TeamsNet に無い
echo   2) Excel(COM)起動不可 / Excel未インストール
echo   3) 出力先(UNC/NAS)への保存権限やロック
exit /b %errorlevel%