@echo off
setlocal EnableExtensions DisableDelayedExpansion

rem === Paths (use files next to this .bat) ===
set "BASE=%~dp0"
set "PS1=%BASE%Generate-TeamsNet-RTT-ByHost.ps1"
if not exist "%PS1%" (
  echo ERROR: PS1 not found: "%PS1%"
  exit /b 1
)

rem === Output: BASE\Output\yyyyMMdd_HHmmss\TeamsNet-RTT-ByHost.xlsx ===
set "OUTROOT=%BASE%Output"
if not exist "%OUTROOT%" mkdir "%OUTROOT%" >nul

for /f "usebackq delims=" %%T in (`powershell -NoProfile -Command "[DateTime]::Now.ToString('yyyyMMdd_HHmmss')"`) do set "TS=%%T"
set "OUTDIR=%OUTROOT%\%TS%"
mkdir "%OUTDIR%" >nul 2>nul
set "OUT=%OUTDIR%\TeamsNet-RTT-ByHost.xlsx"
set "LOG=%OUTDIR%\run.log"

rem === Optional targets file (target.txt beside this .bat) ===
set "TARGET=%BASE%target.txt"
set "TARGETARG="
if exist "%TARGET%" set "TARGETARG=-TargetsFile ""%TARGET%"""

rem === Choose PowerShell (Windows PowerShell -> PowerShell 7) ===
set "PSCMD="
where powershell >nul 2>nul  && set "PSCMD=powershell -NoProfile -ExecutionPolicy Bypass"
if not defined PSCMD (
  where pwsh >nul 2>nul && set "PSCMD=pwsh -NoProfile -ExecutionPolicy Bypass"
)
if not defined PSCMD (
  echo ERROR: PowerShell not found.
  exit /b 1
)

echo Command: %PSCMD% -File "%PS1%" -Output "%OUT%" %TARGETARG% > "%LOG%"
%PSCMD% -File "%PS1%" -Output "%OUT%" %TARGETARG% 1>>"%LOG%" 2>&1
if errorlevel 1 goto :fail

if exist "%OUT%" (
  echo DONE: "%OUT%"
  start "" "%OUT%" 2>nul
  exit /b 0
) else (
  echo ERROR: Script returned success but output missing.
  echo See log: "%LOG%"
  exit /b 2
)

:fail
echo FAILED with exit code %errorlevel%
echo See log: "%LOG%"
echo Hints: >> "%LOG%"
echo 1^) Ensure teams_net_quality.csv exists under %%LOCALAPPDATA%%\TeamsNet >> "%LOG%"
echo 2^) Ensure Microsoft Excel is installed and not blocked >> "%LOG%"
echo 3^) Check write permission or file lock in "!OUTDIR!" >> "%LOG%"
exit /b %errorlevel%