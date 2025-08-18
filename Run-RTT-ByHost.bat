@echo off
setlocal EnableExtensions DisableDelayedExpansion

rem ===== change to script directory (handles UNC via temporary drive mapping) =====
set "BASE=%~dp0"
pushd "%BASE%" >nul 2>nul

rem ===== paths =====
set "PS1=%BASE%Generate-TeamsNet-RTT-ByHost.ps1"
if not exist "%PS1%" (
  echo ERROR: PS1 not found: "%PS1%"
  popd & exit /b 1
)

set "OUTROOT=%BASE%Output"
if not exist "%OUTROOT%" mkdir "%OUTROOT%" >nul 2>nul

rem ===== timestamp (locale-agnostic using cmd variables) =====
set "YYYY=%date:~0,4%"
set "MM=%date:~5,2%"
set "DD=%date:~8,2%"
set "HH=%time:~0,2%"
set "NN=%time:~3,2%"
set "SS=%time:~6,2%"
rem zero-pad hour if space-padded
set "HH=%HH: =0%"
set "TS=%YYYY%%MM%%DD%_%HH%%NN%%SS%"

set "OUTDIR=%OUTROOT%\%TS%"
mkdir "%OUTDIR%" >nul 2>nul
set "OUT=%OUTDIR%\TeamsNet-RTT-ByHost.xlsx"
set "LOG=%OUTDIR%\run.log"

rem ===== optional targets file placed next to this .bat =====
set "TARGET=%BASE%target.txt"
set "TARGETARG="
if exist "%TARGET%" set "TARGETARG=-TargetsFile ""%TARGET%"""

rem ===== choose PowerShell host (Windows PowerShell > PowerShell 7) =====
set "PSCMD="
where powershell.exe >nul 2>nul && set "PSCMD=powershell.exe -NoProfile -ExecutionPolicy Bypass"
if not defined PSCMD (
  where pwsh.exe >nul 2>nul && set "PSCMD=pwsh.exe -NoProfile -ExecutionPolicy Bypass"
)
if not defined PSCMD (
  echo ERROR: PowerShell not found in PATH.
  popd & exit /b 1
)

rem ===== run =====
echo Command: %PSCMD% -File "%PS1%" -Output "%OUT%" %TARGETARG% > "%LOG%"
%PSCMD% -File "%PS1%" -Output "%OUT%" %TARGETARG% 1>>"%LOG%" 2>&1
set "ERR=%ERRORLEVEL%"

if not "%ERR%"=="0" (
  echo FAILED with exit code %ERR%
  echo See log: "%LOG%"
  popd & exit /b %ERR%
)

if exist "%OUT%" (
  echo DONE: "%OUT%"
  start "" "%OUT%" 2>nul
  popd & exit /b 0
) else (
  echo ERROR: Script returned success but output file is missing.
  echo See log: "%LOG%"
  popd & exit /b 2
)