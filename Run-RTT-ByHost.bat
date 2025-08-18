@echo off
setlocal EnableExtensions DisableDelayedExpansion

rem ===== move to script directory (handles UNC/OneDrive) =====
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

rem ===== timestamp =====
set "YYYY=%date:~0,4%"
set "MM=%date:~5,2%"
set "DD=%date:~8,2%"
set "HH=%time:~0,2%"
set "NN=%time:~3,2%"
set "SS=%time:~6,2%"
set "HH=%HH: =0%"
set "TS=%YYYY%%MM%%DD%_%HH%%NN%%SS%"

set "OUTDIR=%OUTROOT%\%TS%"
mkdir "%OUTDIR%" >nul 2>nul
set "OUT=%OUTDIR%\TeamsNet-RTT-ByHost.xlsx"
set "LOG=%OUTDIR%\run.log"

rem ===== optional target.txt =====
set "TARGET=%BASE%target.txt"

rem ===== choose PowerShell host =====
set "PSCMD="
where powershell.exe >nul 2>nul && set "PSCMD=powershell.exe -NoProfile -ExecutionPolicy Bypass"
if not defined PSCMD (
  where pwsh.exe >nul 2>nul && set "PSCMD=pwsh.exe -NoProfile -ExecutionPolicy Bypass"
)
if not defined PSCMD (
  echo ERROR: PowerShell not found in PATH.
  popd & exit /b 1
)

rem ===== run (avoid embedding quotes in variables) =====
if exist "%TARGET%" (
  echo Command: %PSCMD% -File "%PS1%" -Output "%OUT%" -TargetsFile "%TARGET%" > "%LOG%"
  %PSCMD% -File "%PS1%" -Output "%OUT%" -TargetsFile "%TARGET%" 1>>"%LOG%" 2>&1
) else (
  echo Command: %PSCMD% -File "%PS1%" -Output "%OUT%" > "%LOG%"
  %PSCMD% -File "%PS1%" -Output "%OUT%" 1>>"%LOG%" 2>&1
)
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