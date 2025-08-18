@echo off
setlocal EnableExtensions DisableDelayedExpansion

rem ===== move to script directory (handles UNC/OneDrive paths safely) =====
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

rem ===== timestamp (locale-agnostic via %date%/%time% slicing) =====
set "YYYY=%date:~0,4%"
set "MM=%date:~5,2%"
set "DD=%date:~8,2%"
set "HH=%time:~0,2%"
set "NN=%time:~3,2%"
set "SS=%time:~6,2%"
rem space-padded hour -> zero-pad
set "HH=%HH: =0%"
set "TS=%YYYY%%MM%%DD%_%HH%%NN%%SS%"

set "OUTDIR=%OUTROOT%\%TS%"
mkdir "%OUTDIR%" >nul 2>nul
set "OUT=%OUTDIR%\TeamsNet-RTT-ByHost.xlsx"
set "LOG=%OUTDIR%\run.log"

rem ===== optional target.txt beside this .bat =====
set "TARGET=%BASE%target.txt"
set "EXTRA_ARGS="
if exist "%TARGET%" (
  rem -> keep as single variable to avoid double quoting
  set "EXTRA_ARGS=-TargetsFile \"%TARGET%\""
)

rem ===== choose PowerShell host (Windows PowerShell first, then PowerShell 7) =====
set "PSCMD="
where powershell.exe >nul 2>nul && set "PSCMD=powershell.exe -NoProfile -ExecutionPolicy Bypass"
if not defined PSCMD (
  where pwsh.exe >nul 2>nul && set "PSCMD=pwsh.exe -NoProfile -ExecutionPolicy Bypass"
)
if not defined PSCMD (
  echo ERROR: PowerShell not found in PATH.
  popd & exit /b 1
)

rem ===== run and capture stdout/stderr to log =====
echo Command: %PSCMD% -File "%PS1%" -Output "%OUT%" %EXTRA_ARGS% > "%LOG%"
%PSCMD% -File "%PS1%" -Output "%OUT%" %EXTRA_ARGS% 1>>"%LOG%" 2>&1
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