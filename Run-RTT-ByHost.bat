@echo off
setlocal

rem === このbatと同じ場所にあるPS1/target.txtを使う ===
set "BASE=%~dp0"
set "PS1=%BASE%Generate-TeamsNet-RTT-ByHost.ps1"
if not exist "%PS1%" (
  echo ERROR: not found: "%PS1%"
  exit /b 1
)

rem デスクトップの実パスを PowerShell で取得（OneDrive等でも安全）
for /f "usebackq delims=" %%D in (`powershell -NoProfile -Command "[Environment]::GetFolderPath('Desktop')"`) do set "DESK=%%D"
if not defined DESK set "DESK=%USERPROFILE%\Desktop"
set "OUT=%DESK%\TeamsNet-RTT-ByHost.xlsx"

rem ルートの target.txt があれば自動使用
set "TARGET=%BASE%target.txt"
set "TARGETARG="
if exist "%TARGET%" set "TARGETARG=-TargetsFile ""%TARGET%"""

rem 使う PowerShell を決定（Windows PowerShell → PowerShell 7 の順）
set "PSCMD="
where powershell >nul 2>nul && set "PSCMD=powershell -NoProfile -ExecutionPolicy Bypass"
if not defined PSCMD (
  where pwsh >nul 2>nul && set "PSCMD=pwsh -NoProfile"
)
if not defined PSCMD (
  echo ERROR: PowerShell not found. Install Windows PowerShell or PowerShell 7.
  exit /b 1
)

echo Running: %PSCMD% -File "%PS1%" -Output "%OUT%" %TARGETARG%
%PSCMD% -File "%PS1%" -Output "%OUT%" %TARGETARG%
set "ERR=%ERRORLEVEL%"

if not "%ERR%"=="0" (
  echo.
  echo === Failed with exit code %ERR% ===
  echo - Close the Excel file if "%OUT%" is open.
  echo - Ensure teams_net_quality.csv exists in %%LOCALAPPDATA%%\TeamsNet.
  exit /b %ERR%
)

echo Done. Output: "%OUT%"
start "" "%OUT%" 2>nul
endlocal