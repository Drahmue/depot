@echo off
setlocal ENABLEDELAYEDEXPANSION
pushd "%~dp0"

set "LOGDIR=%CD%\logs"
if not exist "%LOGDIR%" mkdir "%LOGDIR%"

for /f %%I in ('powershell -NoProfile -Command "(Get-Date).ToString(\"yyyy-MM\")"') do set "LOGSTAMP=%%I"
set "LOGFILE=%LOGDIR%\depot_!LOGSTAMP!.log"

echo [%date% %time%] START Depot Script >> C:\ProgramData\bat_trace_depot.txt
echo LOGFILE=!LOGFILE! >> C:\ProgramData\bat_trace_depot.txt

set PYTHONIOENCODING=utf-8

rem --- Pull latest updates from GitHub ---
echo [%date% %time%] Pulling updates from GitHub >> "!LOGFILE!"
git pull origin main >> "!LOGFILE!" 2>&1
if errorlevel 1 (
  echo [%date% %time%] Git pull failed, continuing with existing code >> "!LOGFILE!"
)

rem --- Main script execution ---
"%CD%\.venv\Scripts\python.exe" -u "depot.py" >> "!LOGFILE!" 2>&1
set "RC=!ERRORLEVEL!"
echo [%date% %time%] ENDE Depot Script (ExitCode=!RC!) >> "!LOGFILE!"

rem --- Clean up old log files (older than 120 days) ---
powershell -NoProfile -Command ^
  "Get-ChildItem -Path '%LOGDIR%' -Filter 'depot_*.log' | Where-Object LastWriteTime -lt (Get-Date).AddDays(-120) | Remove-Item -Force -ErrorAction SilentlyContinue"  >nul 2>&1

popd
exit /b !RC!