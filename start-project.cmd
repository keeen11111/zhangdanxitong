@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0start-project.ps1" -Watchdog
if errorlevel 1 (
  echo.
  echo Project startup failed. Check the message above or the logs folder.
  pause
  exit /b 1
)
endlocal
