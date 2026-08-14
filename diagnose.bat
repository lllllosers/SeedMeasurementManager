@echo off
cd /d "%~dp0"
echo ========================================
echo Seed Measurement Manager - diagnostics
echo ========================================
echo.
where powershell.exe
if errorlevel 1 (
  echo ERROR: powershell.exe not found.
  pause
  exit /b 1
)
echo.
echo Starting app in diagnostic mode...
echo If an error appears, copy the red text or send error.log.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0app.ps1"
echo.
echo Application exited. Exit code: %errorlevel%
if exist "%~dp0error.log" (
  echo.
  echo ---- error.log ----
  type "%~dp0error.log"
)
pause
