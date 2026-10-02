@echo off
setlocal
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is required. Install the LTS version from https://nodejs.org/
  pause
  exit /b 1
)

if not exist "node_modules\playwright-core" (
  echo First run: installing the uploader package...
  call npm install --no-audit --no-fund
  if errorlevel 1 goto :failed
)

node uploader.mjs "%~1"
if errorlevel 1 goto :failed

echo.
echo Finished.
pause
exit /b 0

:failed
echo.
echo The uploader stopped because of an error.
pause
exit /b 1
