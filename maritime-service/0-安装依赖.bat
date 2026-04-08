@echo off
setlocal
pushd "%~dp0"
call :resolve_python
if errorlevel 1 exit /b 1
echo Installing Python packages...
%PY_CMD% -m pip install -r ".\requirements.txt"
if errorlevel 1 goto :run_failed
echo.
echo Installing Playwright Chromium...
%PY_CMD% -m playwright install chromium
if errorlevel 1 goto :run_failed
echo.
echo [OK] Dependencies are ready.
echo.
pause
exit /b 0

:run_failed
echo.
echo [ERROR] Dependency installation failed.
echo.
pause
exit /b 1

:resolve_python
where python >nul 2>nul
if %errorlevel%==0 (
  set "PY_CMD=python"
  exit /b 0
)
where py >nul 2>nul
if %errorlevel%==0 (
  set "PY_CMD=py"
  exit /b 0
)
echo [ERROR] Python was not found on this computer.
echo Install Python 3.12+ first, then run this BAT again.
echo.
pause
exit /b 1
