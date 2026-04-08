@echo off
setlocal
pushd "%~dp0"
call :resolve_python
if errorlevel 1 exit /b 1
%PY_CMD% ".\launch_ams_desktop_app.py"
exit /b %errorlevel%

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
echo Please install Python 3.12+ first.
echo.
pause
exit /b 1
