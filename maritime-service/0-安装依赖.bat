@echo off
setlocal
pushd "%~dp0"

where python >nul 2>nul
if %errorlevel%==0 (
  echo Installing Python dependencies with python...
  python -m pip install -r ".\requirements.txt"
  echo.
  pause
  exit /b %errorlevel%
)

where py >nul 2>nul
if %errorlevel%==0 (
  echo Installing Python dependencies with py...
  py -m pip install -r ".\requirements.txt"
  echo.
  pause
  exit /b %errorlevel%
)

echo [ERROR] Python was not found on this computer.
echo Please install Python 3.12+ first.
echo.
pause
exit /b 1
