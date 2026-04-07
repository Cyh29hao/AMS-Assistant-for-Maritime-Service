@echo off
setlocal
pushd "%~dp0"
if "%~1"=="" (
  echo Drag a JSON request file onto this BAT file.
  echo.
  pause
  exit /b 1
)
call :resolve_python
if errorlevel 1 exit /b 1
echo Building contract from JSON...
%PY_CMD% ".\scripts\contract_workflow.py" from-json --input "%~1" --output-dir ".\output\contracts"
echo.
echo Output folder:
echo .\output\contracts
echo.
pause
exit /b 0

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
echo Please run 0-Install-Dependencies after installing Python 3.12+.
echo.
pause
exit /b 1
