@echo off
setlocal
pushd "%~dp0"
call :resolve_python
if errorlevel 1 exit /b 1
echo Checking the saved req2 site session...
%PY_CMD% ".\scripts\clearance_site_workflow.py" check-session
if errorlevel 1 goto :run_failed
echo.
echo Check report folder:
echo .\output\clearance\site_checks
echo.
pause
exit /b 0

:run_failed
echo.
echo [ERROR] Req2 site session check failed.
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
echo Please run 0-Install-Dependencies after installing Python 3.12+.
echo.
pause
exit /b 1
