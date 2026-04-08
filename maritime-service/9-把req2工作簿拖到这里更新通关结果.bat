@echo off
setlocal
pushd "%~dp0"
if "%~1"=="" (
  echo Drag a req2 workbook onto this BAT file.
  echo.
  echo The workbook must contain these sheets:
  echo - Business sheet
  echo - Query sheet
  echo.
  pause
  exit /b 1
)
call :resolve_python
if errorlevel 1 exit /b 1
echo Updating req2 workbook...
%PY_CMD% ".\scripts\clearance_workflow.py" from-workbook --input "%~1" --output-dir ".\output\clearance\updated"
if errorlevel 1 goto :run_failed
echo.
echo Output folder:
echo .\output\clearance\updated
echo.
pause
exit /b 0

:run_failed
echo.
echo [ERROR] Req2 workbook update failed.
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
