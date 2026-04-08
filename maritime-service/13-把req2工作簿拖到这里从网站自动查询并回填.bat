@echo off
setlocal
pushd "%~dp0"
if "%~1"=="" (
  echo Drag a req2 workbook onto this BAT file.
  echo.
  echo This flow will:
  echo 1. Use the saved req2 site session
  echo 2. Query pending BL numbers from the real website
  echo 3. Update the workbook automatically
  echo.
  pause
  exit /b 1
)
call :resolve_python
if errorlevel 1 exit /b 1
echo Querying the real req2 website and updating the workbook...
%PY_CMD% ".\scripts\clearance_site_workflow.py" from-workbook --input "%~1" --output-dir ".\output\clearance\updated"
if errorlevel 1 goto :run_failed
echo.
echo Output folders:
echo .\output\clearance\updated
echo .\output\clearance\site_query_results
echo.
pause
exit /b 0

:run_failed
echo.
echo [ERROR] Req2 website query or workbook update failed.
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
