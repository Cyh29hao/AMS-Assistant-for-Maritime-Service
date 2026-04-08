@echo off
setlocal
pushd "%~dp0"
if "%~1"=="" (
  echo Drag a JSON request file onto this BAT file.
  echo.
  pause
  exit /b 1
)
echo Building contract from JSON...
python ".\scripts\contract_workflow.py" from-json --input "%~1" --output-dir ".\output\contracts"
echo.
echo Output folder:
echo .\output\contracts
echo.
pause
