@echo off
setlocal
pushd "%~dp0"
if "%~1"=="" (
  echo Drag an Excel workbook onto this BAT file.
  echo.
  echo If you need a blank workbook first, run:
  echo 2-Create-Blank-Excel-Template
  echo.
  pause
  exit /b 1
)
echo Building contract from Excel...
python ".\scripts\contract_workflow.py" from-workbook --input "%~1" --output-dir ".\output\contracts"
echo.
echo Output folder:
echo .\output\contracts
echo.
pause
