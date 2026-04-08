@echo off
setlocal
pushd "%~dp0"
echo Creating blank Excel template...
python ".\scripts\contract_workflow.py" make-workbook-template --output ".\examples\workbooks\blank-contract-template.xlsx"
echo.
echo Template file:
echo .\examples\workbooks\blank-contract-template.xlsx
echo.
pause
