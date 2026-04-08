@echo off
setlocal
pushd "%~dp0"
echo Building all example contracts...
python ".\scripts\contract_workflow.py" build-examples
echo.
echo Output folder:
echo .\output\contracts
echo.
pause
