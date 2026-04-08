@echo off
setlocal
pushd "%~dp0"
echo Running example verification...
python ".\scripts\contract_workflow.py" verify-examples
echo.
pause
