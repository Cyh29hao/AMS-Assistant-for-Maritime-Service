@echo off
setlocal
pushd "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$f1 = Get-ChildItem '.\docs\00-*' | Where-Object { $_.Name -notlike '*acceptance-guide*' } | Select-Object -First 1; if ($f1) { Invoke-Item $f1.FullName }; $f2 = Get-ChildItem '.\docs\01-*' | Where-Object { $_.Name -notlike '*beginner-guide*' } | Select-Object -First 1; if ($f2) { Invoke-Item $f2.FullName }"
