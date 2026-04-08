@echo off
setlocal
pushd "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$f1 = Get-ChildItem '.\docs\11-*.txt' | Select-Object -First 1; if (-not $f1) { $f1 = Get-ChildItem '.\docs\11-*' | Select-Object -First 1 }; if ($f1) { Invoke-Item $f1.FullName }; $f2 = Get-ChildItem '.\docs\12-*.txt' | Select-Object -First 1; if (-not $f2) { $f2 = Get-ChildItem '.\docs\12-*' | Select-Object -First 1 }; if ($f2) { Invoke-Item $f2.FullName }"
exit /b 0
