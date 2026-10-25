@echo off
chcp 65001 >nul
pushd "%~dp0"
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0Code.ps1" -Mode Merge
popd
pause
