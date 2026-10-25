@echo off
chcp 65001 >nul
pushd "%~dp0"
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0Schedule.ps1" -Action remove
popd
pause
