@echo off
chcp 65001 >nul
pushd "%~dp0"
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0CheckEnv.ps1"
popd
pause
