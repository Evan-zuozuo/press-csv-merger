@echo off
chcp 65001 >nul
pushd "%~dp0"
if not exist "%~dp0Code.ps1" goto NOCODE
powershell -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File "%~dp0Code.ps1" -Mode MergeAndClean -Yes -NoPause
if errorlevel 1 goto ERREXIT
popd
exit /b 0
:NOCODE
echo %date% %time% FATAL: Code.ps1 is missing. Nothing was done.>> "%~dp0_StartupError.txt"
popd
exit /b 1
:ERREXIT
echo %date% %time% FATAL: run failed, exit code %errorlevel%. See Logs folder.>> "%~dp0_StartupError.txt"
popd
exit /b 1
