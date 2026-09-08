@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Run-DpiLayoutProbe.ps1"
set EXIT_CODE=%ERRORLEVEL%
echo.
if not "%EXIT_CODE%"=="0" echo Probe failed with exit code %EXIT_CODE%.
pause
exit /b %EXIT_CODE%
