@echo off
setlocal
python "%~dp0Compare-DpiLayoutProbes.py"
set EXIT_CODE=%ERRORLEVEL%
echo.
if not "%EXIT_CODE%"=="0" echo Comparison failed with exit code %EXIT_CODE%.
pause
exit /b %EXIT_CODE%
