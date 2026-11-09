@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONUTF8=1
py -3.14 "%~dp0solver.py"
if errorlevel 1 echo The run reported an error. See the message above for details.
echo.
echo Press any key to close this window.
pause >nul
