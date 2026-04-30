@echo off
setlocal
cd /d "%~dp0.."
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\scripts\reset-debug-session.ps1" %*
echo.
echo Exit code: %ERRORLEVEL%
endlocal & exit /b %ERRORLEVEL%
