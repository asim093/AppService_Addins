@echo off
cd /d "%~dp0"

echo Stopping the application...

:: Find and kill the Node.js process
taskkill /F /IM node.exe >nul 2>&1

echo Application stopped.
pause
