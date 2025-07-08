@echo off
:: Change directory to the project folder (update path if needed)
cd /d "%~dp0"

:: Check if Node.js is installed
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo Node.js is not installed. Please install Node.js and try again.
    pause
    exit /b
)

:: Check if node_modules folder exists
if not exist node_modules (
    echo Installing dependencies...
    npm install
) else (
    echo Dependencies already installed.
)

:: Start the project
echo Starting the project...
npm start

pause
