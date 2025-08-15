@echo off
REM Setup weekly scheduled task for Kobo to SharePoint transfer
REM This script will run the transfer every Sunday at 2:00 AM

echo Setting up weekly scheduled task for Kobo to SharePoint transfer...

REM Get the current directory
set SCRIPT_DIR=%~dp0
set PYTHON_SCRIPT=%SCRIPT_DIR%new.py

REM Check if the Python script exists
if not exist "%PYTHON_SCRIPT%" (
    echo ❌ Error: new.py not found in %SCRIPT_DIR%
    pause
    exit /b 1
)

REM Check if .env file exists
if not exist "%SCRIPT_DIR%.env" (
    echo ❌ Error: .env file not found in %SCRIPT_DIR%
    echo Please ensure your .env file is configured with all required credentials.
    pause
    exit /b 1
)

REM Create the scheduled task
echo Creating scheduled task...
schtasks /create /tn "Kobo SharePoint Transfer" /tr "python %PYTHON_SCRIPT%" /sc weekly /d SUN /st 02:00 /f

if %errorlevel% equ 0 (
    echo ✅ Weekly scheduled task set up successfully!
    echo 📅 The script will run every Sunday at 2:00 AM
    echo.
    echo To view the task: schtasks /query /tn "Kobo SharePoint Transfer"
    echo To remove the task: schtasks /delete /tn "Kobo SharePoint Transfer" /f
    echo To run the task manually: schtasks /run /tn "Kobo SharePoint Transfer"
) else (
    echo ❌ Failed to create scheduled task
    echo Please run this script as Administrator
)

pause
