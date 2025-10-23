@echo off
REM BCC Bulk Mailer Interactive Cleaner - Windows Launcher

echo ========================================
echo BCC Bulk Mailer Interactive Cleaner
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    pause
    exit /b 1
)

echo Launching Interactive GUI...
echo.

REM Launch the interactive GUI
python gui_interactive.py

if errorlevel 1 (
    echo.
    echo ERROR: Failed to launch
    echo.
    pause
    exit /b 1
)
