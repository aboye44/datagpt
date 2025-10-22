@echo off
REM BCC Bulk Mailer Data Cleaner - Windows Launcher
REM Double-click this file to launch the GUI

echo ========================================
echo BCC Bulk Mailer Data Cleaner
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo Checking dependencies...
echo.

REM Check if required packages are installed
python -c "import pandas, openpyxl, usaddress, rapidfuzz, tkinter" >nul 2>&1
if errorlevel 1 (
    echo Some dependencies are missing. Installing now...
    echo.
    pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo ERROR: Failed to install dependencies
        echo.
        pause
        exit /b 1
    )
)

REM Try to install drag-drop support (optional)
python -c "import tkinterdnd2" >nul 2>&1
if errorlevel 1 (
    echo Installing drag-and-drop support...
    pip install tkinterdnd2
)

echo.
echo Launching GUI...
echo.

REM Launch the GUI
python gui.py

if errorlevel 1 (
    echo.
    echo ERROR: Failed to launch GUI
    echo.
    pause
    exit /b 1
)
