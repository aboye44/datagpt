#!/bin/bash
# BCC Bulk Mailer Data Cleaner - Mac/Linux Launcher
# Double-click this file to launch the GUI (after making it executable)

echo "========================================"
echo "BCC Bulk Mailer Data Cleaner"
echo "========================================"
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed"
    echo ""
    echo "Please install Python 3:"
    echo "  Mac: brew install python3"
    echo "  Ubuntu/Debian: sudo apt install python3 python3-pip"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

echo "Checking dependencies..."
echo ""

# Check if required packages are installed
python3 -c "import pandas, openpyxl, usaddress, rapidfuzz, tkinter" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Some dependencies are missing. Installing now..."
    echo ""
    pip3 install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo ""
        echo "ERROR: Failed to install dependencies"
        echo ""
        read -p "Press Enter to exit..."
        exit 1
    fi
fi

# Try to install drag-drop support (optional)
python3 -c "import tkinterdnd2" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing drag-and-drop support (optional)..."
    pip3 install tkinterdnd2 2>/dev/null || echo "Note: Drag-drop not available, but GUI will still work"
fi

echo ""
echo "Launching GUI..."
echo ""

# Launch the GUI
python3 gui.py

if [ $? -ne 0 ]; then
    echo ""
    echo "ERROR: Failed to launch GUI"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi
