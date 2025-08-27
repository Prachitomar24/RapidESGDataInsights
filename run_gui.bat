@echo off
REM ESG Analysis GUI Launcher Script for Windows

echo 🌍 Launching Rapid ESG Data Insights GUI...

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python is required but not installed.
    echo Please install Python and try again.
    pause
    exit /b 1
)

REM Check if required packages are available
python -c "import tkinter, pandas, matplotlib, seaborn" >nul 2>&1
if errorlevel 1 (
    echo 📦 Installing required packages...
    pip install pandas numpy matplotlib seaborn requests openpyxl xlsxwriter plotly
)

REM Launch the GUI
python esg_gui.py

echo 👋 Thanks for using Rapid ESG Data Insights!
pause