#!/bin/bash
# ESG Analysis GUI Launcher Script

echo "🌍 Launching Rapid ESG Data Insights GUI..."

# Check if Python is available
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is required but not installed."
    echo "Please install Python 3 and try again."
    exit 1
fi

# Check if required packages are available
python3 -c "import tkinter, pandas, matplotlib, seaborn" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "📦 Installing required packages..."
    pip3 install pandas numpy matplotlib seaborn requests openpyxl xlsxwriter plotly
fi

# Launch the GUI
python3 esg_gui.py

echo "👋 Thanks for using Rapid ESG Data Insights!"