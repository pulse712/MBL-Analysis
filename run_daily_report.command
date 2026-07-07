#!/bin/bash
# MLB Daily Betting Report - Mac launcher
# Double-click this file to run the daily report

echo "======================================="
echo " MLB Daily Betting Report Generator"
echo "======================================="
echo ""

# Change to the folder where this script lives
cd "$(dirname "$0")"

# Check for python3
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed."
    echo "Please install it from https://www.python.org/downloads/"
    echo ""
    read -p "Press Enter to close..."
    exit 1
fi

# Install required packages if missing
echo "Checking required packages..."
python3 -m pip install pandas openpyxl xlsxwriter requests -q

echo ""
echo "Running daily report..."
echo ""

python3 daily_report.py

echo ""
if [ $? -eq 0 ]; then
    echo "Report generated successfully!"
else
    echo "Something went wrong. Please check the error above."
fi

echo ""
read -p "Press Enter to close..."
