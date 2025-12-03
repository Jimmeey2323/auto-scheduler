#!/bin/bash

# Schedule Validator Startup Script

echo "📊 Schedule Validator - Starting..."
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Error: Python 3 is not installed"
    echo "Please install Python 3.7 or higher"
    exit 1
fi

# Check if PyPDF2 is installed
python3 -c "import PyPDF2" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "📦 Installing required package: PyPDF2"
    pip3 install -r validator_requirements.txt
    echo ""
fi

# Start the server
echo "🚀 Starting Schedule Validator Server..."
echo ""
python3 schedule_validator.py
