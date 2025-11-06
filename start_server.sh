#!/bin/bash

# Dashboard Server Startup Script

echo "Starting Dashboard Server..."
echo ""

# Check if Python 3 is available
if ! command -v python3 &> /dev/null; then
    echo "Error: Python 3 is not installed"
    exit 1
fi

# Check if required packages are installed
echo "Checking dependencies..."
python3 -c "import flask" 2>/dev/null || {
    echo "Flask not found. Installing dependencies..."
    pip3 install -r requirements.txt || {
        echo "Error: Failed to install dependencies"
        echo "Please run: pip3 install -r requirements.txt"
        exit 1
    }
}

# Check if Excel file exists
if [ ! -f "e2.xlsx" ]; then
    echo "Warning: e2.xlsx not found in current directory"
fi

echo ""
echo "Starting server..."
echo "Dashboard will be available at: http://localhost:5001"
echo "Press Ctrl+C to stop the server"
echo ""
echo "Note: If port 5001 is in use, you can specify a different port:"
echo "  python3 server.py 8000"
echo ""

# Try port 5001 first, if it fails, try 8000
python3 server.py 5001 || python3 server.py 8000

