# Dashboard Server Setup

This server provides real-time dashboard data by reading directly from the Excel file (`e2.xlsx`).

## Installation

1. Install required Python packages:
```bash
pip install -r requirements.txt
```

Or install individually:
```bash
pip install Flask flask-cors openpyxl
```

## Running the Server

Start the server:
```bash
python3 server.py
```

By default, the server runs on port **5001** (to avoid conflicts with macOS AirPlay on port 5000).

You can specify a different port:
```bash
python3 server.py 8000
```

The dashboard will be available at:
- **Dashboard**: http://localhost:5001 (or your specified port)
- **API Endpoint**: http://localhost:5001/api/dashboard-data
- **Board Details API**: http://localhost:5001/api/board-details?name=BOARD_NAME

**Note**: If you get "Address already in use" error:
- On macOS, port 5000 is often used by AirPlay Receiver. The server defaults to port 5001 to avoid this.
- You can disable AirPlay Receiver: System Preferences -> General -> AirDrop & Handoff -> AirPlay Receiver
- Or simply use a different port: `python3 server.py 8000`

## Features

- **Real-time Updates**: Dashboard automatically refreshes every 30 seconds
- **Direct Excel Reading**: Data is read directly from `e2.xlsx` - no need to regenerate JSON files
- **Auto-refresh**: Dashboard updates automatically when Excel file changes
- **Board Details**: Click on any board to see detailed information

## How It Works

1. The server reads data directly from `e2.xlsx` TOTALLIST sheet
2. Dashboard JavaScript fetches data from `/api/dashboard-data` endpoint
3. Data refreshes automatically every 30 seconds
4. When you update the Excel file, the dashboard will show the new values on the next refresh

## Notes

- The server must be running for the dashboard to work with live data
- If the server is not running, the dashboard will fall back to embedded data (if available)
- Make sure `e2.xlsx` is in the same directory as `server.py`

