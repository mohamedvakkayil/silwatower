#!/usr/bin/env python3
"""
Simple HTTP server to serve board details API.
Run with: python3 board_details_api.py
Then access dashboard at: http://localhost:8000/dashboard.html
"""

from http.server import HTTPServer, BaseHTTPRequestHandler
import json
import urllib.parse
from extract_board_details import extract_board_details
import os

class BoardDetailsHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path.startswith('/api/board-details'):
            # Parse query parameters
            parsed_path = urllib.parse.urlparse(self.path)
            query_params = urllib.parse.parse_qs(parsed_path.query)
            board_name = query_params.get('name', [None])[0]
            
            if not board_name:
                self.send_response(400)
                self.send_header('Content-type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({'error': 'Board name required'}).encode())
                return
            
            # Extract board details
            details = extract_board_details(board_name)
            
            # Send response
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps(details, default=str).encode())
            
        elif self.path == '/' or self.path == '/dashboard.html':
            # Serve dashboard.html
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            with open('dashboard.html', 'rb') as f:
                self.wfile.write(f.read())
                
        elif self.path.endswith('.js'):
            # Serve JavaScript files
            self.send_response(200)
            self.send_header('Content-type', 'application/javascript')
            self.end_headers()
            try:
                with open(self.path[1:], 'rb') as f:
                    self.wfile.write(f.read())
            except FileNotFoundError:
                self.send_response(404)
                self.end_headers()
                
        elif self.path.endswith('.css'):
            # Serve CSS files
            self.send_response(200)
            self.send_header('Content-type', 'text/css')
            self.end_headers()
            try:
                with open(self.path[1:], 'rb') as f:
                    self.wfile.write(f.read())
            except FileNotFoundError:
                self.send_response(404)
                self.end_headers()
                
        elif self.path.endswith('.json'):
            # Serve JSON files
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            try:
                with open(self.path[1:], 'rb') as f:
                    self.wfile.write(f.read())
            except FileNotFoundError:
                self.send_response(404)
                self.end_headers()
        else:
            # Try to serve static files
            file_path = self.path[1:] if self.path.startswith('/') else self.path
            if os.path.exists(file_path) and os.path.isfile(file_path):
                self.send_response(200)
                content_type = 'application/octet-stream'
                if file_path.endswith('.html'):
                    content_type = 'text/html'
                elif file_path.endswith('.css'):
                    content_type = 'text/css'
                elif file_path.endswith('.js'):
                    content_type = 'application/javascript'
                self.send_header('Content-type', content_type)
                self.end_headers()
                with open(file_path, 'rb') as f:
                    self.wfile.write(f.read())
            else:
                self.send_response(404)
                self.end_headers()
    
    def log_message(self, format, *args):
        # Suppress default logging
        pass

if __name__ == '__main__':
    port = 8000
    server = HTTPServer(('localhost', port), BoardDetailsHandler)
    print(f'Board Details API Server running on http://localhost:{port}')
    print(f'Access dashboard at: http://localhost:{port}/dashboard.html')
    print('Press Ctrl+C to stop')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nServer stopped')
        server.shutdown()

