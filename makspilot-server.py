#!/usr/bin/env python3
"""
Custom HTTP server with 404 redirect for MaksPilot development
"""

import http.server
import socketserver
import os
import re
from urllib.parse import urlparse

class MaksPilotHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        # Add CORS headers for development
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', '*')
        super().end_headers()
    
    def do_GET(self):
        # Parse the URL
        parsed_path = urlparse(self.path)
        file_path = parsed_path.path.lstrip('/')
        
        # If path ends with '/', try index.html
        if file_path.endswith('/'):
            file_path += 'index.html'
        elif file_path == '':
            file_path = 'index.html'
        
        # Check if file exists
        if os.path.exists(file_path) and os.path.isfile(file_path):
            # File exists, serve it normally
            super().do_GET()
        else:
            # File doesn't exist, serve custom 404
            self.serve_custom_404()
    
    def serve_custom_404(self):
        """Serve custom 404.html page"""
        try:
            # Check if 404.html exists
            if os.path.exists('404.html'):
                with open('404.html', 'rb') as f:
                    content = f.read()
                
                self.send_response(404)
                self.send_header('Content-type', 'text/html; charset=utf-8')
                self.send_header('Content-length', str(len(content)))
                # Add cache-busting headers for 404 pages
                self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
                self.send_header('Pragma', 'no-cache')
                self.send_header('Expires', '0')
                self.end_headers()
                self.wfile.write(content)
            else:
                # Fallback to simple 404
                self.send_simple_404()
        except Exception as e:
            print(f"Error serving 404: {e}")
            self.send_simple_404()
    
    def send_simple_404(self):
        """Send a simple 404 response"""
        message = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>404 - Not Found</title>
            <style>
                body {{ 
                    font-family: 'Inter', Arial, sans-serif; 
                    background-color: #208778; 
                    color: #F9FAFB; 
                    text-align: center; 
                    padding: 50px; 
                    margin: 0;
                }}
                .error-number {{ 
                    font-size: 4rem; 
                    color: #63cfbf; 
                    font-weight: 700;
                    margin-bottom: 1rem;
                }}
                .error-title {{
                    font-size: 2rem;
                    margin-bottom: 1rem;
                }}
                .error-description {{
                    font-size: 1.1rem;
                    margin-bottom: 2rem;
                    opacity: 0.8;
                }}
                a {{ 
                    color: #63cfbf; 
                    text-decoration: none; 
                    padding: 12px 24px; 
                    border: 2px solid #63cfbf; 
                    border-radius: 6px; 
                    display: inline-block; 
                    margin: 10px; 
                    transition: all 0.3s ease;
                }}
                a:hover {{ 
                    background-color: #63cfbf; 
                    color: #208778; 
                }}
            </style>
        </head>
        <body>
            <div class="error-number">404</div>
            <h1 class="error-title">Page Not Found</h1>
            <p class="error-description">The requested page "{self.path}" could not be found.</p>
            <a href="/">Go Home</a>
            <a href="/compare/">Compare Files</a>
        </body>
        </html>
        """.encode('utf-8')
        
        self.send_response(404)
        self.send_header('Content-type', 'text/html; charset=utf-8')
        self.send_header('Content-length', str(len(message)))
        self.end_headers()
        self.wfile.write(message)
    
    def log_message(self, format, *args):
        """Custom log format"""
        print(f"[{self.date_time_string()}] {format % args}")

def run_server(port=8120):
    """Run the custom HTTP server"""
    handler = MaksPilotHTTPRequestHandler
    
    try:
        with socketserver.TCPServer(("", port), handler) as httpd:
            print(f"🚀 MaksPilot Dev Server running at http://localhost:{port}/")
            print(f"📁 Serving from: {os.getcwd()}")
            print("✨ Features:")
            print("   - Custom 404 error pages")
            print("   - Clean URL handling")
            print("   - CORS enabled for development")
            print("\n🧪 Test URLs:")
            print(f"   - http://localhost:{port}/modes/ → Custom 404")
            print(f"   - http://localhost:{port}/ru/nonexistent → Custom 404 with language detection")
            print(f"   - http://localhost:{port}/404.html → Direct 404 page")
            print("\n⭐ Stop server with Ctrl+C")
            print("-" * 60)
            
            try:
                httpd.serve_forever()
            except KeyboardInterrupt:
                print("\n\n✅ Server stopped gracefully.")
                
    except OSError as e:
        if e.errno == 48:  # Address already in use
            print(f"❌ Port {port} is already in use. Try a different port:")
            print(f"   python3 makspilot-server.py {port + 1}")
        else:
            print(f"❌ Error starting server: {e}")

if __name__ == "__main__":
    import sys
    
    # Allow custom port via command line
    port = 8120
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            print("❌ Invalid port number. Using default 8120.")
    
    run_server(port)