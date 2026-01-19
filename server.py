#!/usr/bin/env python3
"""
Simple HTTP server for CSV/Excel Viewer

Usage:
    python server.py              # Runs on default port 8080
    python server.py 3000         # Runs on port 3000
    python server.py --port 5000  # Runs on port 5000
    python server.py -p 9000      # Runs on port 9000
"""

import http.server
import socketserver
import webbrowser
import os
import sys
import argparse

DEFAULT_PORT = 8080

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=os.path.dirname(os.path.abspath(__file__)) or '.', **kwargs)
    
    # Suppress request logging for cleaner output
    def log_message(self, format, *args):
        pass

def parse_args():
    parser = argparse.ArgumentParser(
        description='CSV/Excel Viewer - Local Server',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  python server.py              Start on default port (8080)
  python server.py 3000         Start on port 3000
  python server.py -p 5000      Start on port 5000
  python server.py --port 9000  Start on port 9000
        '''
    )
    parser.add_argument(
        'port_positional',
        nargs='?',
        type=int,
        help='Port number (positional argument)'
    )
    parser.add_argument(
        '-p', '--port',
        type=int,
        default=DEFAULT_PORT,
        help=f'Port number to run the server on (default: {DEFAULT_PORT})'
    )
    parser.add_argument(
        '--no-browser',
        action='store_true',
        help='Do not open browser automatically'
    )
    
    args = parser.parse_args()
    
    # Positional port takes precedence over --port flag
    if args.port_positional is not None:
        args.port = args.port_positional
    
    return args

def main():
    args = parse_args()
    port = args.port
    
    # Validate port
    if port < 1 or port > 65535:
        print(f"❌ Error: Port must be between 1 and 65535")
        sys.exit(1)
    
    # Allow socket reuse
    socketserver.TCPServer.allow_reuse_address = True
    
    try:
        with socketserver.TCPServer(("", port), Handler) as httpd:
            url = f"http://localhost:{port}"
            
            print(f"""
╔══════════════════════════════════════════════════════╗
║         📊 CSV/Excel Viewer is running!             ║
╠══════════════════════════════════════════════════════╣
║                                                      ║
║   🌐 Local:   {url:<37} ║
║                                                      ║
║   Press Ctrl+C to stop the server                    ║
║                                                      ║
╚══════════════════════════════════════════════════════╝
""")
            
            # Open browser unless --no-browser flag is set
            if not args.no_browser:
                try:
                    webbrowser.open(url)
                except:
                    pass
            
            httpd.serve_forever()
            
    except OSError as e:
        if "Address already in use" in str(e) or e.errno == 98:
            print(f"\n❌ Error: Port {port} is already in use!")
            print(f"   Try a different port: python server.py {port + 1}")
        else:
            print(f"\n❌ Error: {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n\n👋 Server stopped. Goodbye!\n")

if __name__ == "__main__":
    main()
