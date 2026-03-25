import http.server
import os
import socketserver

os.chdir(os.path.dirname(os.path.abspath(__file__)))
PORT = 8765
Handler = http.server.SimpleHTTPRequestHandler
with socketserver.TCPServer(("", PORT), Handler) as httpd:
    print(f"Serving at http://localhost:{PORT}")
    httpd.serve_forever()
