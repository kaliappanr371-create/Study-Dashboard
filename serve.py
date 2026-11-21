# Tiny local web server for the Study Dashboard (Python 3). Run: python serve.py
# Always uses the SAME address, http://localhost:8765 - your data is stored per address, so do not change the port.
import http.server, socketserver, os, mimetypes, sys
os.chdir(os.path.dirname(os.path.abspath(__file__)))
mimetypes.add_type('application/manifest+json', '.webmanifest'); mimetypes.add_type('text/javascript', '.js')
class H(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Cache-Control', 'no-cache'); super().end_headers()
    def log_message(self, *a): pass
socketserver.TCPServer.allow_reuse_address = True
with socketserver.TCPServer(('127.0.0.1', 8765), H) as s:
    print('Study Dashboard is running at http://localhost:8765  (close this window to stop)'); s.serve_forever()
