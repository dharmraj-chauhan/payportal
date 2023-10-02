from http.server import BaseHTTPRequestHandler, HTTPServer
import time

class MyHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/form2':
            # Simulate a long-running task
            time.sleep(120)  # Sleep for 2 minutes
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b'Form 2 API Response')
        else:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b'Not Found')

def handler(*args):
    server = HTTPServer(('0.0.0.0', 8080), MyHandler)
    server.serve_forever()

if __name__ == '__main__':
    handler()
