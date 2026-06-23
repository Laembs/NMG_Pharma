"""
Wegwerf-Test-Server fuer den Cloud-/Tunnel-Test.
Startet eine kleine Webseite auf http://localhost:8000

Spaeter ersetzt das echte Backend (FastAPI) diesen Platzhalter -
das hier dient NUR dazu, einmal den ganzen Weg PC -> Internet -> Handy zu testen.

Starten:  python cloud_test_server.py
Beenden:  Strg + C
"""
from http.server import BaseHTTPRequestHandler, HTTPServer
from datetime import datetime
import socket

PORT = 8000

PAGE = """<!doctype html>
<html lang="de"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>NMGone Cloud-Test</title>
<style>
 body{{font-family:system-ui,Arial,sans-serif;background:#f5f7fb;color:#0b4a86;
      display:flex;min-height:100vh;align-items:center;justify-content:center;margin:0}}
 .card{{background:#fff;padding:40px 48px;border-radius:16px;
        box-shadow:0 8px 30px rgba(11,74,134,.12);text-align:center;max-width:420px}}
 h1{{font-size:56px;margin:0 0 8px}}
 p{{margin:6px 0;font-size:18px}}
 .ok{{color:#1a8a3a;font-weight:bold;font-size:22px}}
 small{{color:#7a8aa0}}
</style></head>
<body><div class="card">
 <h1>&#9989;</h1>
 <p class="ok">Es funktioniert!</p>
 <p>Die NMGone-Test-Seite ist erreichbar.</p>
 <p><small>Server: {host}<br>Zeit: {time}</small></p>
</div></body></html>"""

class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        body = PAGE.format(
            host=socket.gethostname(),
            time=datetime.now().strftime("%d.%m.%Y %H:%M:%S"),
        ).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *args):
        pass  # Konsole ruhig halten

if __name__ == "__main__":
    print(f"Test-Server laeuft auf http://localhost:{PORT}")
    print("Zum Pruefen im Browser oeffnen. Beenden mit Strg + C.")
    HTTPServer(("0.0.0.0", PORT), Handler).serve_forever()
