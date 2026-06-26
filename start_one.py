"""Pennone One (ONE) in einem eigenen App-Fenster oeffnen – wie eine lokale App.

Wird vom Cockpit als eigener Prozess gestartet. Erwartet die fertige Anmelde-URL
(mit SSO-Token) in der Umgebungsvariable ONE_SSO_URL. Ist der ONE-Server noch
nicht erreichbar, wird er hier lokal gestartet (uvicorn via Pennones run.py) und
kurz auf Bereitschaft gewartet.

Aufruf erfolgt ueber das Cockpit; direkt nutzbar ist es nur mit gesetzter
ONE_SSO_URL (sonst landet man auf der ONE-Login-Seite).
"""
import os
import socket
import subprocess
import sys
import time
from urllib.parse import urlparse

import webview  # pip install pywebview

ROOT = os.path.dirname(os.path.abspath(__file__))


def _host_port(url: str) -> tuple[str, int]:
    p = urlparse(url)
    return (p.hostname or "127.0.0.1", p.port or (443 if p.scheme == "https" else 80))


def _port_open(host: str, port: int, timeout: float = 0.5) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def _ensure_server(base_url: str, pennone_dir: str, wait_s: float = 20.0) -> bool:
    """ONE-Server sicherstellen. True, wenn erreichbar (ggf. neu gestartet)."""
    host, port = _host_port(base_url)
    if _port_open(host, port):
        return True
    run_py = os.path.join(pennone_dir, "run.py")
    if not os.path.exists(run_py):
        return False  # nur lokaler Start moeglich; zentraler Server muss laufen
    try:
        subprocess.Popen([sys.executable, run_py], cwd=pennone_dir)
    except OSError:
        return False
    deadline = time.time() + wait_s
    while time.time() < deadline:
        if _port_open(host, port):
            return True
        time.sleep(0.4)
    return False


def main():
    sso_url = os.environ.get("ONE_SSO_URL", "")
    base_url = os.environ.get("ONE_BASE_URL", "http://127.0.0.1:8765")
    pennone_dir = os.environ.get("ONE_DIR", r"C:\pennone_one")
    title = os.environ.get("ONE_TITLE", "Pennone One")
    icon = os.environ.get("ONE_ICON", os.path.join(ROOT, "assets", "NMGone.ico"))

    target = sso_url or base_url
    if not _ensure_server(base_url, pennone_dir):
        # Server nicht erreichbar -> trotzdem Fenster zeigen (mit Hinweis statt leerer Seite)
        webview.create_window(
            title,
            html=("<body style='font-family:Segoe UI;padding:40px;color:#0B2A4A'>"
                  "<h2>Pennone One ist nicht erreichbar</h2>"
                  f"<p>Server unter <code>{base_url}</code> antwortet nicht und konnte "
                  "nicht gestartet werden.</p></body>"),
            width=900, height=600)
    else:
        webview.create_window(target, url=target, width=1180, height=780, min_size=(960, 600))

    try:
        webview.start(icon=icon if os.path.exists(icon) else None)
    except TypeError:
        webview.start()


if __name__ == "__main__":
    main()
