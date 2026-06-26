"""NMGone Kasse (Web) – als Desktop-Fenster, kein Browser.

Startet den FastAPI-Server (web.app) lokal auf 127.0.0.1 mit einem freien Port
und öffnet ihn in einem eigenen pywebview-Fenster mit eigenem Icon. Für den
Anwender fühlt es sich an wie ein Programm, technisch ist es die gleiche
Web-Kasse, die auch im Browser / auf dem Handy läuft (siehe docs/Plan_Kasse_Web.pdf).

Start:  python start_kasse_web.py
"""
from __future__ import annotations

import os
import socket
import sys
import threading
import time
import urllib.request

# Repo-Root in den Pfad, damit das 'web'-Paket gefunden wird.
ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

import uvicorn   # noqa: E402
import webview   # noqa: E402  (pip install pywebview)


def _freier_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _warte_auf_server(base: str, sekunden: float = 15.0) -> bool:
    ende = time.time() + sekunden
    while time.time() < ende:
        try:
            with urllib.request.urlopen(base + "/healthz", timeout=1) as r:
                if r.getcode() == 200:
                    return True
        except Exception:
            time.sleep(0.2)
    return False


def main() -> None:
    port = _freier_port()
    base = f"http://127.0.0.1:{port}"

    from web.app import app

    config = uvicorn.Config(app, host="127.0.0.1", port=port, log_level="warning")
    server = uvicorn.Server(config)

    # Server im Hintergrund-Thread; das Fenster läuft im Hauptthread.
    t = threading.Thread(target=server.run, daemon=True)
    t.start()

    if not _warte_auf_server(base):
        print("Server nicht erreichbar – Abbruch.", file=sys.stderr)
        server.should_exit = True
        sys.exit(1)

    webview.create_window(
        "NMGone Kasse",
        url=base + "/kasse",   # nach Login leitet das Gate hierher zurück
        width=1180,
        height=760,
        min_size=(960, 600),
    )
    icon = os.path.join(ROOT, "assets", "kasse.ico")
    try:
        webview.start(icon=icon if os.path.exists(icon) else None)
    except TypeError:
        # ältere pywebview-Versionen kennen den icon-Parameter nicht.
        webview.start()
    finally:
        server.should_exit = True


if __name__ == "__main__":
    main()
