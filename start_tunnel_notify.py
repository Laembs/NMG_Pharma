"""
Startet den Cloudflare-Tunnel und schickt die neue oeffentliche Adresse
automatisch per WhatsApp (ueber den Gratis-Dienst CallMeBot).

Voraussetzung: Datei whatsapp_config.txt im selben Ordner mit:
    phone=49170XXXXXXX        (deine Nummer, Laendervorwahl ohne + und ohne Leerzeichen)
    apikey=DEIN_CALLMEBOT_KEY (den bekommst du einmalig von CallMeBot, siehe Anleitung)

Starten:  python start_tunnel_notify.py
Beenden:  Strg + C
"""
import os
import re
import subprocess
import urllib.parse
import urllib.request
import urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(HERE, "whatsapp_config.txt")
LOCAL_URL = "http://localhost:8000"


def load_config():
    cfg = {}
    if os.path.exists(CONFIG):
        with open(CONFIG, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, val = line.split("=", 1)
                    cfg[key.strip().lower()] = val.strip()
    return cfg


def send_whatsapp(phone, apikey, text):
    url = ("https://api.callmebot.com/whatsapp.php?phone="
           + urllib.parse.quote(phone)
           + "&text=" + urllib.parse.quote(text)
           + "&apikey=" + urllib.parse.quote(apikey))
    try:
        with urllib.request.urlopen(url, timeout=30) as resp:
            body = resp.read().decode("utf-8", "replace").strip()
            print("   CallMeBot-Antwort:", body[:300])
            return 200 <= resp.status < 300
    except urllib.error.HTTPError as exc:
        print("   CallMeBot-Fehler:", exc.code,
              exc.read().decode("utf-8", "replace")[:300])
        return False
    except Exception as exc:
        print("   WhatsApp-Versand fehlgeschlagen:", exc)
        return False


def main():
    cfg = load_config()
    phone = cfg.get("phone")
    apikey = cfg.get("apikey")

    if not phone or not apikey:
        print("!! whatsapp_config.txt fehlt oder ist noch nicht ausgefuellt.")
        print("   Der Tunnel startet trotzdem - es wird nur keine WhatsApp gesendet.")
        print("   Anleitung: Datei whatsapp_config.txt oeffnen und ausfuellen.\n")

    print("Starte Cloudflare-Tunnel ...\n")
    proc = subprocess.Popen(
        ["cloudflared", "tunnel", "--url", LOCAL_URL],
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, bufsize=1,
    )

    url_re = re.compile(r"https://[a-z0-9-]+\.trycloudflare\.com")
    sent = False
    try:
        for line in proc.stdout:
            print(line, end="")
            if not sent:
                match = url_re.search(line)
                if match:
                    public = match.group(0)
                    print("\n>>> Oeffentliche Adresse:", public, "\n")
                    if phone and apikey:
                        ok = send_whatsapp(
                            phone, apikey,
                            "NMGone Cloud-Test ist online: " + public)
                        print("   WhatsApp gesendet.\n" if ok
                              else "   WhatsApp NICHT gesendet (siehe oben).\n")
                    sent = True
        proc.wait()
    except KeyboardInterrupt:
        proc.terminate()
        print("\nTunnel beendet.")


if __name__ == "__main__":
    main()
