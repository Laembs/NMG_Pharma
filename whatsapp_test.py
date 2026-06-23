"""
Schneller Test: schickt EINE Test-WhatsApp ueber CallMeBot und zeigt die Antwort.
So sehen wir genau, woran es haengt (Konfig, Nummer oder apikey).

Starten:  python whatsapp_test.py
"""
import os
import urllib.parse
import urllib.request
import urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(HERE, "whatsapp_config.txt")


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


cfg = load_config()
phone = cfg.get("phone")
apikey = cfg.get("apikey")

print("Gelesene Einstellungen aus whatsapp_config.txt:")
print("  phone  =", repr(phone))
print("  apikey =", repr(apikey))

if not phone or not apikey:
    print("\n!! phone oder apikey fehlt (oder ist noch auskommentiert).")
    print("   Bitte whatsapp_config.txt pruefen.")
    raise SystemExit(1)

url = ("https://api.callmebot.com/whatsapp.php?phone="
       + urllib.parse.quote(phone)
       + "&text=" + urllib.parse.quote("NMGone Test-Nachricht")
       + "&apikey=" + urllib.parse.quote(apikey))

print("\nSende Test-WhatsApp ...\n")
try:
    with urllib.request.urlopen(url, timeout=30) as resp:
        print("HTTP-Status:", resp.status)
        print("Antwort von CallMeBot:")
        print(resp.read().decode("utf-8", "replace")[:600])
except urllib.error.HTTPError as exc:
    print("HTTP-Fehler:", exc.code)
    print(exc.read().decode("utf-8", "replace")[:600])
except Exception as exc:
    print("Fehler:", exc)
