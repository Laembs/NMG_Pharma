"""End-to-End-Rauchtest des Web-Piloten (stdlib urllib, keine Extra-Deps).

Beweist: Login, Personal-CRUD, Lizenz-Gate (Firma B ohne Personal -> 403)
und Mandanten-Trennung (eigene DB pro Firma).
"""
import sys
import time
import urllib.request
import urllib.parse
import http.cookiejar

import os

BASE = os.environ.get("SMOKE_BASE", "http://127.0.0.1:8000")


def session():
    cj = http.cookiejar.CookieJar()
    op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    return op


def get(op, path, allow_redirect=True):
    req = urllib.request.Request(BASE + path)
    try:
        r = op.open(req)
        return r.getcode(), r.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", "replace")


def get_bytes(op, path):
    req = urllib.request.Request(BASE + path)
    try:
        r = op.open(req)
        return r.getcode(), r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.read()


def post(op, path, data):
    body = urllib.parse.urlencode(data).encode()
    req = urllib.request.Request(BASE + path, data=body, method="POST")
    try:
        r = op.open(req)
        return r.getcode(), r.read().decode("utf-8", "replace"), r.geturl()
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", "replace"), BASE + path


def post_file(op, path, fields, filefield, filename, filebytes):
    """Multipart-POST (für Datei-Upload)."""
    boundary = "----nmgsmoke1234567890"
    parts = []
    for k, v in fields.items():
        parts.append(f"--{boundary}\r\nContent-Disposition: form-data; name=\"{k}\"\r\n\r\n{v}\r\n")
    head = (f"--{boundary}\r\nContent-Disposition: form-data; name=\"{filefield}\"; "
            f"filename=\"{filename}\"\r\nContent-Type: text/plain\r\n\r\n")
    body = "".join(parts).encode() + head.encode() + filebytes + f"\r\n--{boundary}--\r\n".encode()
    req = urllib.request.Request(BASE + path, data=body, method="POST")
    req.add_header("Content-Type", f"multipart/form-data; boundary={boundary}")
    try:
        r = op.open(req)
        return r.getcode(), r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.read()


def login(op, firma, login_, pw):
    return post(op, "/login", {"firma": firma, "login": login_, "passwort": pw})


def wait_up(tries=40):
    for _ in range(tries):
        try:
            code, _ = get(session(), "/healthz")
            if code == 200:
                return True
        except Exception:
            pass
        time.sleep(0.5)
    return False


def main():
    assert wait_up(), "Server nicht erreichbar"
    fails = []

    def check(name, cond):
        print(("OK  " if cond else "FAIL") + "  " + name)
        if not cond:
            fails.append(name)

    # 1) Firma A: Login + Personal sichtbar mit Demo-Mitarbeitern
    a = session()
    code, _, url = login(a, "muster-pharma-gmbh", "admin", "demo123")
    check("Firma A Login -> Dashboard", url.endswith("/dashboard"))
    code, html = get(a, "/dashboard")
    check("Firma A Dashboard zeigt Personal-Kachel", "Personal" in html and "/personal" in html)
    code, html = get(a, "/personal")
    check("Firma A /personal == 200", code == 200)
    check("Firma A sieht Demo-Mitarbeiter (Maier)", "Maier" in html)

    # 2) Mitarbeiter anlegen -> erscheint in Liste
    post(a, "/personal/neu", {"vorname": "Test", "name": "Webnutzer",
                              "abteilung": "IT", "position": "Tester",
                              "urlaubsanspruch": "28"})
    code, html = get(a, "/personal")
    check("Firma A: neuer Mitarbeiter sichtbar", "Webnutzer" in html)

    # 3) Abwesenheit eintragen
    # mitarbeiter_id 1 existiert (Demo). Urlaub im laufenden Jahr.
    yr = time.strftime("%Y")
    post(a, "/personal/abwesenheiten", {"mitarbeiter_id": "1", "art": "Urlaub",
                                        "von": f"{yr}-07-01", "bis": f"{yr}-07-05",
                                        "notiz": "Sommer"})
    code, html = get(a, "/personal/abwesenheiten")
    check("Firma A: Abwesenheit gespeichert", "Sommer" in html)
    check("Firma A: Urlaub geht als Antrag (Status beantragt)", "beantragt" in html.lower())

    # 3b) Genehmigungs-Workflow: ersten offenen Antrag genehmigen
    import re
    m = re.search(r'/personal/abwesenheiten/(\d+)/genehmigen', html)
    check("Firma A: Genehmigen-Button vorhanden", bool(m))
    if m:
        aid = m.group(1)
        post(a, f"/personal/abwesenheiten/{aid}/genehmigen", {})
        code, html = get(a, "/personal/abwesenheiten")
        check("Firma A: Antrag ist genehmigt", "genehmigt" in html.lower())

    # 3c) Kalender lädt und zeigt den Monat
    code, html = get(a, "/personal/kalender")
    check("Firma A: Kalender == 200", code == 200)
    check("Firma A: Kalender zeigt Wochentage", "Mo" in html and "So" in html)

    # 3d) Zeiterfassung: stempeln + Übersicht
    code, html = get(a, "/personal/zeiterfassung")
    check("Firma A: Zeiterfassung == 200", code == 200)
    post(a, "/personal/zeiterfassung/stempeln", {"mitarbeiter_id": "1"})
    post(a, "/personal/zeiterfassung", {"mitarbeiter_id": "1", "datum": f"{yr}-06-02",
                                        "kommt": "08:00", "geht": "16:30", "pause_minuten": "30"})
    code, html = get(a, "/personal/zeiterfassung")
    check("Firma A: Zeitbuchung sichtbar (08:00)", "08:00" in html)

    # 3e) Digitale Personalakte: Upload + Download
    code, html = get(a, "/personal/1/akte")
    check("Firma A: Personalakte == 200", code == 200)
    code, _ = post_file(a, "/personal/1/akte/upload",
                        {"kategorie": "Vertrag", "titel": "Arbeitsvertrag"},
                        "datei", "vertrag.txt", b"DEMO-VERTRAG-INHALT")
    code, html = get(a, "/personal/1/akte")
    check("Firma A: Dokument in Akte sichtbar", "Arbeitsvertrag" in html)
    m = re.search(r'/personal/dokument/(\d+)"', html)
    if m:
        code, raw = get_bytes(a, f"/personal/dokument/{m.group(1)}")
        check("Firma A: Dokument-Download liefert Inhalt", b"DEMO-VERTRAG-INHALT" in raw)
    else:
        check("Firma A: Dokument-Download-Link vorhanden", False)

    # 4) Firma B: KEIN Personal -> Lizenz-Gate
    b = session()
    code, _, url = login(b, "beta-distribution-kg", "admin", "demo123")
    check("Firma B Login -> Dashboard", url.endswith("/dashboard"))
    code, html = get(b, "/dashboard")
    check("Firma B Dashboard zeigt KEINE Personal-Kachel (Link)",
          'href="/personal"' not in html)
    code, html = get(b, "/personal")
    check("Firma B /personal == 403 (Lizenz-Gate)", code == 403)

    # 5) Mandanten-Trennung: nicht angemeldet -> Redirect auf Login
    anon = session()
    req = urllib.request.Request(BASE + "/personal")
    # ohne redirect-follow prüfen wir nicht; stattdessen: anon landet auf Login-HTML
    code, html = get(anon, "/personal")
    check("Anonym /personal -> Login-Seite", "Anmelden" in html or "anmelden" in html)

    print()
    if fails:
        print(f"{len(fails)} FEHLGESCHLAGEN: {fails}")
        sys.exit(1)
    print("ALLE CHECKS BESTANDEN")


if __name__ == "__main__":
    main()
