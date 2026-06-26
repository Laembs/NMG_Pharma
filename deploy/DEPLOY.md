# Web-Kasse erstmalig ausrollen — `nmgkasse.pennone.de`

Stufe 1: Kasse online + auf dem Handy, mit **eigenem Login** der Web-Kasse.
ONE-SSO (Login aus ONE, eine Anmeldung wie bei Personal) ist **Stufe 2** — siehe unten.

Server = der bestehende Hetzner-Host (derselbe, auf dem `one.pennone.de` laeuft).
In den Befehlen unten steht `<SERVER-IP>` als Platzhalter fuer die echte IP (nicht ins
oeffentliche Repo schreiben). Die Server-Schritte fuehrst **du** aus (kein Self-Deploy).
Befehle als `root` bzw. mit `sudo`.

---

## 0. Voraussetzung — DNS (erledigt)
`nmgkasse.pennone.de` zeigt bereits auf die Server-IP (geprueft: <SERVER-IP> / IPv6).
Damit kann Caddy in Schritt 5 automatisch ein Zertifikat ziehen.

## 1. Code auf den Server
Vom Entwickler-PC (dieses Worktree), nur das `web/`-Paket + die Server-Requirements:

```bash
ssh root@<SERVER-IP> 'mkdir -p /opt/nmgkasse'
scp -r web                          root@<SERVER-IP>:/opt/nmgkasse/
scp deploy/requirements-server.txt  root@<SERVER-IP>:/opt/nmgkasse/
```
Danach liegt `web.app:app` unter `/opt/nmgkasse/web/app.py`.

## 2. Virtualenv + Abhaengigkeiten (auf dem Server)
```bash
cd /opt/nmgkasse
python3 -m venv .venv
.venv/bin/pip install -U pip
.venv/bin/pip install -r requirements-server.txt
```

## 3. Schreibrechte fuer den Dienst-User
Die App schreibt SQLite nach `web/data` und `web/tenants`:
```bash
chown -R www-data:www-data /opt/nmgkasse/web/data /opt/nmgkasse/web/tenants
```
(Ordner ggf. vorher anlegen: `mkdir -p /opt/nmgkasse/web/{data,tenants}`.)

## 4. systemd-Dienst
```bash
cp /opt/nmgkasse/web/../deploy/nmgkasse.service /etc/systemd/system/nmgkasse.service   # oder per scp
# WICHTIG: in der Unit WEB_SESSION_SECRET durch einen langen Zufallswert ersetzen:
python3 -c "import secrets;print(secrets.token_urlsafe(48))"
systemctl daemon-reload
systemctl enable --now nmgkasse
systemctl status nmgkasse --no-pager        # muss "active (running)" zeigen
curl -s http://127.0.0.1:8770/healthz       # -> {"status":"ok"}
```

## 5. Caddy
Den Block aus `deploy/Caddyfile.nmgkasse` in `/etc/caddy/Caddyfile` ergaenzen (neben `one.pennone.de`), dann:
```bash
caddy validate --config /etc/caddy/Caddyfile
systemctl reload caddy
```
Jetzt sollte `https://nmgkasse.pennone.de/healthz` von aussen `{"status":"ok"}` liefern.

## 6. Firma + Benutzer anlegen (eigener Kasse-Login)
Mit dem DB-Tool `nmgone_db.py` gegen die Server-DB (Pfad = `/opt/nmgkasse/web`):
```bash
python3 nmgone_db.py --web-dir /opt/nmgkasse/web firma-neu "NMG-Pharma" \
    --passwort 'STARKES-PASSWORT' --modul kasse --admin
```
(Erzeugt `platform.sqlite` + Tenant-DB, falls noch nicht vorhanden, und einen Admin-Benutzer.)
Danach `systemctl restart nmgkasse`, damit die App die neue DB sicher sieht.

## 7. Handy
`https://nmgkasse.pennone.de` oeffnen → mit Firma/Benutzer/Passwort aus Schritt 6 anmelden
→ im Browsermenue „Zum Startbildschirm hinzufuegen" (PWA-Icon, manifest+Service-Worker sind drin).

---

## Stufe 2 — Login aus ONE (SSO), spaeter
Damit die Cockpit-Kachel (bereits verdrahtet in `start_cockpit.py` → `WEB_APPS["kasse"]`)
**durchgeloggt** aufgeht, fehlt der Kasse noch der SSO-Endpunkt:

1. In `web/app.py` einen `GET /sso/cockpit` ergaenzen (Token pruefen wie `pennone/routers/auth.py`,
   `jti` einmalig verbrauchen, Session anlegen, auf `next` redirecten) und `/sso/cockpit` in
   `_OEFFENTLICH` aufnehmen.
2. Auf dem Server `PENNONE_SSO_SECRET` setzen — **derselbe Wert** wie bei ONE
   (in `nmgkasse.service` als `Environment=` und im Cockpit als Env).
3. Benutzer der Firma muessen in der Kasse-`platform.sqlite` existieren (Schritt 6), damit der
   Token-Login einen Treffer hat.

Ohne Stufe 2 landet ein Kachel-Klick auf der normalen Kasse-Login-Seite (kein Fehler, nur ein
zweiter Login).
