"""NMGone Cockpit - Prototyp (pywebview).

Eigenes Fenster mit eigenem Icon, KEIN externer Browser. Innen laeuft eine
Web-Oberflaeche (hier noch als eingebettetes HTML; spaeter das Online-Cockpit
aus web/). Ein Kachel-Klick ruft ueber die JS<->Python-Bruecke direkt Python
auf, das die jeweilige App als eigenen Prozess startet.

Design 1:1 aus app/theme.py (dunkle Sidebar + Karten, Palette als CSS-Variablen).

Features:
- NMGone-Logo in der Sidebar (aus assets/NMGone.png, skaliert + eingebettet).
- Frei einstellbare Kacheln pro Mitarbeiter: per Drag&Drop sortieren und
  ein-/ausblenden ("Anpassen"-Modus). Gespeichert je Windows-User.
- Zentrale Sprachauswahl: schreibt language.json (app/i18n) -> jede danach
  gestartete App uebernimmt die Sprache. Das fruehere Sprach-Fenster beim
  NMGone-Start entfaellt.

Start:  python start_cockpit.py
"""
import os
import re
import io
import sys
import json
import base64
import getpass
import subprocess
from datetime import datetime

# Repo-Root in den Pfad, damit das 'app'-Paket im Dev-Modus gefunden wird.
ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

import webview          # pip install pywebview
from app import theme   # zentrale Design-Quelle (Palette/Schrift)
from app import i18n     # zentrale Sprache (DE/EN/SK/CZ) – wie in den lokalen Apps
import cockpit_auth      # Login gegen Pennone One (Cache + woechentliche Pruefung)
import cockpit_sso       # Einmal-Token, damit ONE ohne erneuten Login aufgeht

i18n.load_language()

# Kachel-Schluessel -> Start-Skript (eigener Prozess, eigenes Icon).
APP_STARTER = {
    "analysen":    "start_nmgone.py",   # NMGone-Modul "Analysen" (Vertrieb)
    "kunden":      "start_kunden.py",
    "kasse":       "start_kasse.py",    # PC = volle Desktop-Kasse (Web-Kasse nur fuers Handy)
    "faktura":     "start_faktura.py",
    "gdp":         "start_gdp.py",
    "einkauf":     "start_einkauf.py",
    "buchhaltung": "start_buchhaltung.py",
    "meldungen":   "start_meldungen.py",
    "parameter":   "start_parameter.py",
    "hilfe":       "start_hilfe.py",
}

# Web-Apps (Pennone) – oeffnen in eigenem Fenster, angemeldet per SSO-Token.
# Die lokalen Fallback-Starter (start_personal.py, start_kasse.py) bleiben im
# Repo, sind aber nicht mehr verkachelt.
#   base_url: Ziel-Programm (Default = ONE aus cockpit_config.json).
#   next:     Zielbereich nach der Anmeldung.
#   dir:      lokales Repo fuer Notstart (nur same-machine); "" = nie lokal starten.
WEB_APPS = {
    "personal": {"title": "Personal · Pennone One", "next": "/personal"},
}


def _userdata_base():
    from app.config import USERDATA_ROOT, BASE_DIR
    return str(USERDATA_ROOT) if USERDATA_ROOT else str(BASE_DIR)


class CockpitApi:
    """Wird der Web-Oberflaeche als 'pywebview.api' bereitgestellt. Jede
    oeffentliche Methode ist aus JavaScript aufrufbar (liefert ein Promise)."""

    def __init__(self):
        self.user = None
        self._window = None   # wird in main() gesetzt; fuer Live-Sprachwechsel

    # -- Login (Pennone One als Quelle, lokaler Cache) ----------------------
    def login_prefill(self):
        """Feste Firma (NMG-Pharma) + zuletzt benutztes Login fuers Formular."""
        return {"firma": cockpit_auth.configured_firma(), "last": cockpit_auth.last_login()}

    def login(self, firma, login_name, password):
        res = cockpit_auth.login(firma, login_name, password)
        if res.get("ok"):
            self.user = res
        if res.get("msg"):
            res["msg"] = _t(res["msg"])
        return res

    def logout(self):
        self.user = None
        return {"ok": True}

    # -- Apps starten --------------------------------------------------------
    def start_app(self, key):
        if not self.user:
            return {"ok": False, "msg": _t("Bitte zuerst anmelden.")}
        if key in WEB_APPS:
            return self._start_web_app(key)
        script = APP_STARTER.get(key)
        if not script:
            return {"ok": False, "msg": f"Unbekannte App: {key}"}
        path = os.path.join(ROOT, script)
        if not os.path.exists(path):
            return {"ok": False, "msg": f"Start-Skript fehlt: {script}"}
        try:
            subprocess.Popen([sys.executable, path], cwd=ROOT)
            return {"ok": True, "msg": f"{key} gestartet."}
        except Exception as exc:  # pragma: no cover - Prototyp
            return {"ok": False, "msg": f"Start fehlgeschlagen: {exc}"}

    def _start_web_app(self, key):
        """ONE-Web-App in eigenem Fenster oeffnen – angemeldet per SSO-Token.

        Der Login lief bereits im Cockpit; wir signieren Firma+Login zu einem
        kurzlebigen Einmal-Token und uebergeben die fertige Anmelde-URL an
        start_one.py. Kein Passwort wird gespeichert oder weitergereicht.
        """
        cfg = WEB_APPS[key]
        base_url = cfg.get("base_url") or cockpit_sso.one_base_url()
        try:
            url = cockpit_sso.sso_url(self.user["firma"], self.user["login"],
                                      next_path=cfg.get("next"), base_url=base_url)
        except Exception as exc:  # pragma: no cover - Prototyp
            return {"ok": False, "msg": f"SSO fehlgeschlagen: {exc}"}
        env = dict(os.environ)
        env["ONE_SSO_URL"] = url
        env["ONE_BASE_URL"] = base_url
        # Notstart eines lokalen Servers nur, wenn ein Repo-Pfad gesetzt ist.
        # Remote-Apps (eigene Subdomain) setzen dir="" -> nie lokal hochfahren.
        env["ONE_DIR"] = cfg["dir"] if "dir" in cfg else cockpit_sso.pennone_dir()
        env["ONE_TITLE"] = cfg["title"]
        env["ONE_ICON"] = os.path.join(ROOT, "assets", "NMGone.ico")
        try:
            subprocess.Popen([sys.executable, os.path.join(ROOT, "start_one.py")],
                             cwd=ROOT, env=env)
            return {"ok": True, "msg": f"{cfg['title']} {_t('wird geöffnet …')}"}
        except Exception as exc:  # pragma: no cover - Prototyp
            return {"ok": False, "msg": f"Start fehlgeschlagen: {exc}"}

    # -- Kachel-Layout je Mitarbeiter ---------------------------------------
    def _layout_path(self):
        user = re.sub(r"\W+", "_", getpass.getuser() or "default") or "default"
        return os.path.join(_userdata_base(), f"cockpit_layout_{user}.json")

    def get_layout(self):
        try:
            p = self._layout_path()
            if os.path.exists(p):
                with open(p, encoding="utf-8") as fh:
                    return json.load(fh)
        except Exception:
            pass
        return None

    def save_layout(self, layout):
        try:
            p = self._layout_path()
            os.makedirs(os.path.dirname(p), exist_ok=True)
            with open(p, "w", encoding="utf-8") as fh:
                json.dump(layout, fh, ensure_ascii=False)
            return {"ok": True}
        except Exception as exc:  # pragma: no cover - Prototyp
            return {"ok": False, "msg": str(exc)}

    # -- Aufgaben/Meldungen (geteilte Liste) --------------------------------
    # Eine gemeinsame Liste fuer alle: manuelle Aufgaben UND System-Meldungen
    # der Fach-Apps (z.B. "Warenausgang gemeldet"). Liegt jetzt in der geteilten
    # DB (tbl_cockpit_meldungen, app/userdata.py) statt in einer lokalen JSON-
    # Datei -> an allen Arbeitsplaetzen sichtbar, sobald die DB gemeinsam liegt,
    # und mehrbenutzer-sicher (WAL + busy_timeout). Vorhandene JSON-Aufgaben
    # werden beim ersten Zugriff einmalig in die DB uebernommen.
    def _who(self):
        if self.user and self.user.get("name"):
            return self.user["name"].split()[0]
        try:
            return getpass.getuser()
        except Exception:
            return ""

    def get_todos(self):
        from app import userdata
        return userdata.list_todos()

    def add_todo(self, text):
        from app import userdata
        if not (text or "").strip():
            return {"ok": False, "todos": userdata.list_todos()}
        ok = userdata.add_todo(text, self._who())
        return {"ok": ok, "todos": userdata.list_todos()}

    def toggle_todo(self, todo_id):
        from app import userdata
        userdata.toggle_todo(todo_id, self._who())
        return {"ok": True, "todos": userdata.list_todos()}

    def delete_todo(self, todo_id):
        from app import userdata
        userdata.delete_todo(todo_id)
        return {"ok": True, "todos": userdata.list_todos()}

    # -- Hinweise: Verfall aus dem Lager ------------------------------------
    # Spiegelt die Verfall-Logik der Kasse (kasse_reports.verfall_rows):
    # abgelaufene + bald (<=90 Tage) ablaufende Chargen aus tbl_lagerbestand.
    def get_hinweise(self):
        try:
            from app import kasse_reports
            from app.config import DB_PATH
            rows = kasse_reports.verfall_rows(str(DB_PATH), warn_tage=90)
        except Exception:
            return {"ok": True, "items": [], "abgelaufen": 0, "bald": 0}
        items = [r for r in rows if r["status"] in ("abgelaufen", "bald")]
        abg = sum(1 for r in items if r["status"] == "abgelaufen")
        bald = sum(1 for r in items if r["status"] == "bald")
        return {
            "ok": True, "abgelaufen": abg, "bald": bald,
            "items": [{
                "artikel": r["artikelname"] or r["pzn"],
                "charge": r["charge"] or "—",
                "verfall": r["verfall"] or "—",
                "menge": r["menge"],
                "status": r["status"],
            } for r in items[:8]],
        }

    # -- Sprache (zentral fuer alle Apps) -----------------------------------
    def get_languages(self):
        from app.i18n import LANGUAGES, get_language, load_language
        load_language()
        return {"languages": LANGUAGES, "current": get_language()}

    def set_language(self, code):
        i18n.save_language(code)
        # Cockpit-Fenster sofort in der neuen Sprache neu rendern. Die Anmeldung
        # bleibt erhalten (self.user) -> JS springt per get_session direkt zurueck
        # ins Dashboard, kein erneuter Login.
        if self._window is not None:
            try:
                self._window.load_html(build_html())
            except Exception:  # pragma: no cover - Prototyp
                pass
        return {"ok": True, "current": i18n.get_language()}

    def get_session(self):
        """Aktueller Login (oder {ok:False}). Nach einem Reload (Sprachwechsel)
        kann das JS damit ohne neuen Login direkt ins Dashboard."""
        return self.user if self.user else {"ok": False}

    def whoami(self):
        return getpass.getuser()

    def ping(self):
        return "pong"


# ── Logo aus assets/NMGone.png skalieren + als data-URI einbetten ─────────────
def _logo_data_uri(width=200):
    try:
        from PIL import Image
        p = os.path.join(ROOT, "assets", "NMGone.png")
        im = Image.open(p)
        im.load()
        im = im.convert("RGBA")
        h = max(1, int(im.height * (width / im.width)))
        im = im.resize((width, h), Image.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii")
    except Exception:
        return ""


# ── Login-Hintergrund (NMG-Berge) als data-URI einbetten ─────────────────────
def _login_bg_data_uri():
    try:
        p = os.path.join(ROOT, "assets", "cockpit_login_bg.png")
        with open(p, "rb") as fh:
            data = fh.read()
        return "data:image/png;base64," + base64.b64encode(data).decode("ascii")
    except Exception:
        return ""


# ── Design aus theme.py -> CSS-Variablen + Palette fuer JS ───────────────────
CSSVARS = {
    "--bg": theme.BG, "--card": theme.CARD, "--card-alt": theme.CARD_ALT,
    "--border": theme.BORDER, "--divider": theme.DIVIDER, "--ink": theme.INK,
    "--muted": theme.MUTED, "--faint": theme.FAINT, "--primary": theme.PRIMARY,
    "--primary-dark": theme.PRIMARY_DARK, "--accent": theme.ACCENT,
    "--success": theme.SUCCESS, "--warning": theme.WARNING, "--danger": theme.DANGER,
    "--sidebar": theme.SIDEBAR, "--sidebar-active": theme.SIDEBAR_ACTIVE,
    "--sidebar-text": theme.SIDEBAR_TEXT, "--sidebar-muted": theme.SIDEBAR_MUTED,
    "--select": theme.SELECT_BG,
}
ROOTVARS = "".join(f"{k}:{v};" for k, v in CSSVARS.items())
PALETTE_JSON = json.dumps({
    "primary": theme.PRIMARY, "warning": theme.WARNING, "danger": theme.DANGER,
    "accent": theme.ACCENT, "sidebar": theme.SIDEBAR, "success": theme.SUCCESS,
    "faint": theme.FAINT,
})
_logo = _logo_data_uri()
BRAND_HTML = (f'<img class="logo" src="{_logo}" alt="NMGone">'
              if _logo else '<div class="brand">NMGone</div>')
_login_bg = _login_bg_data_uri()
# CSS fuer den Login-Hintergrund: Bild (mit Zoom) wenn vorhanden, sonst nur var(--bg).
LOGIN_BG_CSS = (f'#07172a url("{_login_bg}") center / cover no-repeat'
                if _login_bg else 'var(--bg)')


RAW = r"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>NMGone Cockpit</title>
<style>
  :root { /*ROOTVARS*/ }
  * { box-sizing: border-box; }
  html, body { height: 100%; }
  body { margin: 0; background: var(--bg); color: var(--ink);
         font-family: "Segoe UI", system-ui, sans-serif; -webkit-user-select: none; }
  .app { display: flex; height: 100vh; }

  /* dunkle Sidebar (wie theme.Sidebar) */
  .sidebar { width: 250px; background: var(--sidebar); color: var(--sidebar-text);
             display: flex; flex-direction: column; flex: none; }
  .logo { display: block; width: 180px; margin: 18px auto 4px; }
  .brand { color: #fff; font-size: 22px; font-weight: 700; padding: 20px 20px 0; }
  .brand-sub { color: var(--sidebar-muted); font-size: 12px; text-align: center; padding: 0 20px 12px; }
  .sb-divider { height: 1px; background: #16395C; margin: 8px 16px; }
  .nav { flex: 1; overflow-y: auto; }
  .nav-section { color: var(--sidebar-muted); font-size: 9px; font-weight: 700;
                 letter-spacing: .06em; padding: 14px 22px 4px; }
  .nav-item { display: flex; align-items: center; gap: 10px; color: var(--sidebar-text);
              font-size: 13px; padding: 10px 20px; margin: 1px 10px; border-radius: 8px; cursor: pointer; }
  .nav-item:hover { background: var(--sidebar-active); }
  .nav-item.active { background: var(--sidebar-active); color: #fff; font-weight: 600; }
  .nav-item .ic { width: 18px; text-align: center; }
  .sb-foot { color: var(--sidebar-muted); font-size: 11px; padding: 14px 22px; }

  /* Hauptbereich */
  .main { flex: 1; overflow-y: auto; padding: 18px 20px; }
  .header { display: flex; align-items: center; justify-content: space-between;
            background: var(--card); border: 1px solid var(--border); border-radius: 12px;
            padding: 14px 18px; margin-bottom: 16px; }
  .hi { display: flex; align-items: center; gap: 12px; }
  .avatar { width: 40px; height: 40px; border-radius: 50%; background: var(--select);
            color: var(--primary); display: flex; align-items: center; justify-content: center;
            font-weight: 700; font-size: 15px; }
  .hi h1 { font-size: 17px; font-weight: 700; margin: 0; color: var(--primary); }
  .hi p { font-size: 12px; color: var(--muted); margin: 2px 0 0; }
  .head-tools { display: flex; align-items: center; gap: 10px; }
  .head-tools select { border: 1px solid var(--border); border-radius: 8px; padding: 6px 8px;
                       font-size: 12px; color: var(--ink); background: var(--card); outline: none; }
  .btn { border: 1px solid var(--border); background: var(--card); color: var(--primary);
         font-size: 12px; font-weight: 700; border-radius: 8px; padding: 7px 12px; cursor: pointer; }
  .btn:hover { background: var(--card-alt); }

  #editbar { display: none; align-items: center; gap: 8px; background: var(--select);
             border: 1px solid var(--border); border-radius: 10px; padding: 8px 12px;
             margin-bottom: 12px; font-size: 12px; color: var(--primary); }

  .layout { display: grid; grid-template-columns: 1fr 260px; gap: 16px; align-items: start; }
  .section-title { font-size: 14px; font-weight: 700; color: var(--ink); margin: 2px 2px 10px; }

  .tiles { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }
  .tile { position: relative; background: var(--card); border: 1px solid var(--border);
          border-radius: 12px; padding: 18px 12px 14px; cursor: pointer; text-align: center;
          display: flex; flex-direction: column; align-items: center;
          transition: border-color .12s, transform .05s; }
  .tile:hover { border-color: var(--accent); }
  .tile:active { transform: scale(.985); }
  .tile.editing { cursor: move; }
  .tile.off { opacity: .42; }
  .tile.drag { outline: 2px dashed var(--accent); opacity: .6; }
  .tile .emoji { font-size: 30px; line-height: 1; margin-bottom: 8px; }
  .tile .tname { font-size: 14px; font-weight: 700; color: var(--ink); }
  .tile .tdesc { font-size: 11px; color: var(--muted); margin: 4px 0 12px; min-height: 28px; }
  .pill { background: var(--primary); color: #fff; font-size: 11px; font-weight: 700;
          border-radius: 8px; padding: 7px 0; width: 100%; }
  .tile:hover .pill { background: var(--primary-dark); }
  .toggle { background: var(--card-alt); border: 1px solid var(--border); color: var(--muted);
            font-size: 11px; font-weight: 700; border-radius: 8px; padding: 7px 0; width: 100%; cursor: pointer; }
  .badge { display: inline-block; font-size: 9px; font-weight: 700; color: var(--muted);
           background: var(--card-alt); border: 1px solid var(--border); border-radius: 6px;
           padding: 1px 6px; margin-bottom: 6px; }

  .card { background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 14px 16px; }
  .card h3 { font-size: 12px; font-weight: 700; color: var(--muted); text-transform: uppercase;
             letter-spacing: .04em; margin: 0 0 10px; }
  .side-card { margin-top: 16px; }
  .empty { font-size: 12px; color: var(--faint); padding: 4px 2px; }

  /* Aufgaben (Team-Todos) */
  .todo-in { display: flex; gap: 6px; margin-bottom: 10px; }
  .todo-in input { flex: 1; border: 1px solid var(--border); border-radius: 8px; padding: 7px 9px;
                   font-size: 12px; outline: none; color: var(--ink); background: var(--card); }
  .todo-in input:focus { border-color: var(--accent); }
  .todo-in button { background: var(--primary); color: #fff; border: none; border-radius: 8px;
                    width: 32px; font-size: 18px; line-height: 1; cursor: pointer; }
  .todo-in button:hover { background: var(--primary-dark); }
  .todos { display: flex; flex-direction: column; gap: 2px; max-height: 220px; overflow-y: auto; }
  .todo { display: flex; align-items: flex-start; gap: 7px; padding: 6px 4px; border-radius: 8px;
          font-size: 12px; }
  .todo:hover { background: var(--card-alt); }
  .todo .tck { cursor: pointer; color: var(--primary); font-size: 14px; line-height: 1.2; flex: none; }
  .todo .ttxt { flex: 1; color: var(--ink); display: flex; flex-direction: column; }
  .todo .tmeta { color: var(--faint); font-size: 10px; margin-top: 1px; }
  .todo.done .ttxt { color: var(--faint); text-decoration: line-through; }
  .todo .tdel { cursor: pointer; color: var(--faint); font-size: 15px; line-height: 1; flex: none;
                visibility: hidden; padding: 0 2px; }
  .todo:hover .tdel { visibility: visible; }
  .todo .tdel:hover { color: var(--danger); }

  /* Hinweise / Verfall */
  .hsum { display: flex; gap: 6px; margin-bottom: 8px; flex-wrap: wrap; }
  .hpill { font-size: 10px; font-weight: 700; border-radius: 6px; padding: 2px 7px; }
  .hpill.abg { background: #fde2e2; color: #a11; }
  .hpill.bald { background: #fff4d6; color: #8a6d00; }
  .hinweise { display: flex; flex-direction: column; gap: 4px; max-height: 240px; overflow-y: auto; }
  .hrow { font-size: 12px; padding: 6px 8px; border-radius: 8px; border-left: 3px solid var(--border); }
  .hrow.abgelaufen { background: #fde2e2; border-left-color: var(--danger); }
  .hrow.bald { background: #fff4d6; border-left-color: var(--warning); }
  .hrow .hart { display: block; font-weight: 700; color: var(--ink); }
  .hrow .hinfo { display: block; color: var(--muted); font-size: 11px; margin-top: 1px; }

  /* Login (Hero mit NMG-Berge-Hintergrund + Intro-Animation, wie ONE) */
  .login { position: relative; overflow: hidden; height: 100vh; display: flex;
           flex-direction: column; align-items: center; justify-content: center;
           gap: 22px; background: #07172a; }
  .login::before { content: ""; position: absolute; inset: 0; z-index: 0;
                   background: /*LOGINBG*/; }
  .login::after { content: ""; position: absolute; inset: 0; z-index: 1;
                  background: linear-gradient(180deg, rgba(7,23,42,.35) 0%, rgba(7,23,42,.82) 100%); }
  .login .brand-hero, .login .login-card, .login .login-foot { position: relative; z-index: 2; }
  .login-card { width: 360px; background: var(--card); border: 1px solid var(--border);
                border-radius: 14px; padding: 26px 26px 22px;
                box-shadow: 0 18px 50px rgba(0,0,0,.45); }

  /* Marke ueber der Karte (wie die Wortmarke bei ONE) */
  .brand-hero { display: flex; flex-direction: column; align-items: center; text-align: center; }
  .brand-hero .one-word { font-size: 40px; font-weight: 800; letter-spacing: -.5px;
                          color: #fff; line-height: 1; }
  .brand-hero .one-by { font-size: 12px; font-weight: 700; letter-spacing: 5px;
                        color: #9DC4FF; margin-top: 6px; }

  /* Intro: Vorhang faellt, Hintergrund zoomt, Marke + Karte blenden gestaffelt ein */
  .intro-curtain { position: fixed; inset: 0; z-index: 50; background: #07172a;
                   pointer-events: none; display: none; }
  .intro-skip { position: fixed; bottom: 18px; right: 18px; z-index: 60;
                display: inline-flex; align-items: center; gap: 8px; padding: 7px 14px;
                font: inherit; font-size: 13px; font-weight: 600; color: #fff; cursor: pointer;
                background: rgba(255,255,255,.14); border: 1px solid rgba(255,255,255,.35);
                border-radius: 999px; -webkit-backdrop-filter: blur(6px); backdrop-filter: blur(6px); }
  .intro-skip:hover { background: rgba(255,255,255,.26); }
  .intro-skip::before { content: ""; width: 8px; height: 8px; border-radius: 50%; background: #6FD3FF; }
  .login.intro-done .brand-hero, .login.intro-done .login-card, .login.intro-done .intro-skip {
    animation: none !important; opacity: 1 !important; transform: none !important; }
  .login.intro-done::before { animation: none !important; transform: none !important; }
  .login.intro-done .intro-curtain { display: none !important; }
  .login.intro .intro-curtain { display: block; animation: curtainOut 1.3s ease-out forwards; }
  .login.intro::before { animation: heroZoom 7s ease-out forwards; }
  .login.intro .brand-hero { opacity: 0; animation: introUp 1s ease-out 1.1s forwards; }
  .login.intro .login-card { opacity: 0; animation: introUp 1.1s ease-out 1.9s forwards; }
  .login.intro .intro-skip { opacity: 0; animation: introFade .6s ease-out .4s forwards; }
  .login.intro .intro-skip::before { animation: skipPulse 1.5s ease-out infinite; }
  @keyframes curtainOut { from { opacity: 1; } to { opacity: 0; } }
  @keyframes heroZoom { from { transform: scale(1.12); } to { transform: scale(1); } }
  @keyframes introUp { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: none; } }
  @keyframes introFade { from { opacity: 0; } to { opacity: 1; } }
  @keyframes skipPulse {
    0% { box-shadow: 0 0 0 0 rgba(111,211,255,.55); }
    70% { box-shadow: 0 0 0 8px rgba(111,211,255,0); }
    100% { box-shadow: 0 0 0 0 rgba(111,211,255,0); } }
  .login-card img.logo { width: 170px; margin: 0 auto 6px; }
  .login-card .brand { color: var(--primary); font-size: 22px; font-weight: 700; text-align: center; }
  .login-card .sub { color: var(--muted); font-size: 12px; text-align: center; margin-bottom: 14px; }
  .firma-fixed { text-align: center; font-size: 15px; font-weight: 700; color: var(--primary);
                 background: var(--select); border: 1px solid var(--border); border-radius: 8px;
                 padding: 8px 0; margin-bottom: 4px; }
  .login-card label { display: block; font-size: 11px; font-weight: 700; color: var(--muted);
                      text-transform: uppercase; letter-spacing: .04em; margin: 12px 2px 4px; }
  .login-card select, .login-card input { width: 100%; border: 1px solid var(--border); border-radius: 8px;
                      padding: 9px 10px; font-size: 13px; color: var(--ink); background: var(--card); outline: none; }
  .login-card select:focus, .login-card input:focus { border-color: var(--accent); }
  .login-card .go { width: 100%; margin-top: 18px; background: var(--primary); color: #fff; border: none;
                    border-radius: 9px; padding: 11px 0; font-size: 14px; font-weight: 700; cursor: pointer; }
  .login-card .go:hover { background: var(--primary-dark); }
  .login-err { color: var(--danger); font-size: 12px; min-height: 16px; margin-top: 10px; text-align: center; }
  .login-hint { color: var(--faint); font-size: 11px; text-align: center; margin-top: 12px; }

  .toast { position: fixed; left: 50%; bottom: 18px; transform: translateX(-50%);
           background: var(--sidebar); color: #fff; padding: 9px 16px; border-radius: 999px;
           font-size: 13px; opacity: 0; transition: opacity .2s; pointer-events: none; }
  .toast.show { opacity: 1; }
</style>
</head>
<body>
<div class="login intro" id="login">
  <div class="intro-curtain" aria-hidden="true"></div>
  <button type="button" class="intro-skip" id="lg-skip" aria-label="Animation überspringen">Animation überspringen</button>
  <div class="brand-hero">
    <span class="one-word">NMGone</span>
    <span class="one-by">COCKPIT</span>
  </div>
  <div class="login-card">
    <!--LOGO-->
    <div class="sub">Cockpit · Anmelden</div>
    <div class="firma-fixed" id="lg-firma-name">NMG-Pharma</div>
    <label>Benutzer</label>
    <input id="lg-login" type="text" autocomplete="username" placeholder="z. B. achefin">
    <label>Passwort</label>
    <input id="lg-pw" type="password" autocomplete="current-password" placeholder="Passwort">
    <button class="go" id="lg-go">Anmelden</button>
    <div class="login-err" id="lg-err"></div>
    <div class="login-hint">Konten kommen aus Pennone One · 1×/Woche wird online geprüft</div>
  </div>
</div>

<div class="app" id="app" style="display:none">
  <aside class="sidebar">
    <!--LOGO-->
    <div class="brand-sub">Cockpit</div>
    <div class="sb-divider"></div>
    <div class="nav">
      <div class="nav-item active"><span class="ic">&#127968;</span> Cockpit</div>
      <div class="nav-section">Arbeiten</div>
      <div class="nav-item" onclick="start('analysen')"><span class="ic">&#128202;</span> Analysen</div>
      <div class="nav-item" onclick="start('kunden')"><span class="ic">&#128199;</span> Kunden</div>
      <div class="nav-item" onclick="start('kasse')"><span class="ic">&#128722;</span> Kasse</div>
      <div class="nav-item" onclick="start('einkauf')"><span class="ic">&#128666;</span> Einkauf</div>
      <div class="nav-item" onclick="start('gdp')"><span class="ic">&#128230;</span> Wareneingang</div>
      <div class="nav-section">Mehr</div>
      <div class="nav-item" onclick="start('parameter')"><span class="ic">&#128272;</span> Parameter</div>
      <div class="nav-item" onclick="start('hilfe')"><span class="ic">&#10067;</span> Hilfe</div>
    </div>
    <div class="sb-foot">Prototyp &middot; pywebview</div>
  </aside>

  <main class="main">
    <div class="header">
      <div class="hi">
        <div class="avatar" id="avatar">–</div>
        <div>
          <h1 id="greet">Guten Tag</h1>
          <p>NMGone Cockpit &middot; eigenes Fenster, kein Browser</p>
        </div>
      </div>
      <div class="head-tools">
        <select id="lang" title="Sprache fuer alle Apps"></select>
        <button class="btn" id="btn-edit">&#9881; Anpassen</button>
      </div>
    </div>

    <div id="editbar">
      <span>&#9776; Kacheln ziehen zum Sortieren &middot; &bdquo;Ausblenden&ldquo; blendet aus. Wird automatisch gespeichert.</span>
    </div>

    <div class="layout">
      <div>
        <div class="section-title">Meine Apps</div>
        <div class="tiles" id="tiles"></div>
      </div>

      <div>
        <div class="card">
          <h3>Aufgaben &middot; Team</h3>
          <div class="todo-in">
            <input id="todo-in" type="text" placeholder="Neue Aufgabe&hellip;" maxlength="200">
            <button id="todo-add" title="Aufgabe hinzufügen">+</button>
          </div>
          <div class="todos" id="todos"></div>
        </div>
        <div class="card side-card">
          <h3>&#9888; Verfall im Blick</h3>
          <div class="hinweise" id="hinweise"></div>
        </div>
      </div>
    </div>
  </main>
</div>

<div class="toast" id="toast"></div>

<script>
  var PALETTE = /*PALETTEJSON*/;

  var ALL_TILES = [
    {k:"analysen",    e:"📊", n:"Analysen",     d:"Auswertungen für den Vertrieb.", c:PALETTE.primary, kind:"NMGone-Modul"},
    {k:"kunden",      e:"📇", n:"Kunden",       d:"Apotheken-CRM, ABC & Landkarte.",     c:PALETTE.accent,  kind:"lokal"},
    {k:"kasse",       e:"🛒", n:"Kasse",        d:"Verkauf & Wareneingang.",             c:PALETTE.warning, kind:"lokal"},
    {k:"gdp",         e:"📦", n:"Wareneingang", d:"Chargen & Retouren (GDP).",           c:"#0B6E6E",       kind:"lokal"},
    {k:"einkauf",     e:"🚚", n:"Einkauf",      d:"Beschaffung EU-Ausland.",             c:PALETTE.primary, kind:"lokal"},
    {k:"faktura",     e:"🧾", n:"Faktura",      d:"Rechnungen & Gutschriften.",          c:PALETTE.primary, kind:"lokal"},
    {k:"buchhaltung", e:"📒", n:"Buchhaltung",  d:"Export ans Steuerbüro.",         c:"#0B6E6E",       kind:"lokal"},
    {k:"personal",    e:"👥", n:"Personal",     d:"Mitarbeiter & Abwesenheiten (ONE).",  c:"#6B4FB3",       kind:"web · ONE"},
    {k:"meldungen",   e:"🔔", n:"Meldungen",    d:"Abweichungen & Kontrollen.",          c:PALETTE.danger,  kind:"lokal"},
    {k:"parameter",   e:"🔒", n:"Parameter",    d:"Berechtigungen: wer darf was.",       c:PALETTE.sidebar, kind:"lokal"},
    {k:"hilfe",       e:"❓",       n:"Hilfe",        d:"Bebildertes Handbuch.",               c:PALETTE.accent,  kind:"lokal"}
  ];
  var TMAP = {}; ALL_TILES.forEach(function(t){ TMAP[t.k] = t; });
  var state = { order: ALL_TILES.map(function(t){ return t.k; }), hidden: [] };
  var editing = false, dragKey = null, api = null;

  function persist() { if (api && api.save_layout) api.save_layout({order: state.order, hidden: state.hidden}); }

  function makeTile(t) {
    var hiddenNow = state.hidden.indexOf(t.k) >= 0;
    var d = document.createElement("div");
    d.className = "tile" + (editing ? " editing" : "") + (editing && hiddenNow ? " off" : "");
    d.setAttribute("data-k", t.k);
    d.innerHTML =
        '<div class="badge">' + t.kind + '</div>'
      + '<div class="emoji" style="color:' + t.c + '">' + t.e + '</div>'
      + '<div class="tname">' + t.n + '</div>'
      + '<div class="tdesc">' + t.d + '</div>'
      + (editing
          ? '<button class="toggle">' + (hiddenNow ? "Einblenden" : "Ausblenden") + '</button>'
          : '<div class="pill">Öffnen &#8594;</div>');
    if (editing) {
      d.setAttribute("draggable", "true");
      d.addEventListener("dragstart", function(){ dragKey = t.k; d.classList.add("drag"); });
      d.addEventListener("dragend", function(){ d.classList.remove("drag"); });
      d.addEventListener("dragover", function(e){ e.preventDefault(); });
      d.addEventListener("drop", function(e){ e.preventDefault(); reorder(dragKey, t.k); });
      d.querySelector(".toggle").addEventListener("click", function(e){ e.stopPropagation(); toggleHide(t.k); });
    } else {
      d.addEventListener("click", function(){ start(t.k); });
    }
    return d;
  }

  function reorder(from, to) {
    if (!from || from === to) return;
    var o = state.order.slice();
    var fi = o.indexOf(from); if (fi < 0) return;
    o.splice(fi, 1);
    var ti = o.indexOf(to); if (ti < 0) ti = o.length;
    o.splice(ti, 0, from);
    state.order = o; persist(); render();
  }

  function toggleHide(k) {
    var i = state.hidden.indexOf(k);
    if (i >= 0) state.hidden.splice(i, 1); else state.hidden.push(k);
    persist(); render();
  }

  function render() {
    var c = document.getElementById("tiles"); c.innerHTML = "";
    var keys = editing ? state.order
                       : state.order.filter(function(k){ return state.hidden.indexOf(k) < 0; });
    keys.forEach(function(k){ if (TMAP[k]) c.appendChild(makeTile(TMAP[k])); });
    document.getElementById("editbar").style.display = editing ? "flex" : "none";
    document.getElementById("btn-edit").innerHTML = editing ? "✓ Fertig" : "⚙ Anpassen";
  }

  function toggleEdit() { editing = !editing; render(); }

  function buildLang(r) {
    var sel = document.getElementById("lang"); if (!sel || !r) return;
    sel.innerHTML = "";
    Object.keys(r.languages).forEach(function(code){
      var o = document.createElement("option");
      o.value = code; o.textContent = r.languages[code] + " (" + code + ")";
      if (code === r.current) o.selected = true;
      sel.appendChild(o);
    });
    sel.addEventListener("change", function(){
      if (api && api.set_language)
        api.set_language(sel.value).then(function(){
          toast("Sprache gesetzt – gilt für neu gestartete Apps.");
        });
    });
  }

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, function(c){
      return {"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c];
    });
  }

  // ── Aufgaben (geteilte Team-Todos) ───────────────────────────────────────
  function renderTodos(todos) {
    var box = document.getElementById("todos");
    if (!box) return;
    if (!todos || !todos.length) {
      box.innerHTML = '<div class="empty">Noch keine Aufgaben &#8211; trag die erste ein.</div>';
      return;
    }
    box.innerHTML = todos.map(function(t){
      var meta = esc(t.by || "") + (t.ts ? " &middot; " + esc(t.ts) : "");
      return '<div class="todo' + (t.done ? ' done' : '') + '">'
        + '<span class="tck" data-id="' + t.id + '">' + (t.done ? '&#9745;' : '&#9744;') + '</span>'
        + '<span class="ttxt">' + esc(t.text) + '<span class="tmeta">' + meta + '</span></span>'
        + '<span class="tdel" data-id="' + t.id + '" title="Löschen">&times;</span>'
        + '</div>';
    }).join("");
    box.querySelectorAll(".tck").forEach(function(el){
      el.addEventListener("click", function(){
        api.toggle_todo(Number(el.getAttribute("data-id"))).then(function(r){ renderTodos(r.todos); });
      });
    });
    box.querySelectorAll(".tdel").forEach(function(el){
      el.addEventListener("click", function(){
        api.delete_todo(Number(el.getAttribute("data-id"))).then(function(r){ renderTodos(r.todos); });
      });
    });
  }

  function addTodo() {
    var inp = document.getElementById("todo-in");
    var txt = (inp.value || "").trim();
    if (!txt || !api || !api.add_todo) return;
    api.add_todo(txt).then(function(r){ if (r.ok) { inp.value = ""; } renderTodos(r.todos); });
  }

  // ── Hinweise / Verfall (aus dem Lager) ───────────────────────────────────
  function renderHinweise(r) {
    var box = document.getElementById("hinweise");
    if (!box) return;
    if (!r || !r.items || !r.items.length) {
      box.innerHTML = '<div class="empty">Keine Verfall-Warnungen &#128077;</div>';
      return;
    }
    var head = '<div class="hsum">'
      + (r.abgelaufen ? '<span class="hpill abg">' + r.abgelaufen + ' abgelaufen</span>' : '')
      + (r.bald ? '<span class="hpill bald">' + r.bald + ' bald</span>' : '')
      + '</div>';
    box.innerHTML = head + r.items.map(function(it){
      return '<div class="hrow ' + it.status + '">'
        + '<span class="hart">' + esc(it.artikel) + '</span>'
        + '<span class="hinfo">Charge ' + esc(it.charge) + ' &middot; ' + esc(it.verfall)
        + ' &middot; ' + it.menge + ' St.</span>'
        + '</div>';
    }).join("");
  }

  function toast(msg) {
    var t = document.getElementById("toast");
    t.textContent = msg; t.classList.add("show");
    clearTimeout(window._tt);
    window._tt = setTimeout(function(){ t.classList.remove("show"); }, 2400);
  }

  function start(key) {
    if (editing) return;
    pywebview.api.start_app(key).then(function(res){
      toast(res.ok ? ("✓ " + res.msg) : ("✗ " + res.msg));
    });
  }

  function initials(name) {
    var p = (name || "").trim().split(/\s+/);
    if (!p[0]) return "–";
    return (p.length > 1 ? p[0][0] + p[p.length - 1][0] : p[0].slice(0, 2)).toUpperCase();
  }

  function initDashboard() {
    if (api && api.get_languages) api.get_languages().then(buildLang);
    if (api && api.get_todos) api.get_todos().then(renderTodos);
    if (api && api.get_hinweise) api.get_hinweise().then(renderHinweise);
    // Aufgaben/Meldungen automatisch nachladen, damit Eintraege anderer
    // Arbeitsplaetze (z.B. ein gemeldeter Warenausgang) ohne Neuladen erscheinen.
    if (window._todoPoll) clearInterval(window._todoPoll);
    window._todoPoll = setInterval(function(){
      if (api && api.get_todos) api.get_todos().then(renderTodos);
    }, 20000);
    if (api && api.get_layout) {
      api.get_layout().then(function(saved){
        if (saved && saved.order) {
          var valid = saved.order.filter(function(k){ return TMAP[k]; });
          ALL_TILES.forEach(function(t){ if (valid.indexOf(t.k) < 0) valid.push(t.k); });
          state.order = valid;
          state.hidden = (saved.hidden || []).filter(function(k){ return TMAP[k]; });
        }
        render();
      });
    } else {
      render();
    }
  }

  function showApp(res) {
    document.getElementById("greet").textContent =
        "Guten Tag, " + (res.name ? res.name.split(/\s+/)[0] : "");
    document.getElementById("avatar").textContent = initials(res.name);
    document.getElementById("login").style.display = "none";
    document.getElementById("app").style.display = "flex";
    initDashboard();
  }

  var FIRMA_SLUG = "";

  function doLogin() {
    var err = document.getElementById("lg-err");
    err.textContent = "";
    var login = document.getElementById("lg-login").value;
    var pw = document.getElementById("lg-pw").value;
    if (!login || !pw) { err.textContent = "Benutzer und Passwort eingeben."; return; }
    api.login(FIRMA_SLUG, login, pw).then(function(res){
      if (res.ok) { if (res.msg) toast(res.msg); showApp(res); }
      else { err.textContent = res.msg || "Anmeldung fehlgeschlagen."; }
    });
  }

  function bootLogin() {
    api = (window.pywebview && window.pywebview.api) ? window.pywebview.api : null;
    if (!api) return;
    api.login_prefill().then(function(p){
      if (p.firma) {
        FIRMA_SLUG = p.firma.slug || "";
        document.getElementById("lg-firma-name").textContent = p.firma.name || "NMG-Pharma";
      }
      if (p.last) document.getElementById("lg-login").value = p.last.login || "";
      document.getElementById((p.last && p.last.login) ? "lg-pw" : "lg-login").focus();
    });
  }

  document.getElementById("btn-edit").addEventListener("click", toggleEdit);
  document.getElementById("todo-add").addEventListener("click", addTodo);
  document.getElementById("todo-in").addEventListener("keydown", function(e){ if (e.key === "Enter") addTodo(); });
  document.getElementById("lg-go").addEventListener("click", doLogin);
  document.getElementById("lg-pw").addEventListener("keydown", function(e){ if (e.key === "Enter") doLogin(); });
  document.getElementById("lg-login").addEventListener("keydown", function(e){ if (e.key === "Enter") doLogin(); });

  // Login-Intro (wie ONE): Animation ueberspringen + Skip-Button nach Ablauf ausblenden
  function skipIntro() {
    var w = document.getElementById("login");
    if (w) { w.classList.add("intro-done"); w.classList.remove("intro"); }
    var b = document.getElementById("lg-skip");
    if (b) { b.style.display = "none"; }
  }
  document.getElementById("lg-skip").addEventListener("click", skipIntro);
  window.setTimeout(function () {
    var b = document.getElementById("lg-skip");
    if (!b || !document.querySelector(".login.intro")) return;
    b.style.animation = "none";
    b.style.transition = "opacity .5s ease";
    b.style.opacity = "0";
    window.setTimeout(function () { if (b) b.style.display = "none"; }, 500);
  }, 4000);

  // Boot: nach einem Reload (z. B. Sprachwechsel) ist evtl. noch ein Login aktiv
  // -> direkt ins Dashboard, sonst Login-Maske.
  function boot() {
    api = (window.pywebview && window.pywebview.api) ? window.pywebview.api : null;
    if (!api) return;
    if (api.get_session) {
      api.get_session().then(function(s){ if (s && s.ok) showApp(s); else bootLogin(); });
    } else {
      bootLogin();
    }
  }
  if (window.pywebview && window.pywebview.api) boot();
  else window.addEventListener("pywebviewready", boot);
</script>
</body>
</html>
"""

# ── Cockpit-Uebersetzungen ───────────────────────────────────────────────────
# Das Cockpit ist eine eigenstaendige HTML/JS-Oberflaeche; die Tkinter-Auto-
# Uebersetzung der lokalen Apps greift hier nicht. Darum ein eigenes, kleines
# Woerterbuch (deutscher Anzeigetext -> SK). Es wird ins zentrale i18n
# registriert, damit es derselben Logik folgt (Fallback DE, spaeter EN/CZ).
COCKPIT_SK = {
    # Login
    "Cockpit · Anmelden": "Cockpit · Prihlásenie",
    "Benutzer": "Používateľ",
    "Passwort": "Heslo",
    "z. B. achefin": "napr. achefin",
    "Anmelden": "Prihlásiť",
    "Konten kommen aus Pennone One · 1×/Woche wird online geprüft":
        "Účty pochádzajú z Pennone One · 1×/týždeň sa overuje online",
    "Benutzer und Passwort eingeben.": "Zadajte používateľa a heslo.",
    "Anmeldung fehlgeschlagen.": "Prihlásenie zlyhalo.",
    # Login-/Auth-Meldungen (aus cockpit_auth)
    "Pennone One nicht erreichbar – lokal angemeldet.":
        "Pennone One nedostupné – prihlásené lokálne.",
    "Firma, Benutzer oder Passwort stimmen nicht.":
        "Firma, používateľ alebo heslo nesúhlasia.",
    "Bitte zuerst das Passwort in Pennone One ändern, dann hier anmelden.":
        "Najprv zmeňte heslo v Pennone One, potom sa prihláste tu.",
    "Bitte zuerst anmelden.": "Najprv sa prihláste.",
    "wird geöffnet …": "sa otvára …",
    # Sidebar / Kopf
    "Arbeiten": "Práca",
    "Analysen": "Analýzy",
    "Kunden": "Zákazníci",
    "Kasse": "Pokladňa",
    "Einkauf": "Nákup",
    "Wareneingang": "Príjem tovaru",
    "Mehr": "Viac",
    "Parameter": "Parametre",
    "Hilfe": "Pomocník",
    "Guten Tag, ": "Dobrý deň, ",
    "Guten Tag": "Dobrý deň",
    "NMGone Cockpit · eigenes Fenster, kein Browser":
        "NMGone Cockpit · vlastné okno, žiadny prehliadač",
    "Sprache fuer alle Apps": "Jazyk pre všetky aplikácie",
    "⚙ Anpassen": "⚙ Prispôsobiť",
    "✓ Fertig": "✓ Hotovo",
    "Anpassen": "Prispôsobiť",
    "Kacheln ziehen zum Sortieren &middot; &bdquo;Ausblenden&ldquo; blendet aus. Wird automatisch gespeichert.":
        "Presúvajte dlaždice na zoradenie &middot; &bdquo;Skryť&ldquo; skryje. Ukladá sa automaticky.",
    "Meine Apps": "Moje aplikácie",
    "Aufgaben &middot; Team": "Úlohy &middot; Tím",
    "Neue Aufgabe&hellip;": "Nová úloha&hellip;",
    "Aufgabe hinzufügen": "Pridať úlohu",
    "Noch keine Aufgaben &#8211; trag die erste ein.":
        "Zatiaľ žiadne úlohy &#8211; pridajte prvú.",
    "Löschen": "Odstrániť",
    "Verfall im Blick": "Exspirácia na očiach",
    "Keine Verfall-Warnungen": "Žiadne upozornenia na exspiráciu",
    # Kacheln (Beschreibungen + Namen, die nicht schon in der Sidebar stehen)
    "Auswertungen für den Vertrieb.": "Vyhodnotenia pre obchod.",
    "Apotheken-CRM, ABC & Landkarte.": "CRM lekární, ABC a mapa.",
    "Verkauf & Wareneingang (online).": "Predaj a príjem tovaru (online).",
    "Chargen & Retouren (GDP).": "Šarže a vratky (GDP).",
    "Beschaffung EU-Ausland.": "Obstarávanie v zahraničí EÚ.",
    "Rechnungen & Gutschriften.": "Faktúry a dobropisy.",
    "Export ans Steuerbüro.": "Export pre účtovníka.",
    "Mitarbeiter & Abwesenheiten (ONE).": "Zamestnanci a neprítomnosti (ONE).",
    "Abweichungen & Kontrollen.": "Odchýlky a kontroly.",
    "Berechtigungen: wer darf was.": "Oprávnenia: kto môže čo.",
    "Bebildertes Handbuch.": "Obrázková príručka.",
    "Faktura": "Fakturácia",
    "Buchhaltung": "Účtovníctvo",
    "Personal": "Personál",
    "Meldungen": "Hlásenia",
    "NMGone-Modul": "NMGone modul",
    "lokal": "lokálne",
    # Kachel-Buttons / Toasts
    "Einblenden": "Zobraziť",
    "Ausblenden": "Skryť",
    "Öffnen &#8594;": "Otvoriť &#8594;",
    "Sprache gesetzt – gilt für neu gestartete Apps.":
        "Jazyk nastavený – platí pre novo spustené aplikácie.",
}
i18n.register_translations(COCKPIT_SK, "SK")


def _t(text):
    """Uebersetzt eine Python-Meldung (Toasts/Login) in die aktive Sprache."""
    return i18n.translate(text)


def _localize(html):
    """Ersetzt die bekannten deutschen Cockpit-Strings durch die aktive Sprache.
    Bei DE unveraendert. Laengste zuerst, damit Teilstrings nicht kollidieren."""
    if i18n.get_language() == "DE":
        return html
    for de in sorted(COCKPIT_SK, key=len, reverse=True):
        tr = i18n.translate(de)
        if tr and tr != de:
            html = html.replace(de, tr)
    return html


def build_html():
    """Baut das Cockpit-HTML in der aktuell gesetzten Sprache (DE = unveraendert).

    Wichtig: zuerst die Texte uebersetzen (nur die RAW-Vorlage), DANACH Logo
    (Base64), Farben und Palette einsetzen. So kann die String-Ersetzung niemals
    den Logo-Datenstrom treffen (sonst zerschiesst ein kurzer Treffer das PNG)."""
    return (_localize(RAW)
            .replace("/*ROOTVARS*/", ROOTVARS)
            .replace("/*PALETTEJSON*/", PALETTE_JSON)
            .replace("/*LOGINBG*/", LOGIN_BG_CSS)
            .replace("<!--LOGO-->", BRAND_HTML))


HTML = build_html()


def main():
    api = CockpitApi()
    # Maximiert (Vollbild-Fenster) starten – wie die lokalen Apps. width/height
    # bleiben als Groesse fuer den wiederhergestellten Zustand erhalten.
    win_kwargs = dict(
        html=HTML,
        js_api=api,
        width=1100,
        height=720,
        min_size=(960, 600),
        maximized=True,
    )
    try:
        win = webview.create_window("NMGone Cockpit", **win_kwargs)
    except TypeError:
        # aeltere pywebview-Versionen kennen 'maximized' nicht.
        win_kwargs.pop("maximized", None)
        win = webview.create_window("NMGone Cockpit", **win_kwargs)
    api._window = win   # ermoeglicht Live-Reload bei Sprachwechsel
    icon = os.path.join(ROOT, "assets", "NMGone.ico")
    try:
        webview.start(icon=icon if os.path.exists(icon) else None)
    except TypeError:
        # aeltere pywebview-Versionen kennen den icon-Parameter nicht.
        webview.start()


if __name__ == "__main__":
    main()
