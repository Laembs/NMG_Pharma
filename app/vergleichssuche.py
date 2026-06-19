"""SP31 / V1.1 SP2: Vergleichs-Such-Tool.

Sucht uebergreifend in den Wissens-Tabellen:
- tbl_austauschdatenbank   (pzn_alt, pzn_nmg, artikel_alt, artikel_nmg, freitext_austausch)
- tbl_nmg_stamm            (pzn, artikelname, herstellerkuerzel, apu)
- tbl_artikelstamm         (pzn, artikel, df, pck, herst, ek)
- tbl_pzn_basisdaten       (pzn, artikelname, herstellerkuerzel, df, pck)
- tbl_wirkstoff_staerke    (pzn, wirkstoff, staerke) - V1.1 SP2

Eingabe ist entweder eine PZN (Ziffern, optional mit '/N'-Suffix) oder ein
Textbegriff. Bei Text wird parallel gesucht in:
  Artikelname, Wirkstoff, Hersteller, Staerke (Freitext_Austausch zusaetzlich
  in der Austauschdatenbank). Treffer pro normalisierter PZN dedupliziert.

Performance: gespeicherte PZNs sind beim Import bereits auf 8 Ziffern
normalisiert. Die Suche kann deshalb direkt mit `pzn LIKE '%digits%'` arbeiten
und muss die Python-UDF pzn_norm() nicht pro Zeile aufrufen.
"""

import re
import sqlite3

from .config import DB_PATH


def _pzn_norm(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    # SP29: '/N'-Suffix der SK-Lagervariante VOR der Ziffern-Extraktion abschneiden.
    text = re.sub(r"\s*/\s*N\s*$", "", text, flags=re.IGNORECASE)
    digits = "".join(ch for ch in text if ch.isdigit())
    if not digits:
        return ""
    return digits.zfill(8)


def _connect():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def _is_pzn_input(query: str) -> bool:
    cleaned = re.sub(r"\s*/\s*N\s*$", "", query.strip(), flags=re.IGNORECASE)
    if not cleaned:
        return False
    non_digit = sum(1 for ch in cleaned if not (ch.isdigit() or ch.isspace()))
    return non_digit == 0


# V1.1 SP4: Filter-Tags in der Sucheingabe.
_KNOWN_FILTERS = {"nmg", "austausch", "schulbank", "wirkstoff", "herst"}


def _parse_query(raw: str) -> tuple[str, set[str]]:
    """Trennt die User-Eingabe in Suchbegriff und Filter-Set.

    Tokens, die mit '#' beginnen und einen bekannten Filter-Namen tragen,
    werden in das Filter-Set aufgenommen. Alles andere bleibt Suchbegriff.

    Beispiele:
        'gliv #nmg'          -> ('gliv', {'nmg'})
        '#schulbank ramipril' -> ('ramipril', {'schulbank'})
        'noris #nmg #herst'  -> ('noris', {'nmg', 'herst'})
        '#unbekannt foo'     -> ('foo', set())     # unbekannte Tags ignoriert
        'gliv'               -> ('gliv', set())
    """
    tokens = (raw or "").split()
    filters: set[str] = set()
    rest: list[str] = []
    for tok in tokens:
        if tok.startswith("#") and len(tok) > 1:
            name = tok[1:].lower()
            if name in _KNOWN_FILTERS:
                filters.add(name)
                continue
            # Unbekannte #-Tags werden verworfen, damit der User merkt,
            # dass sie keinen Effekt haben. Sonst landen sie versehentlich
            # als Such-LIKE in den Wissens-Tabellen und liefern Null-Treffer.
            continue
        rest.append(tok)
    return " ".join(rest).strip(), filters


def search_unified(query: str, limit: int = 200) -> list[dict]:
    """Liefert dedupliziert pro PZN. Jeder Eintrag:
        pzn, artikel, df, pck, herst, wirkstoff, staerke,
        ist_nmg, hat_austausch, hat_schulbank, via (Set: 'name', 'herst',
        'wirkstoff').

    V1.1 SP4: Filter-Syntax mit '#'. Beispiele:
        'gliv #nmg'               -> nur NMG-Stamm-Treffer fuer 'gliv'
        'noris #nmg #austausch'   -> NMG-Stamm UND Austausch-Eintrag
        '#schulbank #wirkstoff'   -> nur Schulbank + Wirkstoff-Match
    UND-Logik. Unbekannte #-Tags werden verworfen.
    """
    raw = (query or "").strip()
    if not raw:
        return []
    search_q, filters = _parse_query(raw)
    if len(search_q) < 3 and not filters:
        return []
    if len(search_q) < 3 and filters:
        # Nur Filter ohne Suchbegriff -> kein Match-Driver, leeres Ergebnis.
        return []

    results: dict[str, dict] = {}

    with _connect() as con:
        is_pzn = _is_pzn_input(search_q)
        pzn_digits = "".join(ch for ch in search_q if ch.isdigit())
        pzn_like = f"%{pzn_digits}%"
        name_like = f"%{search_q.lower()}%"
        name_lower = search_q.lower()

        def add(pzn_raw, artikel=None, df=None, pck=None, herst=None,
                wirkstoff=None, staerke=None, *, source="", via=None):
            pzn = _pzn_norm(pzn_raw)
            if not pzn:
                return
            row = results.setdefault(pzn, {
                "pzn": pzn,
                "artikel": "",
                "df": "",
                "pck": "",
                "herst": "",
                "wirkstoff": "",
                "staerke": "",
                "ist_nmg": False,
                "hat_austausch": False,
                "hat_schulbank": False,
                "via": set(),
            })
            if artikel and not row["artikel"]:
                row["artikel"] = str(artikel).strip()
            if df and not row["df"]:
                row["df"] = str(df).strip()
            if pck and not row["pck"]:
                row["pck"] = str(pck).strip()
            if herst and not row["herst"]:
                row["herst"] = str(herst).strip()
            if wirkstoff and not row["wirkstoff"]:
                row["wirkstoff"] = str(wirkstoff).strip()
            if staerke and not row["staerke"]:
                row["staerke"] = str(staerke).strip()
            if source == "nmg":
                row["ist_nmg"] = True
            if source == "austausch":
                row["hat_austausch"] = True
            if source == "schulbank":
                row["hat_austausch"] = True
                row["hat_schulbank"] = True
            if via:
                row["via"].add(via)

        # --- 1. tbl_austauschdatenbank (mit Quelle-Tracking V1.1 SP4) ---
        try:
            if is_pzn:
                rows = con.execute("""
                    SELECT pzn_alt, artikel_alt, pzn_nmg, artikel_nmg, quelle
                    FROM tbl_austauschdatenbank
                    WHERE status = 'aktiv'
                      AND (pzn_alt LIKE ? OR pzn_nmg LIKE ?)
                    LIMIT ?
                """, (pzn_like, pzn_like, limit)).fetchall()
            else:
                rows = con.execute("""
                    SELECT pzn_alt, artikel_alt, pzn_nmg, artikel_nmg, quelle
                    FROM tbl_austauschdatenbank
                    WHERE status = 'aktiv'
                      AND (LOWER(COALESCE(artikel_alt,'')) LIKE ?
                           OR LOWER(COALESCE(artikel_nmg,'')) LIKE ?
                           OR LOWER(COALESCE(freitext_austausch,'')) LIKE ?)
                    LIMIT ?
                """, (name_like, name_like, name_like, limit)).fetchall()
            for r in rows:
                src = "schulbank" if (r["quelle"] or "").strip().lower() == "schulbank" else "austausch"
                if r["pzn_alt"]:
                    add(r["pzn_alt"], artikel=r["artikel_alt"], source=src)
                if r["pzn_nmg"]:
                    add(r["pzn_nmg"], artikel=r["artikel_nmg"], source=src)
        except sqlite3.OperationalError:
            pass

        # --- 2. tbl_nmg_stamm ---
        try:
            if is_pzn:
                rows = con.execute(
                    "SELECT pzn, artikelname, herstellerkuerzel FROM tbl_nmg_stamm WHERE pzn LIKE ? LIMIT ?",
                    (pzn_like, limit),
                ).fetchall()
            else:
                rows = con.execute(
                    "SELECT pzn, artikelname, herstellerkuerzel FROM tbl_nmg_stamm WHERE LOWER(COALESCE(artikelname,'')) LIKE ? LIMIT ?",
                    (name_like, limit),
                ).fetchall()
            for r in rows:
                add(r["pzn"], artikel=r["artikelname"], herst=r["herstellerkuerzel"], source="nmg")
        except sqlite3.OperationalError:
            pass

        # --- 3. tbl_artikelstamm (mit Match-Source-Tracking V1.1 SP4) ---
        try:
            if is_pzn:
                rows = con.execute(
                    "SELECT pzn, artikel, df, pck, herst FROM tbl_artikelstamm WHERE pzn LIKE ? LIMIT ?",
                    (pzn_like, limit),
                ).fetchall()
            else:
                rows = con.execute(
                    """SELECT pzn, artikel, df, pck, herst FROM tbl_artikelstamm
                       WHERE LOWER(COALESCE(artikel,'')) LIKE ?
                          OR LOWER(COALESCE(herst,'')) LIKE ?
                       LIMIT ?""",
                    (name_like, name_like, limit),
                ).fetchall()
            for r in rows:
                # Tracking: welches Feld hat den Match getrieben? (#herst-Filter)
                via_set = []
                if not is_pzn:
                    if name_lower in (r["artikel"] or "").lower():
                        via_set.append("name")
                    if name_lower in (r["herst"] or "").lower():
                        via_set.append("herst")
                for v in via_set or [None]:
                    add(r["pzn"], artikel=r["artikel"], df=r["df"], pck=r["pck"], herst=r["herst"], via=v)
        except sqlite3.OperationalError:
            pass

        # --- 4. tbl_pzn_basisdaten (mit Match-Source-Tracking V1.1 SP4) ---
        try:
            if is_pzn:
                rows = con.execute(
                    "SELECT pzn, artikelname, herstellerkuerzel, df, pck FROM tbl_pzn_basisdaten WHERE pzn LIKE ? LIMIT ?",
                    (pzn_like, limit),
                ).fetchall()
            else:
                rows = con.execute(
                    """SELECT pzn, artikelname, herstellerkuerzel, df, pck FROM tbl_pzn_basisdaten
                       WHERE LOWER(COALESCE(artikelname,'')) LIKE ?
                          OR LOWER(COALESCE(herstellerkuerzel,'')) LIKE ?
                       LIMIT ?""",
                    (name_like, name_like, limit),
                ).fetchall()
            for r in rows:
                via_set = []
                if not is_pzn:
                    if name_lower in (r["artikelname"] or "").lower():
                        via_set.append("name")
                    if name_lower in (r["herstellerkuerzel"] or "").lower():
                        via_set.append("herst")
                for v in via_set or [None]:
                    add(r["pzn"], artikel=r["artikelname"], df=r["df"], pck=r["pck"],
                        herst=r["herstellerkuerzel"], via=v)
        except sqlite3.OperationalError:
            pass

        # --- 5. tbl_wirkstoff_staerke (V1.1 SP2; via='wirkstoff' V1.1 SP4) ---
        # Wirkstoff- und Staerke-Suche; PZN-Suche ist hier ueberfluessig, weil
        # die Stammtabellen die PZN bereits abdecken. Nur fuer Text-Eingaben.
        if not is_pzn:
            try:
                rows = con.execute(
                    """SELECT pzn, wirkstoff, staerke FROM tbl_wirkstoff_staerke
                       WHERE LOWER(COALESCE(wirkstoff,'')) LIKE ?
                          OR LOWER(COALESCE(staerke,'')) LIKE ?
                       LIMIT ?""",
                    (name_like, name_like, limit),
                ).fetchall()
                for r in rows:
                    add(r["pzn"], wirkstoff=r["wirkstoff"], staerke=r["staerke"], via="wirkstoff")
            except sqlite3.OperationalError:
                pass

        # Nachfueller fuer Schulbank: wenn die initiale Austausch-Query den
        # Treffer nicht via Schulbank-Eintrag gefunden hat (sondern z.B. via
        # Artikelname-Match aus tbl_artikelstamm), wissen wir die Quelle nicht.
        # Wenn Filter '#schulbank' gesetzt ist, holen wir die Quellen fuer alle
        # bisherigen Treffer in einer Sammelabfrage nach.
        if "schulbank" in filters and results:
            pzns = list(results.keys())
            placeholders = ",".join("?" * len(pzns))
            try:
                rows = con.execute(
                    f"""SELECT pzn_alt, pzn_nmg, quelle FROM tbl_austauschdatenbank
                        WHERE status='aktiv'
                          AND (pzn_alt IN ({placeholders}) OR pzn_nmg IN ({placeholders}))""",
                    pzns + pzns,
                ).fetchall()
                for r in rows:
                    if (r["quelle"] or "").strip().lower() != "schulbank":
                        continue
                    for p in (r["pzn_alt"], r["pzn_nmg"]):
                        if p and p in results:
                            results[p]["hat_schulbank"] = True
                            results[p]["hat_austausch"] = True
            except sqlite3.OperationalError:
                pass

        # Fehlende Felder pro Treffer aus den Stammtabellen nachfuellen (PK-Lookup, billig).
        # Nur fuer Treffer die mindestens ein Feld noch leer haben.
        for pzn, row in results.items():
            if (row["artikel"] and row["df"] and row["pck"] and row["herst"]
                    and row["wirkstoff"] and row["staerke"]):
                continue
            stamm = _lookup_stamm(con, pzn)
            if not row["artikel"] and stamm["artikel"]:
                row["artikel"] = stamm["artikel"]
            if not row["df"] and stamm["df"]:
                row["df"] = stamm["df"]
            if not row["pck"] and stamm["pck"]:
                row["pck"] = stamm["pck"]
            if not row["herst"] and stamm["herst"]:
                row["herst"] = stamm["herst"]
            if not row["wirkstoff"] and stamm.get("wirkstoff"):
                row["wirkstoff"] = stamm["wirkstoff"]
            if not row["staerke"] and stamm.get("staerke"):
                row["staerke"] = stamm["staerke"]

        # ist_nmg-Markierung fuer Treffer, die nicht via NMG-Stamm gefunden wurden.
        # Eine einzige Sammelabfrage statt einer pro Treffer.
        if results:
            pzns = list(results.keys())
            try:
                placeholders = ",".join("?" * len(pzns))
                nmg_rows = con.execute(
                    f"SELECT pzn FROM tbl_nmg_stamm WHERE pzn IN ({placeholders})",
                    pzns,
                ).fetchall()
                nmg_set = {r["pzn"] for r in nmg_rows}
                for pzn in nmg_set:
                    if pzn in results:
                        results[pzn]["ist_nmg"] = True
            except sqlite3.OperationalError:
                pass

    # V1.1 SP4: Post-Hoc-Filter mit UND-Logik anwenden.
    # via-Set wird vor der Rueckgabe entfernt (intern, fuer GUI nicht relevant).
    out = list(results.values())
    if filters:
        def keep(r) -> bool:
            if "nmg" in filters and not r["ist_nmg"]:
                return False
            if "austausch" in filters and not r["hat_austausch"]:
                return False
            if "schulbank" in filters and not r["hat_schulbank"]:
                return False
            if "wirkstoff" in filters and "wirkstoff" not in r["via"]:
                return False
            if "herst" in filters and "herst" not in r["via"]:
                return False
            return True
        out = [r for r in out if keep(r)]
    for r in out:
        r.pop("via", None)

    out.sort(key=lambda r: (not r["ist_nmg"], (r["artikel"] or "").lower(), r["pzn"]))
    return out[:limit]


def _lookup_stamm(con, pzn_norm: str) -> dict:
    """Vereinheitlichte Stammdaten zu einer PZN: artikel, df, pck, herst, apu, ek,
    wirkstoff, staerke. Reihenfolge: NMG-Stamm > Artikelstamm > Basisdaten >
    Wirkstoff/Staerke-Tabelle (erst gefundener Wert pro Feld gewinnt).
    PK-Lookups, keine UDF noetig.
    """
    info = {"artikel": "", "df": "", "pck": "", "herst": "",
            "apu": None, "ek": None, "wirkstoff": "", "staerke": ""}
    if not pzn_norm:
        return info

    try:
        r = con.execute(
            "SELECT artikelname, herstellerkuerzel, apu, taxe_ek "
            "FROM tbl_nmg_stamm WHERE pzn = ? LIMIT 1",
            (pzn_norm,),
        ).fetchone()
        if r:
            if r["artikelname"]:
                info["artikel"] = r["artikelname"]
            if r["herstellerkuerzel"]:
                info["herst"] = r["herstellerkuerzel"]
            if r["apu"] is not None:
                info["apu"] = r["apu"]
            if r["taxe_ek"] is not None:
                info["ek"] = r["taxe_ek"]
    except sqlite3.OperationalError:
        pass

    try:
        r = con.execute(
            "SELECT artikel, df, pck, herst, ek FROM tbl_artikelstamm "
            "WHERE pzn = ? LIMIT 1",
            (pzn_norm,),
        ).fetchone()
        if r:
            if not info["artikel"] and r["artikel"]:
                info["artikel"] = r["artikel"]
            if not info["herst"] and r["herst"]:
                info["herst"] = r["herst"]
            if r["df"]:
                info["df"] = r["df"]
            if r["pck"]:
                info["pck"] = r["pck"]
            if info["ek"] is None and r["ek"] is not None:
                info["ek"] = r["ek"]
    except sqlite3.OperationalError:
        pass

    try:
        r = con.execute(
            "SELECT artikelname, herstellerkuerzel, df, pck "
            "FROM tbl_pzn_basisdaten WHERE pzn = ? LIMIT 1",
            (pzn_norm,),
        ).fetchone()
        if r:
            if not info["artikel"] and r["artikelname"]:
                info["artikel"] = r["artikelname"]
            if not info["herst"] and r["herstellerkuerzel"]:
                info["herst"] = r["herstellerkuerzel"]
            if not info["df"] and r["df"]:
                info["df"] = r["df"]
            if not info["pck"] and r["pck"]:
                info["pck"] = r["pck"]
    except sqlite3.OperationalError:
        pass

    # V1.1 SP2: Wirkstoff + Staerke.
    try:
        r = con.execute(
            "SELECT wirkstoff, staerke FROM tbl_wirkstoff_staerke WHERE pzn = ? LIMIT 1",
            (pzn_norm,),
        ).fetchone()
        if r:
            if r["wirkstoff"]:
                info["wirkstoff"] = r["wirkstoff"]
            if r["staerke"]:
                info["staerke"] = r["staerke"]
    except sqlite3.OperationalError:
        pass

    return info


def get_pzn_details(pzn: str) -> dict:
    """Kompakte Karte fuer die selektierte PZN plus Austausch-Alternativen,
    jede schon mit DF/PCK/Hersteller aus den Stammtabellen angereichert.
    """
    pzn_q = _pzn_norm(pzn)
    out = {
        "pzn": pzn_q,
        "info": {"artikel": "", "df": "", "pck": "", "herst": "", "apu": None, "ek": None},
        "ist_nmg": False,
        "nmg_rabatt": None,
        "alternativen": [],
    }
    if not pzn_q:
        return out

    with _connect() as con:
        out["info"] = _lookup_stamm(con, pzn_q)

        try:
            r = con.execute(
                "SELECT 1 FROM tbl_nmg_stamm WHERE pzn = ? LIMIT 1",
                (pzn_q,),
            ).fetchone()
            out["ist_nmg"] = bool(r)
        except sqlite3.OperationalError:
            pass

        try:
            r = con.execute(
                "SELECT rabatt FROM nmg_rabatte WHERE nmg_pzn = ? LIMIT 1",
                (pzn_q,),
            ).fetchone()
            if r and r[0] is not None:
                out["nmg_rabatt"] = r[0]
        except sqlite3.OperationalError:
            pass

        try:
            rows = con.execute("""
                SELECT id, pzn_alt, artikel_alt, pzn_nmg, artikel_nmg,
                       freitext_austausch, quelle, erstellt_am
                FROM tbl_austauschdatenbank
                WHERE status = 'aktiv'
                  AND (pzn_alt = ? OR pzn_nmg = ?)
                ORDER BY erstellt_am DESC
            """, (pzn_q, pzn_q)).fetchall()
        except sqlite3.OperationalError:
            rows = []

        seen = set()
        for r in rows:
            pzn_alt_n = _pzn_norm(r["pzn_alt"])
            pzn_nmg_n = _pzn_norm(r["pzn_nmg"])
            if pzn_alt_n == pzn_q:
                alt_pzn = pzn_nmg_n
                alt_name_fallback = r["artikel_nmg"] or ""
            else:
                alt_pzn = pzn_alt_n
                alt_name_fallback = r["artikel_alt"] or ""

            key = alt_pzn or f"freitext:{(r['freitext_austausch'] or '').lower()}"
            if key in seen:
                continue
            seen.add(key)

            stamm = _lookup_stamm(con, alt_pzn) if alt_pzn else {
                "artikel": "", "df": "", "pck": "", "herst": "", "apu": None, "ek": None,
            }
            ist_alt_nmg = False
            if alt_pzn:
                try:
                    rn = con.execute(
                        "SELECT 1 FROM tbl_nmg_stamm WHERE pzn = ? LIMIT 1",
                        (alt_pzn,),
                    ).fetchone()
                    ist_alt_nmg = bool(rn)
                except sqlite3.OperationalError:
                    pass

            out["alternativen"].append({
                "pzn": alt_pzn,
                "artikel": stamm["artikel"] or alt_name_fallback,
                "df": stamm["df"],
                "pck": stamm["pck"],
                "herst": stamm["herst"],
                "freitext": r["freitext_austausch"] or "",
                "quelle": r["quelle"] or "",
                "erstellt_am": (r["erstellt_am"] or "")[:10],
                "ist_nmg": ist_alt_nmg,
            })

    return out
