# -*- coding: utf-8 -*-
"""
Erzeugt eine Demo-Datenbank mit mind. 12 Demodaten pro operativem Schritt.

- Quelle:  data/nmg_startdatenbank.sqlite  (wird NUR gelesen + kopiert)
- Ziel:    data/nmg_demodatenbank.sqlite

Stammdaten (Artikelstamm, Austauschdatenbank, Biosimilar, nmg_rabatte,
Wirkstoffe, ...) werden 1:1 uebernommen. Nur die operativen Tabellen der
Module Kasse, Faktura und Personal werden mit Demodaten neu befuellt.

Idempotent: kann beliebig oft laufen, leert die operativen Zieltabellen vorher.
"""
import os
import shutil
import sqlite3
import random
from datetime import date, timedelta

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "data", "nmg_startdatenbank.sqlite")
DST = os.path.join(ROOT, "data", "nmg_demodatenbank.sqlite")

random.seed(42)
HEUTE = date(2026, 6, 24)
N = 12  # mind. 12 Datensaetze pro Schritt


def d(offset_tage):
    return (HEUTE + timedelta(days=offset_tage)).isoformat()


def ts(offset_tage):
    return (HEUTE + timedelta(days=offset_tage)).isoformat() + " 09:00:00"


# ---------------------------------------------------------------- Artikelpool
def _df_pck(con, pzn):
    """df/pck aus dem Artikelstamm nachschlagen (falls vorhanden)."""
    r = con.execute(
        "SELECT df, pck FROM tbl_artikelstamm WHERE pzn=? LIMIT 1", (pzn,)
    ).fetchone()
    return (r[0] or "FER", r[1] or "1 St") if r else ("FER", "1 St")


def artikel_pool(con):
    """Gemischter Pool aus echten Quellen: NMG-Stamm, Biosimilars, Stamm.

    Jeder Eintrag traegt seine Herkunft + den echten Rabatt (sofern vorhanden):
      herkunft  -> 'NMG' | 'Biosimilar' | 'Standard'
      rabatt    -> Prozentwert (z.B. 20.0)
      rabatt_quelle -> Klartext fuer die Belegzeile
    """
    pool = []
    seen = set()

    # 1) NMG-Artikel mit echtem Rabatt aus nmg_rabatte ---------------------
    rows = con.execute(
        """SELECT s.pzn, s.artikelname, s.apu, s.taxe_ek, r.rabatt, r.quelle
           FROM tbl_nmg_stamm s
           LEFT JOIN nmg_rabatte r ON r.nmg_pzn = s.pzn
           WHERE s.artikelname IS NOT NULL
           ORDER BY (r.rabatt IS NULL), s.pzn"""
    ).fetchall()
    for pzn, name, apu, taxe_ek, rabatt, rquelle in rows:
        if not pzn or pzn in seen:
            continue
        seen.add(pzn)
        df, pck = _df_pck(con, pzn)
        apu = float(apu) if apu else round(random.uniform(120, 950), 2)
        ek = float(taxe_ek) if taxe_ek else round(apu * 0.82, 2)
        rab = round(float(rabatt) * 100, 1) if rabatt else random.choice([5, 7.5, 10])
        # Hinweis: r.quelle aus nmg_rabatte ("Partnerkonditionen ...") ist
        # ein interner Klartext und wird NICHT uebernommen -> neutrale Label.
        pool.append(dict(
            pzn=pzn, artikel=(name or "").strip(), df=df, pck=pck,
            apu=apu, ek=ek, rabatt=rab,
            rabatt_quelle="NMG-Kondition", herkunft="NMG"))
        if sum(1 for a in pool if a["herkunft"] == "NMG") >= 22:
            break

    # 2) Biosimilars aus der Biosimilar-Wissensbasis ----------------------
    rows = con.execute(
        """SELECT bp.pzn, bp.artikelname, prod.rolle, g.wirkstoff
           FROM tbl_biosimilar_pzn bp
           JOIN tbl_biosimilar_produkt prod ON prod.id = bp.produkt_id
           JOIN tbl_biosimilar_gruppe g ON g.id = prod.gruppe_id
           WHERE bp.artikelname IS NOT NULL
           ORDER BY bp.id"""
    ).fetchall()
    random.shuffle(rows)
    for pzn, name, rolle, wirkstoff in rows:
        if not pzn or pzn in seen:
            continue
        seen.add(pzn)
        df, pck = _df_pck(con, pzn)
        # Preis aus NMG-Stamm ziehen, sonst plausibel
        pr = con.execute(
            "SELECT apu, taxe_ek FROM tbl_nmg_stamm WHERE pzn=? LIMIT 1", (pzn,)
        ).fetchone()
        apu = float(pr[0]) if pr and pr[0] else round(random.uniform(200, 1400), 2)
        ek = float(pr[1]) if pr and pr[1] else round(apu * 0.82, 2)
        # Biosimilars: hoeherer Rabatt als Original
        rab = random.choice([12, 15, 18, 22]) if rolle == "biosimilar" else \
            random.choice([3, 5])
        pool.append(dict(
            pzn=pzn, artikel=(name or "").strip(), df=df, pck=pck,
            apu=apu, ek=ek, rabatt=float(rab),
            rabatt_quelle=f"Biosimilar/{wirkstoff} ({rolle})",
            herkunft="Biosimilar"))
        if sum(1 for a in pool if a["herkunft"] == "Biosimilar") >= 22:
            break

    # 3) Standard-Artikel aus dem Artikelstamm ----------------------------
    rows = con.execute(
        """SELECT pzn, artikel, df, pck FROM tbl_artikelstamm
           WHERE artikel LIKE '%mg%' AND length(artikel) > 8
           ORDER BY treffer DESC, pzn LIMIT 400"""
    ).fetchall()
    for pzn, artikel, df, pck in rows:
        if not pzn or pzn in seen:
            continue
        seen.add(pzn)
        apu = round(random.uniform(35, 950), 2)
        pool.append(dict(
            pzn=pzn, artikel=(artikel or "").strip(),
            df=df or "FER", pck=pck or "1 St",
            apu=apu, ek=round(apu * 0.78, 2),
            rabatt=random.choice([0, 0, 3, 5]),
            rabatt_quelle="Standard", herkunft="Standard"))
        if sum(1 for a in pool if a["herkunft"] == "Standard") >= 20:
            break

    random.shuffle(pool)
    return pool


# ------------------------------------------------------------------- Personal
VORNAMEN = ["Anna", "Bernd", "Clara", "David", "Eva", "Frank", "Greta",
            "Hannes", "Ines", "Jakob", "Katrin", "Lukas", "Maria", "Nils"]
NACHNAMEN = ["Adler", "Berg", "Conrad", "Dorn", "Engel", "Forster", "Gross",
             "Hahn", "Iller", "Jung", "Kraus", "Lang", "Meyer", "Neumann"]
ABTEILUNGEN = ["Geschaeftsfuehrung", "Vertrieb", "Lager/Logistik",
               "Buchhaltung", "Qualitaet", "IT", "Einkauf", "Marketing"]
POSITIONEN = ["Geschaeftsfuehrer", "Teamleiter", "Sachbearbeiter",
              "Lagerist", "Aussendienst", "Controller", "Sachbearbeiterin",
              "Werkstudent"]


def seed_personal(con):
    cur = con.cursor()
    for t in ("tbl_mitarbeiter_wert", "tbl_mitarbeiter_feld",
              "tbl_mitarbeiter_vorgesetzter", "tbl_mitarbeiter_arbeitsbereich",
              "tbl_abwesenheit", "tbl_arbeitsbereich", "tbl_mitarbeiter"):
        cur.execute(f"DELETE FROM {t}")

    # Mitarbeiter (14)
    ma_ids = []
    for i in range(14):
        vn = VORNAMEN[i]
        nn = NACHNAMEN[i]
        abt = ABTEILUNGEN[0] if i == 0 else ABTEILUNGEN[(i % len(ABTEILUNGEN))]
        pos = POSITIONEN[0] if i == 0 else POSITIONEN[(i % len(POSITIONEN))]
        pv = 1 if pos in ("Geschaeftsfuehrer", "Teamleiter") else 0
        cur.execute(
            """INSERT INTO tbl_mitarbeiter
               (vorname,name,abteilung,position,telefon,mobil,email,
                aufgaben,aktiv,erstellt_am,bearbeiter,board_x,board_y,
                urlaubsanspruch,personalverantwortlich)
               VALUES (?,?,?,?,?,?,?,?,1,?,?,?,?,?,?)""",
            (vn, nn, abt, pos,
             f"030-1000-{100+i}", f"0151-2000-{200+i}",
             f"{vn.lower()}.{nn.lower()}@nmg-pharma.de",
             f"Verantwortlich fuer {abt}", ts(-300), "demo",
             40 + (i % 5) * 180, 40 + (i // 5) * 140, 30, pv))
        ma_ids.append(cur.lastrowid)

    # Arbeitsbereiche (12)
    bereiche = [
        ("Wareneingang", "Lager"), ("Kommissionierung", "Lager"),
        ("Versand", "Lager"), ("Key-Account", "Vertrieb"),
        ("Innendienst", "Vertrieb"), ("Aussendienst", "Vertrieb"),
        ("Debitoren", "Buchhaltung"), ("Kreditoren", "Buchhaltung"),
        ("Reklamation", "Qualitaet"), ("Audit", "Qualitaet"),
        ("Support", "IT"), ("Infrastruktur", "IT"),
    ]
    ber_ids = []
    for name, kat in bereiche:
        cur.execute("INSERT INTO tbl_arbeitsbereich (name,kategorie) VALUES (?,?)",
                    (name, kat))
        ber_ids.append(cur.lastrowid)

    # Mitarbeiter <-> Arbeitsbereich (je MA 1-3 Bereiche)
    for mid in ma_ids:
        for bid in random.sample(ber_ids, random.randint(1, 3)):
            cur.execute(
                """INSERT OR IGNORE INTO tbl_mitarbeiter_arbeitsbereich
                   (mitarbeiter_id,bereich_id,stufe) VALUES (?,?,?)""",
                (mid, bid, random.randint(1, 3)))

    # Vorgesetzten-Matrix: GF(0) -> Teamleiter -> Rest
    gf = ma_ids[0]
    leiter = [m for m, i in zip(ma_ids, range(14))
              if POSITIONEN[i % len(POSITIONEN)] == "Teamleiter"]
    for mid in ma_ids[1:]:
        chef = gf if mid in leiter else (random.choice(leiter) if leiter else gf)
        cur.execute(
            """INSERT INTO tbl_mitarbeiter_vorgesetzter
               (mitarbeiter_id,vorgesetzter_id,art,ist_primaer,erstellt_am)
               VALUES (?,?, 'disziplinarisch', 1, ?)""",
            (mid, chef, ts(-300)))

    # Abwesenheiten (14)
    arten = [("Urlaub", "Erholung"), ("Krank", "AU"),
             ("Fortbildung", "Seminar"), ("Urlaub", "Sonderurlaub")]
    for i, mid in enumerate(ma_ids):
        art, unterart = arten[i % len(arten)]
        start = -60 + i * 9
        cur.execute(
            """INSERT INTO tbl_abwesenheit
               (mitarbeiter_id,art,von,bis,notiz,unterart) VALUES (?,?,?,?,?,?)""",
            (mid, art, d(start), d(start + random.randint(1, 10)),
             f"{art} geplant", unterart))

    # Custom-Felder (EAV) + Werte
    felder = [("Personalnummer", "Text", ""),
              ("Eintrittsdatum", "Datum", ""),
              ("Fuehrerschein", "Auswahl", "Ja;Nein"),
              ("Standort", "Auswahl", "Berlin;Muenchen;Hamburg"),
              ("Kostenstelle", "Text", ""),
              ("Befristung bis", "Datum", ""),
              ("Wochenstunden", "Text", ""),
              ("Schichtmodell", "Auswahl", "Frueh;Spaet;Tag"),
              ("Notfallkontakt", "Text", ""),
              ("Ersthelfer", "Auswahl", "Ja;Nein")]
    feld_ids = []
    for i, (fn, ft, opt) in enumerate(felder):
        cur.execute(
            """INSERT INTO tbl_mitarbeiter_feld
               (feld_name,feld_typ,optionen,reihenfolge,aktiv,erstellt_am)
               VALUES (?,?,?,?,1,?)""", (fn, ft, opt, i, ts(-300)))
        feld_ids.append(cur.lastrowid)
    standorte = ["Berlin", "Muenchen", "Hamburg"]
    for n, mid in enumerate(ma_ids):
        werte = [f"P-{1000+n}", d(-300 - n * 30),
                 random.choice(["Ja", "Nein"]), random.choice(standorte),
                 f"KST-{100+n}", d(400 + n * 20),
                 str(random.choice([20, 30, 40])),
                 random.choice(["Frueh", "Spaet", "Tag"]),
                 f"0151-3000-{300+n}", random.choice(["Ja", "Nein"])]
        for fid, wert in zip(feld_ids, werte):
            cur.execute(
                """INSERT OR REPLACE INTO tbl_mitarbeiter_wert
                   (mitarbeiter_id,feld_id,wert) VALUES (?,?,?)""",
                (mid, fid, wert))
    return ma_ids


# --------------------------------------------------------------------- Kunden
APO_NAMEN = ["Adler", "Baeren", "Central", "Dom", "Engel", "Forsthaus",
             "Gronau", "Hirsch", "Insel", "Jupiter", "Kranich", "Loewen"]
ORTE = [("10115", "Berlin"), ("80331", "Muenchen"), ("50667", "Koeln"),
        ("20095", "Hamburg"), ("70173", "Stuttgart"), ("01067", "Dresden"),
        ("60311", "Frankfurt"), ("04109", "Leipzig"), ("30159", "Hannover"),
        ("90402", "Nuernberg"), ("28195", "Bremen"), ("40213", "Duesseldorf")]


def seed_kunden(con):
    """kunden_center auf 12 Apotheken erweitern (Demo)."""
    cur = con.cursor()
    cur.execute("DELETE FROM tbl_kunden_center")
    kunden = []
    for i, nm in enumerate(APO_NAMEN):
        knr = str(10001 + i)
        plz, ort = ORTE[i]
        name = f"{nm}-Apotheke"
        cur.execute(
            """INSERT INTO tbl_kunden_center
               (kundennummer,kundenname,kundentyp,email,status,
                erstellt_am,bearbeiter,plz,ort,strasse,inhaber,
                rechnungsart)
               VALUES (?,?,?,?, 'aktiv', ?, 'demo', ?,?,?,?,?)""",
            (knr, name, "PK", f"info@{nm.lower()}-apo.de", ts(-200),
             plz, ort, f"Hauptstrasse {i+1}",
             f"{VORNAMEN[i]} {nm}", "Sofortige Rechnung"))
        kunden.append(dict(knr=knr, name=name, plz=plz, ort=ort,
                           strasse=f"Hauptstrasse {i+1}"))
    return kunden


# ---------------------------------------------------------------------- Kasse
LIEFERANTEN = ["PHOENIX", "Sanacorp", "AEP", "NOWEDA", "Gehe",
               "Alliance Healthcare", "Pharma Privat"]


def seed_kasse(con, pool, kunden):
    cur = con.cursor()
    for t in ("tbl_wareneingang_positionen", "tbl_wareneingang",
              "tbl_lagerbestand", "tbl_kasse_log",
              "tbl_kasse_tagesabschluss", "tbl_kasse_einstellungen"):
        cur.execute(f"DELETE FROM {t}")

    # Wareneingaenge (12) + Positionen
    lager = {}  # (pzn,charge,verfall) -> menge
    for i in range(N):
        cur.execute(
            """INSERT INTO tbl_wareneingang
               (datum,lieferant,lieferschein,bearbeiter,notizen,erstellt_am)
               VALUES (?,?,?,?,?,?)""",
            (d(-90 + i * 5), random.choice(LIEFERANTEN),
             f"LS-2026-{5000+i}", "demo",
             "Demo-Wareneingang", ts(-90 + i * 5)))
        we_id = cur.lastrowid
        for art in random.sample(pool, random.randint(2, 4)):
            menge = random.randint(5, 40)
            charge = f"CH-{random.randint(1000,9999)}"
            verfall = f"{random.randint(1,12):02d}/{random.randint(2027,2030)}"
            cur.execute(
                """INSERT INTO tbl_wareneingang_positionen
                   (we_id,pzn,artikelname,charge,verfall,menge,ek)
                   VALUES (?,?,?,?,?,?,?)""",
                (we_id, art["pzn"], art["artikel"], charge, verfall,
                 menge, art["ek"]))
            key = (art["pzn"], art["artikel"], charge, verfall, art["ek"])
            lager[key] = lager.get(key, 0) + menge

    # Lagerbestand (>=12 aus Wareneingaengen aggregiert)
    for (pzn, name, charge, verfall, ek), menge in list(lager.items())[:max(N, 14)]:
        cur.execute(
            """INSERT OR REPLACE INTO tbl_lagerbestand
               (pzn,artikelname,charge,verfall,menge,aktualisiert_am,ek)
               VALUES (?,?,?,?,?,?,?)""",
            (pzn, name, charge, verfall, menge, ts(-1), ek))

    # Kasse-Log (12)
    aktionen = ["Verkauf erfasst", "Bestellung angelegt", "Storno",
                "Tagesabschluss", "Lagerkorrektur", "Wareneingang gebucht"]
    for i in range(N):
        k = kunden[i % len(kunden)]
        cur.execute(
            """INSERT INTO tbl_kasse_log
               (zeitpunkt,bearbeiter,aktion,bestell_id,kunde,details)
               VALUES (?,?,?,?,?,?)""",
            (ts(-30 + i), "demo", random.choice(aktionen),
             i + 1, k["name"], "Demo-Vorgang"))

    # Tagesabschluss (12 verschiedene Tage)
    for i in range(N):
        anzahl = random.randint(8, 25)
        pakete = anzahl * random.randint(1, 4)
        brutto = round(random.uniform(800, 6000), 2)
        rabatt = round(brutto * random.uniform(0, 0.05), 2)
        cur.execute(
            """INSERT OR REPLACE INTO tbl_kasse_tagesabschluss
               (datum,nr,erzeugt_am,erzeugt_von,anzahl,pakete,brutto,rabatt,netto)
               VALUES (?,?,?,?,?,?,?,?,?)""",
            (d(-i - 1), i + 1, ts(-i - 1), "demo", anzahl, pakete,
             brutto, rabatt, round(brutto - rabatt, 2)))

    # Einstellungen (12 Schluessel)
    settings = {
        "kasse_aktiv": "1", "standard_lieferzeit": "10 Uhr",
        "rabatt_anzeigen": "1", "drucker": "Standard",
        "bon_fusszeile": "Vielen Dank", "mwst_satz": "19",
        "auto_tagesabschluss": "0", "lager_warnschwelle": "5",
        "filiale": "Zentrale", "waehrung": "EUR",
        "beleg_praefix": "K", "demo_modus": "1",
    }
    for k, v in settings.items():
        cur.execute(
            """INSERT OR REPLACE INTO tbl_kasse_einstellungen
               (schluessel,wert,geaendert_am,bearbeiter) VALUES (?,?,?, 'demo')""",
            (k, v, ts(-1)))


# -------------------------------------------------------------------- Faktura
def seed_faktura(con, pool, kunden):
    cur = con.cursor()
    for t in ("tbl_faktura_positionen", "tbl_faktura_belege",
              "tbl_faktura_mitarbeiter"):
        cur.execute(f"DELETE FROM {t}")

    # Faktura-Mitarbeiter (12)
    for i in range(N):
        vn, nn = VORNAMEN[i], NACHNAMEN[i]
        cur.execute(
            """INSERT INTO tbl_faktura_mitarbeiter
               (benutzer,name,email,telefon,aktiv,erstellt_am)
               VALUES (?,?,?,?,1,?)""",
            (f"{vn.lower()}{nn.lower()[0]}", f"{vn} {nn}",
             f"{vn.lower()}.{nn.lower()}@nmg-pharma.de",
             f"030-1000-{100+i}", ts(-200)))

    # Belege (12) + Positionen
    belegarten = ["rechnung"] * 9 + ["gutschrift", "rechnung", "rechnung"]
    for i in range(N):
        k = kunden[i % len(kunden)]
        belegart = belegarten[i]
        praefix = "RE" if belegart == "rechnung" else "GU"
        beleg_nr = f"{praefix}-2026-{1001+i}"
        positionen = random.sample(pool, random.randint(2, 4))
        netto = ust = brutto = 0.0
        cur.execute(
            """INSERT INTO tbl_faktura_belege
               (belegart,beleg_nr,kunde_nr,kunde_name,kunde_adresse,
                beleg_datum,leistungsdatum,netto,ust_betrag,brutto,
                status,mitarbeiter,erstellt_am,festgeschrieben_am)
               VALUES (?,?,?,?,?,?,?,0,0,0, 'festgeschrieben', 'demo', ?, ?)""",
            (belegart, beleg_nr, k["knr"], k["name"],
             f"{k['strasse']}, {k['plz']} {k['ort']}",
             d(-60 + i * 4), d(-60 + i * 4), ts(-60 + i * 4), ts(-60 + i * 4)))
        beleg_id = cur.lastrowid
        for pn, art in enumerate(positionen, start=1):
            menge = random.randint(1, 6)
            rabatt = art["rabatt"]  # echter Rabatt aus NMG/Biosimilar/Stamm
            apu = art["apu"]
            netto_z = round(menge * apu * (1 - rabatt / 100), 2)
            ust_z = round(netto_z * 0.19, 2)
            brutto_z = round(netto_z + ust_z, 2)
            cur.execute(
                """INSERT INTO tbl_faktura_positionen
                   (beleg_id,pos_nr,pzn,bezeichnung,menge,apu_einzel,rabatt,
                    ust_satz,netto_zeile,ust_zeile,brutto_zeile)
                   VALUES (?,?,?,?,?,?,?,19,?,?,?)""",
                (beleg_id, pn, art["pzn"], art["artikel"], menge, apu,
                 rabatt, netto_z, ust_z, brutto_z))
            netto += netto_z
            ust += ust_z
            brutto += brutto_z
        sign = -1 if belegart == "gutschrift" else 1
        cur.execute(
            "UPDATE tbl_faktura_belege SET netto=?,ust_betrag=?,brutto=? WHERE id=?",
            (round(sign * netto, 2), round(sign * ust, 2),
             round(sign * brutto, 2), beleg_id))

    # Nummernkreis-Zaehler nachziehen
    cur.execute("UPDATE tbl_faktura_nummernkreis SET letzter_zaehler=1012 "
                "WHERE belegart='rechnung' AND jahr=2026")
    cur.execute("UPDATE tbl_faktura_nummernkreis SET letzter_zaehler=1001 "
                "WHERE belegart='gutschrift' AND jahr=2026")


# ------------------------------------------------------------------- Bestell.
def seed_bestellungen(con, pool, kunden):
    cur = con.cursor()
    cur.execute("DELETE FROM tbl_bestellpositionen")
    cur.execute("DELETE FROM tbl_bestellungen")
    status_w = ["offen", "offen", "geliefert", "in Bearbeitung", "geliefert"]
    for i in range(N):
        k = kunden[i % len(kunden)]
        cur.execute(
            """INSERT INTO tbl_bestellungen
               (datum,kundennummer,apotheke,bestellart,lieferzeit,status,
                bearbeiter,erstellt_am,msk_erfasst,msk_von,msk_am)
               VALUES (?,?,?, 'Bestellung','10 Uhr',?, 'demo', ?,1,'demo',?)""",
            (d(-40 + i * 3), k["knr"], k["name"], random.choice(status_w),
             ts(-40 + i * 3), ts(-40 + i * 3)))
        b_id = cur.lastrowid
        for art in random.sample(pool, random.randint(2, 4)):
            cur.execute(
                """INSERT INTO tbl_bestellpositionen
                   (bestell_id,pzn,artikelname,df,pck,apu,menge,
                    rabatt_prozent,rabatt_quelle,charge,verfall,bestellart,lieferzeit)
                   VALUES (?,?,?,?,?,?,?,?,?, 'CH-A', ?, 'Bestellung','10 Uhr')""",
                (b_id, art["pzn"], art["artikel"], art["df"], art["pck"],
                 art["apu"], random.randint(1, 8), art["rabatt"],
                 art["rabatt_quelle"], f"{random.randint(1,12):02d}/2028"))


# ------------------------------------------------------------------ Austausch
def seed_austausch(con, pool):
    """Demo-Austauscheintraege Original -> NMG (klar als 'Demo' markiert).

    Die echten Bestandsdaten (quelle != 'Demo') bleiben unberuehrt.
    """
    cur = con.cursor()
    cur.execute("DELETE FROM tbl_austauschdatenbank WHERE quelle='Demo'")
    cur.execute("DELETE FROM tbl_austauschartikel  WHERE quelle='Demo'")

    nmg = [a for a in pool if a["herkunft"] in ("NMG", "Biosimilar")]
    orig = [a for a in pool if a["herkunft"] == "Standard"]
    if len(nmg) < 4 or len(orig) < 4:
        return

    # vorhandene PKs, damit tbl_austauschartikel (PK=original_pzn) nicht kollidiert
    belegt = {r[0] for r in con.execute(
        "SELECT original_pzn FROM tbl_austauschartikel")}

    paare = 0
    for o in orig:
        if paare >= N:
            break
        if o["pzn"] in belegt:
            continue
        ziel = random.choice(nmg)
        # 1) Austauschdatenbank (Freitext-Zuordnung)
        cur.execute(
            """INSERT INTO tbl_austauschdatenbank
               (pzn_alt,artikel_alt,pzn_nmg,artikel_nmg,freitext_austausch,
                quelle,status,erstellt_am,gueltig_ab,bemerkung,bearbeiter)
               VALUES (?,?,?,?,?, 'Demo','aktiv', ?, ?, 'Demo-Austausch','demo')""",
            (o["pzn"], o["artikel"], ziel["pzn"], ziel["artikel"],
             ziel["artikel"], ts(-120 + paare), d(-120 + paare)))
        # 2) Austauschartikel (1:1, PK=original_pzn)
        cur.execute(
            """INSERT INTO tbl_austauschartikel
               (original_pzn,original_artikel,nmg_pzn,austauschbar_gegen,
                quelle,letzte_aktualisierung,treffer_anzahl,bemerkung)
               VALUES (?,?,?,?, 'Demo', ?, ?, 'Demo-Austausch')""",
            (o["pzn"], o["artikel"], ziel["pzn"], ziel["artikel"],
             ts(-120 + paare), random.randint(1, 5)))
        belegt.add(o["pzn"])
        paare += 1


# ----------------------------------------------------------------------- GDP
GDP_DDL = """
CREATE TABLE IF NOT EXISTS tbl_gdp_auslieferung(
    id INTEGER PRIMARY KEY AUTOINCREMENT, datum TEXT, kundennummer TEXT, kunde_name TEXT,
    pzn TEXT, artikelname TEXT, charge TEXT, verfall TEXT, menge INTEGER DEFAULT 0,
    beleg_nr TEXT, bearbeiter TEXT, erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS tbl_gdp_we_pruefung(
    id INTEGER PRIMARY KEY AUTOINCREMENT, we_id INTEGER, datum TEXT, lieferant TEXT,
    transport_temp_c REAL, temp_ok INTEGER DEFAULT 1, unversehrt INTEGER DEFAULT 1,
    dokumente_ok INTEGER DEFAULT 1, gdp_konform INTEGER DEFAULT 1, geprueft_von TEXT,
    bemerkung TEXT, erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS tbl_gdp_messpunkt(
    id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, typ TEXT, soll_min REAL, soll_max REAL,
    aktiv INTEGER DEFAULT 1);
CREATE TABLE IF NOT EXISTS tbl_gdp_temperatur(
    id INTEGER PRIMARY KEY AUTOINCREMENT, messpunkt_id INTEGER, zeitpunkt TEXT, temp_c REAL,
    status TEXT, erfasst_von TEXT, notiz TEXT, massnahme TEXT, behoben INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS tbl_gdp_inspektion(
    id INTEGER PRIMARY KEY AUTOINCREMENT, datum TEXT, titel TEXT, typ TEXT, status TEXT,
    durchgefuehrt_von TEXT, naechste_faellig TEXT, bemerkung TEXT,
    erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS tbl_gdp_inspektion_punkt(
    id INTEGER PRIMARY KEY AUTOINCREMENT, inspektion_id INTEGER, kategorie TEXT, frage TEXT,
    ergebnis TEXT, bemerkung TEXT, massnahme TEXT);
CREATE TABLE IF NOT EXISTS tbl_gdp_kunde_quali(
    id INTEGER PRIMARY KEY AUTOINCREMENT, kundennummer TEXT UNIQUE, kunde_name TEXT,
    lizenznummer TEXT, lizenz_typ TEXT, lizenz_gueltig_bis TEXT, qualifiziert INTEGER DEFAULT 0,
    geprueft_am TEXT, geprueft_von TEXT, bemerkung TEXT);
CREATE TABLE IF NOT EXISTS tbl_gdp_retoure(
    id INTEGER PRIMARY KEY AUTOINCREMENT, datum TEXT, typ TEXT, kundennummer TEXT, kunde_name TEXT,
    pzn TEXT, artikelname TEXT, charge TEXT, verfall TEXT, menge INTEGER DEFAULT 0, grund TEXT,
    temperaturbruch INTEGER DEFAULT 0, status TEXT, entscheidung TEXT, gutschrift_beleg TEXT,
    faktura_beleg_id INTEGER, bearbeiter TEXT, erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP,
    abgeschlossen_am TEXT, notiz TEXT);
CREATE TABLE IF NOT EXISTS tbl_gdp_rueckruf(
    id INTEGER PRIMARY KEY AUTOINCREMENT, datum TEXT, charge TEXT, pzn TEXT, artikelname TEXT,
    grund TEXT, betroffene_kunden INTEGER DEFAULT 0, betroffene_menge INTEGER DEFAULT 0,
    status TEXT, ausgeloest_von TEXT, bemerkung TEXT, erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS tbl_gdp_log(
    id INTEGER PRIMARY KEY AUTOINCREMENT, zeitpunkt TEXT DEFAULT CURRENT_TIMESTAMP, bearbeiter TEXT,
    modul TEXT, aktion TEXT, bezug_id INTEGER, details TEXT);
CREATE TABLE IF NOT EXISTS tbl_gdp_abschreibung(
    id INTEGER PRIMARY KEY AUTOINCREMENT, datum TEXT, pzn TEXT, artikelname TEXT, charge TEXT,
    verfall TEXT, menge INTEGER DEFAULT 0, grund TEXT, wert_ek REAL, retoure_id INTEGER,
    bearbeiter TEXT, erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS tbl_gdp_einstellungen(schluessel TEXT PRIMARY KEY, wert TEXT);
CREATE TABLE IF NOT EXISTS tbl_gdp_produktionsbestand(
    id INTEGER PRIMARY KEY AUTOINCREMENT, pzn TEXT, artikelname TEXT, charge TEXT, verfall TEXT,
    menge INTEGER DEFAULT 0, ek REAL, aktualisiert_am TEXT, UNIQUE(pzn, charge, verfall));
CREATE TABLE IF NOT EXISTS tbl_gdp_warenausgang(
    id INTEGER PRIMARY KEY AUTOINCREMENT, datum TEXT, nummer TEXT, quelle TEXT,
    ziel TEXT DEFAULT 'Verkaufsbestand', status TEXT DEFAULT 'avisiert', bemerkung TEXT,
    rechnungsnummer TEXT, erstellt_von TEXT, erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP,
    bestaetigt_am TEXT, bestaetigt_von TEXT, we_id INTEGER);
CREATE TABLE IF NOT EXISTS tbl_gdp_warenausgang_pos(
    id INTEGER PRIMARY KEY AUTOINCREMENT, wa_id INTEGER, pzn TEXT, artikelname TEXT,
    charge TEXT, verfall TEXT, menge INTEGER DEFAULT 0, ek REAL);
CREATE TABLE IF NOT EXISTS tbl_gdp_bestandsdiff(
    id INTEGER PRIMARY KEY AUTOINCREMENT, datum TEXT, bereich TEXT, pzn TEXT, artikelname TEXT,
    charge TEXT, verfall TEXT, menge_vorher INTEGER, menge_diff INTEGER, menge_nachher INTEGER,
    grund TEXT, bemerkung TEXT, bearbeiter TEXT, erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS tbl_gdp_meldung(
    id INTEGER PRIMARY KEY AUTOINCREMENT, datum TEXT, typ TEXT, titel TEXT, prioritaet TEXT,
    pzn TEXT, charge TEXT, betrifft TEXT, beschreibung TEXT, status TEXT DEFAULT 'Offen',
    massnahme TEXT, verantwortlich TEXT, faellig_am TEXT, erledigt_am TEXT,
    gemeldet_von TEXT, erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP);
"""

INSPEKT_TEMPLATE = [
    ("Raeumlichkeiten", "Lager sauber, trocken, abschliessbar?"),
    ("Raeumlichkeiten", "Wareneingang/-ausgang getrennt?"),
    ("Temperatur", "Kuehlkette luekenlos dokumentiert?"),
    ("Temperatur", "Messgeraete kalibriert?"),
    ("Dokumentation", "Chargen rueckverfolgbar (Kunde<->Charge)?"),
    ("Dokumentation", "Lieferantenqualifizierung aktuell?"),
    ("Kunden", "Nur lizenzierte Apotheken beliefert?"),
    ("Retouren", "Retouren-Verfahren eingehalten?"),
    ("Retouren", "Vernichtung dokumentiert?"),
    ("Personal", "Schulungen GDP aktuell?"),
    ("Faelschungsschutz", "securPharm / Verifizierung aktiv?"),
    ("Selbstinspektion", "Massnahmen aus letzter Inspektion erledigt?"),
]

RET_GRUENDE = ["Falschlieferung", "Ablauf/Verfall nahe", "Transportschaden",
               "Temperaturbruch", "Ueberbestellung", "Qualitaetsmangel",
               "Bestellfehler Apotheke", "Sonstiges"]


def seed_gdp(con, pool, kunden):
    """GDP-Modul: Auslieferungen (Kunde<->Charge), Temperatur, Inspektionen,
    Kundenqualifizierung, Retouren, Rueckrufe, Protokoll. Mind. 12 je Schritt."""
    cur = con.cursor()
    con.executescript(GDP_DDL)
    for t in ("tbl_gdp_auslieferung", "tbl_gdp_we_pruefung", "tbl_gdp_messpunkt",
              "tbl_gdp_temperatur", "tbl_gdp_inspektion", "tbl_gdp_inspektion_punkt",
              "tbl_gdp_kunde_quali", "tbl_gdp_retoure", "tbl_gdp_rueckruf", "tbl_gdp_log",
              "tbl_gdp_abschreibung", "tbl_gdp_meldung", "tbl_gdp_produktionsbestand",
              "tbl_gdp_warenausgang", "tbl_gdp_warenausgang_pos", "tbl_gdp_bestandsdiff"):
        cur.execute(f"DELETE FROM {t}")
    # Retourenbestand-Spalte am Lager sicherstellen + zuruecksetzen
    lcols = {r[1] for r in cur.execute("PRAGMA table_info(tbl_lagerbestand)")}
    if "menge_retoure" not in lcols:
        cur.execute("ALTER TABLE tbl_lagerbestand ADD COLUMN menge_retoure INTEGER DEFAULT 0")
    else:
        cur.execute("UPDATE tbl_lagerbestand SET menge_retoure=0")
    # Wareneingangs-Art sicherstellen; Demo zeigt beide Wareneingaenge aktiv
    wcols = {r[1] for r in cur.execute("PRAGMA table_info(tbl_wareneingang)")}
    if "art" not in wcols:
        cur.execute("ALTER TABLE tbl_wareneingang ADD COLUMN art TEXT DEFAULT 'bestand'")
    else:
        cur.execute("UPDATE tbl_wareneingang SET art='bestand'")
    # Demo startet im Verkauf-Modus (zeigt Retouren/Avis-Bestaetigung); zum
    # Produktionsteil in den Einstellungen auf 'produktion' umschalten.
    cur.execute("INSERT OR REPLACE INTO tbl_gdp_einstellungen(schluessel,wert) "
                "VALUES('betriebsmodus','verkauf')")

    # Reale Chargen aus dem Wareneingang ziehen (Basis fuer Rueckverfolgung)
    charges = cur.execute(
        "SELECT pzn, artikelname, charge, verfall FROM tbl_wareneingang_positionen "
        "WHERE charge IS NOT NULL").fetchall()
    if not charges:
        charges = [(a["pzn"], a["artikel"], f"CH-{1000+i}", "06/2028")
                   for i, a in enumerate(pool[:12])]

    # GDP-Wareneingangspruefungen zu vorhandenen Wareneingaengen
    we_ids = [r[0] for r in cur.execute(
        "SELECT id, lieferant, datum FROM tbl_wareneingang ORDER BY id").fetchall()]
    for n, wid in enumerate(we_ids):
        # Die letzten 3 Eingaenge bewusst OHNE GDP-Pruefung lassen (History-Filter
        # "Offen" demonstrierbar + offene Pflicht auf der Uebersicht).
        if n >= len(we_ids) - 3:
            continue
        temp = round(random.uniform(2.5, 7.8), 1)
        konform = 0 if n % 7 == 0 else 1
        cur.execute(
            """INSERT INTO tbl_gdp_we_pruefung
               (we_id,datum,lieferant,transport_temp_c,temp_ok,unversehrt,dokumente_ok,
                gdp_konform,geprueft_von,bemerkung)
               VALUES (?,?,?,?,?,?,?,?, 'demo', ?)""",
            (wid, d(-90 + n * 5), random.choice(LIEFERANTEN), temp, 1, konform, 1, konform,
             "Demo-Pruefung" if konform else "Karton beschaedigt - Klaerung"))

    # Zweiter Wareneingang (Einkauf -> Produktion): die ersten 3 Eingaenge als
    # 'produktion' markieren und ihre Ware in den getrennten Produktionsbestand
    # buchen (NICHT in den verkaufbaren Lagerbestand).
    prod_we = we_ids[:3]
    for wid in prod_we:
        cur.execute("UPDATE tbl_wareneingang SET art='produktion', lieferant=? WHERE id=?",
                    (random.choice(["EuroPharma S.L.", "Medis d.o.o.", "PharmaParallel BV"]), wid))
        for pzn, name, charge, verfall, menge, ek in cur.execute(
                "SELECT pzn,artikelname,charge,verfall,menge,ek FROM tbl_wareneingang_positionen "
                "WHERE we_id=?", (wid,)).fetchall():
            cur.execute(
                """INSERT INTO tbl_gdp_produktionsbestand(pzn,artikelname,charge,verfall,menge,ek,aktualisiert_am)
                   VALUES (?,?,?,?,?,?,?)
                   ON CONFLICT(pzn,charge,verfall) DO UPDATE SET menge=menge+excluded.menge""",
                (pzn, name, charge, verfall, menge, ek or round(random.uniform(60, 400), 2), ts(-1)))

    # Warenausgaenge aus der Produktion (Avis an den Verkaufs-Wareneingang):
    # 3 avisiert (warten auf Bestaetigung) + 1 bereits bestaetigt.
    for j in range(4):
        pzn, name, charge, verfall = charges[(j + 1) % len(charges)]
        status = "bestaetigt" if j == 3 else "avisiert"
        nummer = f"WA-2026-{1001 + j}"
        # Demo: jeder zweite Warenausgang mit Rechnungsnummer (Feld ist optional)
        rnr = f"RE-2026-{4200 + j}" if j % 2 == 0 else ""
        cur.execute(
            """INSERT INTO tbl_gdp_warenausgang
               (datum,nummer,quelle,ziel,status,bemerkung,rechnungsnummer,erstellt_von,bestaetigt_am,bestaetigt_von)
               VALUES (?,?, 'Produktion', 'Verkaufsbestand', ?, 'Demo-Produktionscharge', ?, 'demo', ?, ?)""",
            (ts(-6 + j), nummer, status, rnr,
             ts(-2) if status == "bestaetigt" else None,
             "demo" if status == "bestaetigt" else None))
        wa_id = cur.lastrowid
        for _ in range(random.randint(1, 3)):
            p2, n2, c2, v2 = charges[random.randrange(len(charges))]
            cur.execute(
                "INSERT INTO tbl_gdp_warenausgang_pos(wa_id,pzn,artikelname,charge,verfall,menge,ek) "
                "VALUES (?,?,?,?,?,?,?)",
                (wa_id, p2, n2, f"FW-{random.randint(1000,9999)}", v2,
                 random.randint(5, 30), round(random.uniform(80, 500), 2)))

    # Manuelle Bestandsdifferenzen (Demo): Verkaufs- und Produktionsbestand
    bd_gruende = ["Inventur", "Bruch / Beschaedigung", "Schwund / Diebstahl",
                  "Fund / Mehrbestand", "Buchungsfehler"]
    lager_lines = cur.execute(
        "SELECT pzn,artikelname,charge,verfall,menge FROM tbl_lagerbestand WHERE menge>0 LIMIT 6").fetchall()
    prod_lines = cur.execute(
        "SELECT pzn,artikelname,charge,verfall,menge FROM tbl_gdp_produktionsbestand WHERE menge>0 LIMIT 4").fetchall()
    # Datums ueber mehrere Zeitraeume streuen, damit Monat/Quartal/Jahr/Vorjahr
    # im Protokoll-Zeitraumfilter unterschiedliche Ergebnisse zeigen.
    bd_offsets = [-2, -8, -20, -55, -120, -200, -300, -400, -430, -460]
    for n, (bereich, line) in enumerate(
            [("Verkauf", l) for l in lager_lines] + [("Produktion", l) for l in prod_lines]):
        pzn, name, charge, verfall, vorher = line
        diff = random.choice([-3, -2, -1, 1, 2])
        nachher = max(0, (vorher or 0) + diff)
        echt_diff = nachher - (vorher or 0)
        off = bd_offsets[n % len(bd_offsets)]
        grund = bd_gruende[n % len(bd_gruende)]
        cur.execute(
            """INSERT INTO tbl_gdp_bestandsdiff
               (datum,bereich,pzn,artikelname,charge,verfall,menge_vorher,menge_diff,menge_nachher,
                grund,bemerkung,bearbeiter)
               VALUES (?,?,?,?,?,?,?,?,?,?, 'Demo-Korrektur', 'demo')""",
            (ts(off), bereich, pzn, name, charge, verfall, vorher, echt_diff, nachher, grund))
        # passender Protokoll-Eintrag (damit Modul=Bestandsdifferenz + Zeitraum greift)
        cur.execute(
            "INSERT INTO tbl_gdp_log(zeitpunkt,bearbeiter,modul,aktion,bezug_id,details) "
            "VALUES (?, 'demo', 'Bestandsdifferenz', ?, ?, ?)",
            (ts(off), f"{bereich}: {'+' if echt_diff > 0 else ''}{echt_diff}", n + 1,
             f"{name} Ch {charge}: {vorher} -> {nachher} ({grund})"))

    # Auslieferungen Kunde <-> Charge (>= 24)
    beleg = 7000
    for i, (pzn, name, charge, verfall) in enumerate(charges):
        for k in random.sample(kunden, random.randint(1, 3)):
            beleg += 1
            cur.execute(
                """INSERT INTO tbl_gdp_auslieferung
                   (datum,kundennummer,kunde_name,pzn,artikelname,charge,verfall,menge,beleg_nr,bearbeiter)
                   VALUES (?,?,?,?,?,?,?,?,?, 'demo')""",
                (d(-70 + i), k["knr"], k["name"], pzn, name, charge, verfall,
                 random.randint(1, 12), f"LF-2026-{beleg}"))

    # Messpunkte (4) + Temperatur-Messungen (>=12, einige Abweichungen)
    messpunkte = [("Kuehlschrank 1 (2-8 C)", "Kuehlschrank", 2.0, 8.0),
                  ("Kuehlschrank 2 (2-8 C)", "Kuehlschrank", 2.0, 8.0),
                  ("Lager Trockenbereich", "Lager", 15.0, 25.0),
                  ("Transportbox Kuehlkette", "Transport", 2.0, 8.0)]
    mp_ids = []
    for name, typ, smin, smax in messpunkte:
        cur.execute("INSERT INTO tbl_gdp_messpunkt(name,typ,soll_min,soll_max,aktiv) VALUES (?,?,?,?,1)",
                    (name, typ, smin, smax))
        mp_ids.append((cur.lastrowid, smin, smax))
    for i in range(16):
        mid, smin, smax = random.choice(mp_ids)
        if i % 6 == 0:  # bewusste Abweichung
            temp = round(smax + random.uniform(0.5, 3.0), 1)
            status, behoben = "Abweichung", (1 if i % 12 else 0)
            massn = "Kuehlaggregat geprueft, Ware umgelagert" if behoben else None
        else:
            temp = round(random.uniform(smin + 0.3, smax - 0.3), 1)
            status, behoben, massn = "ok", 0, None
        cur.execute(
            """INSERT INTO tbl_gdp_temperatur
               (messpunkt_id,zeitpunkt,temp_c,status,erfasst_von,notiz,massnahme,behoben)
               VALUES (?,?,?,?, 'demo', ?,?,?)""",
            (mid, ts(-20 + i) , temp, status, "Routinemessung", massn, behoben))

    # Selbstinspektionen (2) + Pruefpunkte (>=12 je)
    for j, (titel, status, faellig_off) in enumerate([
            ("GDP-Selbstinspektion 2025", "abgeschlossen", 200),
            ("GDP-Selbstinspektion 2026", "laeuft", 365)]):
        cur.execute(
            """INSERT INTO tbl_gdp_inspektion
               (datum,titel,typ,status,durchgefuehrt_von,naechste_faellig,bemerkung)
               VALUES (?,?, 'Selbstinspektion', ?, 'demo', ?, 'Demo-Inspektion')""",
            (d(-200 + j * 200), titel, status, d(faellig_off - 200)))
        iid = cur.lastrowid
        for kat, frage in INSPEKT_TEMPLATE:
            if status == "abgeschlossen":
                erg = "Abweichung" if random.random() < 0.15 else "ok"
            else:
                erg = random.choice(["ok", "ok", "offen", "Abweichung"])
            massn = "Korrekturmassnahme eingeleitet" if erg == "Abweichung" else None
            cur.execute(
                "INSERT INTO tbl_gdp_inspektion_punkt(inspektion_id,kategorie,frage,ergebnis,massnahme) "
                "VALUES (?,?,?,?,?)", (iid, kat, frage, erg, massn))

    # Kundenqualifizierung (alle Kunden; einige abgelaufen/gesperrt)
    typen = ["Apothekenbetriebserlaubnis", "Apothekenbetriebserlaubnis",
             "Grosshandelserlaubnis (83 AMG)", "Krankenhausapotheke"]
    for i, k in enumerate(kunden):
        if i % 6 == 5:
            gueltig, quali = d(-30), 1          # abgelaufen
        elif i % 6 == 4:
            gueltig, quali = d(400), 0          # gesperrt (nicht freigegeben)
        else:
            gueltig, quali = d(300 + i * 10), 1  # ok
        cur.execute(
            """INSERT INTO tbl_gdp_kunde_quali
               (kundennummer,kunde_name,lizenznummer,lizenz_typ,lizenz_gueltig_bis,
                qualifiziert,geprueft_am,geprueft_von,bemerkung)
               VALUES (?,?,?,?,?,?,?, 'demo', ?)""",
            (k["knr"], k["name"], f"ABE-{2000+i}", typen[i % len(typen)], gueltig,
             quali, d(-100 + i), "Demo-Qualifizierung"))

    # Retouren / Reklamationen (>=12, gemischte Status)
    status_flow = [("Neu", None, None), ("In Pruefung", None, None),
                   ("Gutschrift", "Gutschrift erteilt", "GU-2026-{}"),
                   ("Im Retourenbestand", "In Quarantaene", None),
                   ("Vernichtet", "Vernichtung", None), ("Abgelehnt", "Abgelehnt", None)]
    quarantaene = {}  # (pzn,charge,verfall) -> Stueck im Retourenbestand
    for i in range(14):
        pzn, name, charge, verfall = charges[i % len(charges)]
        k = kunden[i % len(kunden)]
        status, entsch, gut_t = status_flow[i % len(status_flow)]
        grund = RET_GRUENDE[i % len(RET_GRUENDE)]
        tbruch = 1 if grund == "Temperaturbruch" else 0
        gut = gut_t.format(1100 + i) if gut_t else None
        menge = random.randint(1, 8)
        abg = d(-10 + i) if status in ("Gutschrift", "Im Retourenbestand", "Vernichtet", "Abgelehnt") else None
        cur.execute(
            """INSERT INTO tbl_gdp_retoure
               (datum,typ,kundennummer,kunde_name,pzn,artikelname,charge,verfall,menge,grund,
                temperaturbruch,status,entscheidung,gutschrift_beleg,bearbeiter,abgeschlossen_am,notiz)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?, 'demo', ?, ?)""",
            (d(-25 + i), "Reklamation" if i % 4 == 0 else "Retoure", k["knr"], k["name"],
             pzn, name, charge, verfall, menge, grund, tbruch, status,
             entsch, gut, abg, "Demo-Vorgang"))
        if status == "Im Retourenbestand":
            quarantaene[(pzn, charge, verfall, name)] = quarantaene.get((pzn, charge, verfall, name), 0) + menge

    # Retourenbestand (Quarantaene) ins Lager schreiben - gesperrt, nicht verkaufbar
    for (pzn, charge, verfall, name), stk in quarantaene.items():
        ex = cur.execute(
            "SELECT id, ek FROM tbl_lagerbestand WHERE pzn=? AND COALESCE(charge,'')=? "
            "AND COALESCE(verfall,'')=?", (pzn, charge or "", verfall or "")).fetchone()
        if ex:
            cur.execute("UPDATE tbl_lagerbestand SET menge_retoure=COALESCE(menge_retoure,0)+? WHERE id=?",
                        (stk, ex[0]))
        else:
            cur.execute(
                "INSERT INTO tbl_lagerbestand(pzn,artikelname,charge,verfall,menge,menge_retoure,aktualisiert_am,ek) "
                "VALUES (?,?,?,?,0,?,?,?)", (pzn, name, charge, verfall, stk, ts(-1),
                                            round(random.uniform(80, 600), 2)))

    # Abschreibungen / Write-offs (Demo, >=4)
    for i in range(5):
        pzn, name, charge, verfall = charges[(i + 2) % len(charges)]
        menge = random.randint(1, 5)
        ek = round(random.uniform(80, 600), 2)
        cur.execute(
            """INSERT INTO tbl_gdp_abschreibung
               (datum,pzn,artikelname,charge,verfall,menge,grund,wert_ek,bearbeiter)
               VALUES (?,?,?,?,?,?,?,?, 'demo')""",
            (d(-8 + i), pzn, name, charge, verfall, menge,
             ["Beschaedigt", "Verfall ueberschritten", "Temperaturbruch", "Nicht GDP-konform", "Rueckruf"][i],
             round(ek * menge, 2)))

    # Rueckrufe (2)
    for i in range(2):
        pzn, name, charge, verfall = charges[i]
        n_k = cur.execute("SELECT COUNT(DISTINCT kundennummer) FROM tbl_gdp_auslieferung WHERE charge=?",
                          (charge,)).fetchone()[0]
        summe = cur.execute("SELECT COALESCE(SUM(menge),0) FROM tbl_gdp_auslieferung WHERE charge=?",
                            (charge,)).fetchone()[0]
        cur.execute(
            """INSERT INTO tbl_gdp_rueckruf
               (datum,charge,pzn,artikelname,grund,betroffene_kunden,betroffene_menge,status,ausgeloest_von,bemerkung)
               VALUES (?,?,?,?,?,?,?, 'ausgeloest', 'demo', 'Demo-Rueckruf')""",
            (d(-12 + i), charge, pzn, name,
             "Qualitaetsmangel" if i == 0 else "Temperaturabweichung", n_k, summe))

    # Protokoll (>=12)
    log_aktionen = [
        ("Wareneingang", "GDP-Pruefung erfasst"), ("Retoure", "Retoure angelegt"),
        ("Retoure", "Gutschrift erstellt"), ("Kuehlkette", "Messung Abweichung"),
        ("Kuehlkette", "Massnahme dokumentiert"), ("Rueckruf", "Rueckruf ausgeloest"),
        ("Qualifizierung", "Lizenz/Qualifizierung aktualisiert"),
        ("Inspektion", "Inspektion abgeschlossen"), ("Retoure", "Wiedereingelagert"),
        ("Retoure", "Vernichtung dokumentiert"), ("Wareneingang", "GDP-Pruefung erfasst"),
        ("Rueckruf", "Verteiler exportiert"),
    ]
    for i, (modul, aktion) in enumerate(log_aktionen):
        cur.execute(
            "INSERT INTO tbl_gdp_log(zeitpunkt,bearbeiter,modul,aktion,bezug_id,details) "
            "VALUES (?, 'demo', ?,?,?, 'Demo-Eintrag')",
            (ts(-15 + i), modul, aktion, i + 1))

    # Meldungen (>=12, gemischte Typen/Prioritaeten/Status; einige offen/ueberfaellig)
    meld_vorlagen = [
        ("Temperaturbruch", "Kuehlschrank 1 ueber 8 C", "Hoch", "Kuehlschrank 1 (2-8 C)",
         "Temperatur kurzzeitig auf 11 C gestiegen, Ware geprueft."),
        ("Qualitaetsmangel", "Beschaedigte Umverpackung", "Mittel", None,
         "Mehrere Faltschachteln eingedrueckt geliefert."),
        ("Abweichung", "Lieferschein unvollstaendig", "Niedrig", None,
         "Chargenangabe auf dem Lieferschein fehlte."),
        ("Reklamation", "Apotheke meldet Falschlieferung", "Mittel", None,
         "Falsche PZN geliefert, Austausch veranlasst."),
        ("Faelschungsverdacht / securPharm", "securPharm-Alarm beim Verifizieren", "Kritisch", None,
         "Packung liess sich nicht verifizieren - gesperrt und gemeldet."),
        ("Transportschaden", "Kuehlbox undicht angeliefert", "Hoch", None,
         "Kuehlkette nicht sicher - Ware in Quarantaene."),
        ("Lieferengpass", "Wirkstoff voraussichtlich nicht lieferbar", "Mittel", None,
         "Alternative Bezugsquelle wird geprueft."),
        ("Abweichung", "Messgeraet ueberfaellig kalibriert", "Mittel", None,
         "Kalibrierung des Thermometers terminieren."),
    ]
    status_zyklus = ["Offen", "In Bearbeitung", "Erledigt", "Offen", "Verworfen", "In Bearbeitung"]
    n = 0
    for runde in range(2):
        for j, (typ, titel, prio, betrifft, beschr) in enumerate(meld_vorlagen):
            status = status_zyklus[n % len(status_zyklus)]
            ch = charges[n % len(charges)]
            betr = betrifft or kunden[n % len(kunden)]["name"]
            # ein paar bewusst ueberfaellige offene Meldungen
            faellig = d(-3 + n) if (status in ("Offen", "In Bearbeitung") and n % 3 == 0) else d(10 + n)
            erledigt = d(-2 + n) if status in ("Erledigt", "Verworfen") else None
            massn = ("Korrektur eingeleitet, Wirksamkeit geprueft."
                     if status in ("In Bearbeitung", "Erledigt") else None)
            cur.execute(
                """INSERT INTO tbl_gdp_meldung
                   (datum,typ,titel,prioritaet,pzn,charge,betrifft,beschreibung,status,
                    massnahme,verantwortlich,faellig_am,erledigt_am,gemeldet_von)
                   VALUES (?,?,?,?,?,?,?,?,?,?, 'demo', ?,?, 'demo')""",
                (d(-30 + n), typ, titel, prio, ch[0], ch[2], betr, beschr, status,
                 massn, faellig, erledigt))
            cur.execute(
                "INSERT INTO tbl_gdp_log(zeitpunkt,bearbeiter,modul,aktion,bezug_id,details) "
                "VALUES (?, 'demo', 'Meldung', 'Meldung angelegt', ?, ?)",
                (ts(-30 + n), cur.lastrowid, f"{typ}: {titel}"))
            n += 1


def main():
    if not os.path.exists(SRC):
        raise SystemExit(f"Quell-DB fehlt: {SRC}")
    shutil.copy2(SRC, DST)
    print(f"Kopie angelegt: {DST}")

    con = sqlite3.connect(DST)
    con.execute("PRAGMA foreign_keys=ON")
    try:
        pool = artikel_pool(con)
        if len(pool) < 10:
            raise SystemExit("Zu wenige Artikel im Stamm fuer Demodaten.")
        from collections import Counter
        mix = Counter(a["herkunft"] for a in pool)
        print(f"Artikelpool: {len(pool)} echte Artikel "
              f"(NMG={mix['NMG']}, Biosimilar={mix['Biosimilar']}, "
              f"Standard={mix['Standard']})")

        seed_personal(con)
        kunden = seed_kunden(con)
        seed_kasse(con, pool, kunden)
        seed_faktura(con, pool, kunden)
        seed_bestellungen(con, pool, kunden)
        seed_austausch(con, pool)
        seed_gdp(con, pool, kunden)
        con.commit()
    finally:
        # Report
        tabs = ["tbl_mitarbeiter", "tbl_arbeitsbereich", "tbl_abwesenheit",
                "tbl_mitarbeiter_arbeitsbereich", "tbl_mitarbeiter_vorgesetzter",
                "tbl_mitarbeiter_feld", "tbl_mitarbeiter_wert",
                "tbl_kunden_center", "tbl_wareneingang",
                "tbl_wareneingang_positionen", "tbl_lagerbestand",
                "tbl_kasse_log", "tbl_kasse_tagesabschluss",
                "tbl_kasse_einstellungen", "tbl_faktura_mitarbeiter",
                "tbl_faktura_belege", "tbl_faktura_positionen",
                "tbl_bestellungen", "tbl_bestellpositionen",
                "tbl_gdp_auslieferung", "tbl_gdp_we_pruefung", "tbl_gdp_temperatur",
                "tbl_gdp_inspektion", "tbl_gdp_inspektion_punkt", "tbl_gdp_kunde_quali",
                "tbl_gdp_retoure", "tbl_gdp_rueckruf", "tbl_gdp_log",
                "tbl_gdp_abschreibung", "tbl_gdp_produktionsbestand",
                "tbl_gdp_warenausgang", "tbl_gdp_warenausgang_pos", "tbl_gdp_bestandsdiff"]
        print("\n=== Befuellte Tabellen ===")
        for t in tabs:
            n = con.execute(f"SELECT count(*) FROM {t}").fetchone()[0]
            flag = "OK " if n >= 10 else "!! "
            print(f"  {flag}{t:38s} {n}")
        # Austausch: nur Demo-Zeilen zaehlen, Echtdaten bleiben unberuehrt
        for t in ("tbl_austauschdatenbank", "tbl_austauschartikel"):
            dn = con.execute(
                f"SELECT count(*) FROM {t} WHERE quelle='Demo'").fetchone()[0]
            ges = con.execute(f"SELECT count(*) FROM {t}").fetchone()[0]
            flag = "OK " if dn >= 10 else "!! "
            print(f"  {flag}{t:38s} {dn} Demo  ({ges} gesamt, Rest unberuehrt)")
        con.close()


if __name__ == "__main__":
    main()
