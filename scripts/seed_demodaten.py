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
        pool.append(dict(
            pzn=pzn, artikel=(name or "").strip(), df=df, pck=pck,
            apu=apu, ek=ek, rabatt=rab,
            rabatt_quelle=(rquelle or "NMG-Kondition"), herkunft="NMG"))
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
                "tbl_bestellungen", "tbl_bestellpositionen"]
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
